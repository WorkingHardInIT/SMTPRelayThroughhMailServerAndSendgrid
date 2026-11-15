#Requires -Version 5.1
<#
.SYNOPSIS
Installs and reconciles an hMailServer SendGrid relay.
.DESCRIPTION
The JSON config is the complete desired allowlist, not incremental additions.
The default HMailServerRelay.json is read beside this script, independent of the
working directory. Its supplied device addresses are examples; verify them first.
Devices, domains, sender rewrite address, SMTP identity, and upstream host/port
come from JSON. smtpHostname is the relay's FQDN (for example
smtp01.datawisetech.corp), not the SendGrid host. Create its DNS record separately.
Credentials are prompted for and must not be stored in the config. Missing,
invalid, or unknown settings stop execution before any changes. An explicit
empty devices array removes managed device entries but retains loopback access.
Reruns add, update, and remove script-managed hMailServer ranges, including when
Windows Firewall is skipped. Names starting with 'SendGrid relay ' are reserved
for this script. Older unprefixed example ranges (and currently configured device
names) are migrated only when their settings match this script's relay settings.
Unrelated ranges and domains are preserved. Previously removed custom unprefixed
ranges cannot be identified automatically and must be reviewed manually once.
The three named firewall rules are updated in place. -SkipWindowsFirewall leaves
existing firewall rules unchanged; -KeepInternetRange leaves Internet unchanged.
SMTP access on 'My computer' is disabled; the managed loopback relay is retained.
Use -WhatIf to preview without prompting for credentials or changing anything.
When installation or firewall changes require elevation, the script requests UAC
and relaunches in Windows PowerShell, preserving parameters and relative paths.
The original invocation returns after launch; finish in the elevated window.
Installation checks include the service, interop DLL, and COM registration.
An incomplete standard installation is repaired using the verified installer;
the existing administrator password is required when repairing an installation.
.PARAMETER ConfigPath
Path to the plain JSON configuration. Defaults to HMailServerRelay.json beside
this script. Device names must be unique and CIDRs must be individual IPv4 /32s.
.EXAMPLE
.\Configure-HMailServerRelay.ps1 -WhatIf
.EXAMPLE
.\Configure-HMailServerRelay.ps1 -ConfigPath 'C:\SysAdmin\Relay.json'
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidateNotNullOrEmpty()]
    [string]$ConfigPath,

    [string]$HMailAdminUser = 'Administrator',

    # This installer is pinned to the latest stable release verified when this script was updated.
    [string]$InstallerPath = 'C:\hMailServerFiles\Installer\hMailServer-5.7.1-B3116-x64.exe',

    [switch]$KeepInternetRange,

    [switch]$SkipWindowsFirewall
)

if (-not $PSBoundParameters.ContainsKey('ConfigPath')) {
    $ConfigPath = Join-Path $PSScriptRoot 'HMailServerRelay.json'
}

$useUnicodeIcons = $PSVersionTable.PSVersion.Major -ge 6 -or
    -not [string]::IsNullOrWhiteSpace($env:WT_SESSION)
if ($useUnicodeIcons) {
    $script:StatusIcons = @{
        Added = [char]::ConvertFromUtf32(0x2795)
        Existing = [char]::ConvertFromUtf32(0x2705)
        Updated = [char]::ConvertFromUtf32(0x1F504)
        Missing = [char]::ConvertFromUtf32(0x2139)
        Warning = [char]::ConvertFromUtf32(0x26A0)
        Error = [char]::ConvertFromUtf32(0x274C)
    }
}
else {
    $script:StatusIcons = @{ Added = '[+]'; Existing = '[=]'; Updated = '[~]'; Missing = '[i]'; Warning = '[!]'; Error = '[X]' }
}
$script:StatusColors = @{ Added = 'Green'; Existing = 'DarkGreen'; Updated = 'Cyan'; Missing = 'DarkGray'; Warning = 'Yellow'; Error = 'Red' }

function Write-Status {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Added', 'Existing', 'Updated', 'Missing', 'Warning', 'Error')]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host ("{0} {1}" -f $script:StatusIcons[$Status], $Message) -ForegroundColor $script:StatusColors[$Status]
}

function Assert-HMailConfigObject {
    param($Value, [string]$Setting, [string[]]$Required, [string[]]$Optional = @())

    if ($Value -isnot [pscustomobject]) {
        throw "Config '$Setting' must be a JSON object."
    }
    $names = @($Value.PSObject.Properties.Name)
    foreach ($name in $Required) {
        if ($names -cnotcontains $name) {
            throw "Config '$Setting' is missing required property '$name'."
        }
    }
    foreach ($name in $names) {
        if (($Required + $Optional) -cnotcontains $name) {
            throw "Unknown config property '$Setting.$name'. Check spelling; credentials do not belong in this file."
        }
    }
}

function Assert-HMailConfigString {
    param($Value, [string]$Setting)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value) -or $Value -cne $Value.Trim()) {
        throw "Config '$Setting' must be a nonempty string without leading or trailing whitespace."
    }
}

function Read-HMailRelayConfig {
    param([Parameter(Mandatory = $true)][string]$Path)

    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
    try {
        $config = ConvertFrom-Json -InputObject $json -ErrorAction Stop
    }
    catch {
        throw "Could not parse relay JSON config '$Path': $($_.Exception.Message)"
    }
    if ($json -notmatch '^\s*\{') {
        throw 'Relay config must be a JSON object.'
    }
    Assert-HMailConfigObject $config 'root' @('devices', 'internalDomain', 'authorizedDomain', 'authorizedFrom', 'sendGrid', 'smtpHostname', 'smtpAuthentication', 'tls')
    Assert-HMailConfigString $config.smtpHostname 'smtpHostname'
    if ([Uri]::CheckHostName($config.smtpHostname) -ne [UriHostNameType]::Dns -or
        $config.smtpHostname -notmatch '\.' -or $config.smtpHostname.Length -gt 253 -or
        $config.smtpHostname -match '(^|\.)[^a-zA-Z0-9]|[^a-zA-Z0-9](\.|$)' -or
        $config.smtpHostname -match '[^a-zA-Z0-9.-]') {
        throw "Config 'smtpHostname' must be a fully qualified DNS hostname such as smtp01.datawisetech.corp."
    }
    foreach ($name in 'internalDomain', 'authorizedDomain', 'authorizedFrom') {
        Assert-HMailConfigString $config.$name $name
    }
    foreach ($name in 'internalDomain', 'authorizedDomain') {
        if ([Uri]::CheckHostName($config.$name) -ne [UriHostNameType]::Dns) {
            throw "Config '$name' must be a DNS domain name."
        }
    }
    if ($config.authorizedFrom -notmatch ('^[^@\s]+@' + [regex]::Escape($config.authorizedDomain) + '$')) {
        throw "Config 'authorizedFrom' must use the authorized domain '$($config.authorizedDomain)'."
    }
    Assert-HMailConfigObject $config.sendGrid 'sendGrid' @('host', 'port')
    Assert-HMailConfigString $config.sendGrid.host 'sendGrid.host'
    if ([Uri]::CheckHostName($config.sendGrid.host) -eq [UriHostNameType]::Unknown) {
        throw "Config 'sendGrid.host' must be a hostname or IP address, not a URL."
    }
    $port = $config.sendGrid.port
    if (($port -isnot [int] -and $port -isnot [long]) -or $port -lt 1 -or $port -gt 65535) {
        throw "Config 'sendGrid.port' must be an integer from 1 to 65535."
    }
    Assert-HMailConfigObject $config.tls 'tls' @('preferServerCipherOrder', 'prioritizeChaCha20')
    foreach ($name in 'preferServerCipherOrder', 'prioritizeChaCha20') {
        if ($config.tls.$name -isnot [bool]) {
            throw "Config 'tls.$name' must be a boolean."
        }
    }
    if ($config.tls.prioritizeChaCha20 -and -not $config.tls.preferServerCipherOrder) {
        throw "Config 'tls.prioritizeChaCha20' requires 'tls.preferServerCipherOrder' to be true."
    }
    Assert-HMailConfigObject $config.smtpAuthentication 'smtpAuthentication' @(
        'enabled', 'generateDemoCertificate', 'accountAddress', 'port', 'certificateName', 'certificateFile', 'privateKeyFile'
    )
    if ($config.smtpAuthentication.enabled -isnot [bool]) {
        throw "Config 'smtpAuthentication.enabled' must be a boolean."
    }
    if ($config.smtpAuthentication.generateDemoCertificate -isnot [bool]) {
        throw "Config 'smtpAuthentication.generateDemoCertificate' must be a boolean."
    }
    if ($config.smtpAuthentication.generateDemoCertificate -and -not $config.smtpAuthentication.enabled) {
        throw "Config 'smtpAuthentication.generateDemoCertificate' requires 'smtpAuthentication.enabled' to be true."
    }
    Assert-HMailConfigString $config.smtpAuthentication.accountAddress 'smtpAuthentication.accountAddress'
    Assert-HMailConfigString $config.smtpAuthentication.certificateName 'smtpAuthentication.certificateName'
    if ($config.smtpAuthentication.accountAddress -notmatch ('^[^@\s]+@' + [regex]::Escape($config.authorizedDomain) + '$')) {
        throw "Config 'smtpAuthentication.accountAddress' must use the authorized domain '$($config.authorizedDomain)'."
    }
    if (($config.smtpAuthentication.port -isnot [int] -and $config.smtpAuthentication.port -isnot [long]) -or
        $config.smtpAuthentication.port -ne 587) {
        throw "Config 'smtpAuthentication.port' must be 587."
    }
    foreach ($pathProperty in 'certificateFile', 'privateKeyFile') {
        if ($config.smtpAuthentication.$pathProperty -isnot [string]) {
            throw "Config 'smtpAuthentication.$pathProperty' must be a string."
        }
        if ($config.smtpAuthentication.enabled -and
            [string]::IsNullOrWhiteSpace($config.smtpAuthentication.$pathProperty)) {
            throw "Config 'smtpAuthentication.$pathProperty' is required when SMTP authentication is enabled."
        }
    }
    if ($config.devices -isnot [array]) {
        throw "Config 'devices' must be an explicit JSON array (use [] for no devices)."
    }
    $names = @{ '127.0.0.1' = $true }
    $addresses = @{ '127.0.0.1/32' = $true }
    foreach ($device in $config.devices) {
        Assert-HMailConfigObject $device 'devices[]' @('name', 'cidr') @('description', 'requireSmtpAuth')
        Assert-HMailConfigString $device.name 'devices[].name'
        Assert-HMailConfigString $device.cidr 'devices[].cidr'
        if ($device.PSObject.Properties.Name -contains 'description' -and $device.description -isnot [string]) {
            throw "Config 'devices[].description' must be a string when specified."
        }
        if ($device.PSObject.Properties.Name -contains 'requireSmtpAuth' -and $device.requireSmtpAuth -isnot [bool]) {
            throw "Config 'devices[].requireSmtpAuth' must be a boolean when specified."
        }
        if ($device.cidr -notmatch '^(?:[0-9]{1,3}\.){3}[0-9]{1,3}/32$') {
            throw "Device '$($device.name)' must use an individual dotted IPv4 address with /32."
        }
        $bounds = ConvertFrom-IPv4Cidr -CIDR $device.cidr
        if ($device.cidr -cne "$($bounds.LowerIP)/32") {
            throw "Device '$($device.name)' must use a canonical IPv4 /32 address."
        }
        if ($names.ContainsKey($device.name)) {
            throw "Duplicate or reserved device name '$($device.name)'. Names must be unique; loopback is automatic."
        }
        if ($addresses.ContainsKey($device.cidr)) {
            throw "Duplicate device address '$($device.cidr)'. Addresses must be unique; loopback is automatic."
        }
        $names[$device.name] = $true
        $addresses[$device.cidr] = $true
    }
    return $config
}

function Get-HMailAccountByAddressOrNull {
    param($Accounts, [string]$Address)

    for ($index = 0; $index -lt $Accounts.Count; $index++) {
        $account = $Accounts.Item($index)
        if ($account.Address -ieq $Address) {
            return $account
        }
    }
    return $null
}

function Set-HMailSMTPAuthAccount {
    param($Domain, [string]$Address, [string]$Password)

    $account = Get-HMailAccountByAddressOrNull -Accounts $Domain.Accounts -Address $Address
    if (-not $account) {
        $account = $Domain.Accounts.Add()
        $account.Address = $Address
        $status = 'Added'
    }
    else {
        $status = 'Updated'
    }
    $account.Password = $Password
    $account.Active = $true
    $account.Save()
    Write-Status -Status $status -Message "Provisioned SMTP authentication account '$Address'."
}

function Set-HMailSSLCertificate {
    param(
        $Certificates,
        [string]$Name,
        [string]$CertificateFile,
        [string]$PrivateKeyFile
    )

    $certificate = Get-ItemByNameOrNull -Collection $Certificates -Name $Name
    if (-not $certificate) {
        $certificate = $Certificates.Add()
        $status = 'Added'
    }
    else {
        $status = 'Updated'
    }
    $changed = $false
    foreach ($property in 'Name', 'CertificateFile', 'PrivateKeyFile') {
        $value = switch ($property) {
            'Name' { $Name }
            'CertificateFile' { $CertificateFile }
            'PrivateKeyFile' { $PrivateKeyFile }
        }
        if ($certificate.$property -cne $value) {
            $certificate.$property = $value
            $changed = $true
        }
    }
    if ($changed) {
        $certificate.Save()
    }
    else {
        $status = 'Existing'
    }
    Write-Status -Status $status -Message "hMailServer TLS certificate '$Name' is configured."
    return $certificate
}

function Set-HMailTLSSettings {
    param(
        $Settings,
        [bool]$PreferServerCipherOrder,
        [bool]$PrioritizeChaCha20
    )

    $desiredSettings = [ordered]@{
        TlsVersion10Enabled = $false
        TlsVersion11Enabled = $false
        TlsVersion12Enabled = $true
        TlsVersion13Enabled = $true
        TlsOptionPreferServerCiphersEnabled = $PreferServerCipherOrder
        TlsOptionPrioritizeChaChaEnabled = $PrioritizeChaCha20
    }
    $changed = $false
    foreach ($property in $desiredSettings.Keys) {
        if ($Settings.$property -ne $desiredSettings[$property]) {
            $Settings.$property = $desiredSettings[$property]
            $changed = $true
        }
    }
    Write-Status -Status $(if ($changed) { 'Updated' } else { 'Existing' }) `
        -Message "TLS policy: TLS 1.0/1.1 disabled; TLS 1.2/1.3 enabled; server cipher order: $PreferServerCipherOrder; prioritize ChaCha20: $PrioritizeChaCha20."
    return $changed
}

function Set-HMailSMTPSubmissionPort {
    param(
        $Ports,
        [int]$PortNumber,
        [int]$CertificateID,
        $SMTPProtocol,
        $StartTlsRequired
    )

    $port = $null
    for ($index = 0; $index -lt $Ports.Count; $index++) {
        $candidate = $Ports.Item($index)
        if ($candidate.PortNumber -ne $PortNumber) {
            continue
        }
        if ($candidate.Address -eq '0.0.0.0' -and $candidate.Protocol -eq $SMTPProtocol) {
            if ($port) {
                throw "Multiple hMailServer TCP/IP listeners conflict with the configured SMTP port $PortNumber."
            }
            $port = $candidate
        }
        else {
            throw "hMailServer TCP/IP port $PortNumber is already assigned to '$($candidate.Address)' and another protocol or binding."
        }
    }
    if (-not $port) {
        $port = $Ports.Add()
    }

    $changed = $false
    $properties = @{
        Address = '0.0.0.0'
        PortNumber = $PortNumber
        Protocol = $SMTPProtocol
        ConnectionSecurity = $StartTlsRequired
        SSLCertificateID = $CertificateID
    }
    foreach ($property in $properties.Keys) {
        if ($port.$property -ne $properties[$property]) {
            $port.$property = $properties[$property]
            $changed = $true
        }
    }
    if ($changed) {
        $port.Save()
    }
    Write-Status -Status $(if ($changed) { 'Updated' } else { 'Existing' }) `
        -Message "hMailServer SMTP listener on port $PortNumber requires STARTTLS."
    return $changed
}

function Start-HMailElevatedSession {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][hashtable]$Parameters
    )

    $forwarded = @{}
    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]
        if ($value -is [System.Management.Automation.SwitchParameter]) {
            $value = $value.IsPresent
        }
        $forwarded[$key] = $value
    }
    $payload = [System.Management.Automation.PSSerializer]::Serialize(@{
        ScriptPath = $ScriptPath
        Parameters = $forwarded
        WorkingDirectory = (Get-Location).ProviderPath
    })
    $encodedPayload = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payload))
    $command = @"
`$ErrorActionPreference = 'Stop'
`$invocation = [System.Management.Automation.PSSerializer]::Deserialize([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('$encodedPayload')))
Set-Location -LiteralPath `$invocation.WorkingDirectory
`$parameters = `$invocation.Parameters
& `$invocation.ScriptPath @parameters
"@
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -Verb RunAs -ArgumentList @('-NoProfile', '-NoExit', '-EncodedCommand', $encodedCommand) `
        -ErrorAction Stop | Out-Null
    Write-Status -Status Missing -Message 'Launched elevated PowerShell. Complete confirmation and credential prompts in that window.'
}

function ConvertTo-PlainText {
    param([Parameter(Mandatory = $true)][Security.SecureString]$SecureString)

    $credential = New-Object System.Management.Automation.PSCredential ('temporary', $SecureString)
    return $credential.GetNetworkCredential().Password
}

function Get-HMailInteropPath {
    foreach ($programFiles in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not [string]::IsNullOrWhiteSpace($programFiles)) {
            $candidate = Join-Path $programFiles 'hMailServer\Bin\Interop.hMailServer.dll'
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }
    }
    return $null
}

function Get-HMailInstaller {
    param([Parameter(Mandatory = $true)][string]$Path)

    $installerUrl = 'https://github.com/hmailserver/hmailserver/releases/download/hMailServer-5.7.1-b3116/hMailServer-5.7.1-B3116-x64.exe'
    $expectedHash = 'e86a7d4a6a8f44227b730b1e2f2173ae191a753d79b32d3633e55fa83f10ae21'
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $parent = Split-Path -Path $Path -Parent
        if ([string]::IsNullOrWhiteSpace($parent)) {
            $parent = (Get-Location).Path
        }

        if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
            New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop | Out-Null
        }
        $temporaryPath = Join-Path $parent ('hMailServer-download-' + [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            try {
                Invoke-WebRequest -UseBasicParsing -Uri $installerUrl -OutFile $temporaryPath -ErrorAction Stop
            }
            catch {
                throw "Could not download hMailServer 5.7.1 Build 3116: $($_.Exception.Message). Download the official x64 installer from '$installerUrl', save it as '$Path', and rerun the script. The installer will be checked against its SHA-256 checksum before use."
            }
            if ((Get-FileHash -LiteralPath $temporaryPath -Algorithm SHA256 -ErrorAction Stop).Hash -ne $expectedHash) {
                throw 'The downloaded hMailServer installer failed SHA-256 verification.'
            }
            Move-Item -LiteralPath $temporaryPath -Destination $Path -ErrorAction Stop
            Write-Status -Status Added -Message "Downloaded hMailServer 5.7.1 Build 3116 to '$Path'."
        }
        finally {
            if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
                Remove-Item -LiteralPath $temporaryPath -ErrorAction Stop
            }
        }
    }

    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash -ne $expectedHash) {
        throw "Installer '$Path' does not match the official hMailServer 5.7.1 Build 3116 SHA-256 checksum. It was not run or replaced."
    }
    return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
}

function ConvertTo-DerLength {
    param([Parameter(Mandatory = $true)][int]$Length)

    if ($Length -lt 128) {
        return ,([byte[]]@([byte]$Length))
    }

    $lengthBytes = New-Object 'System.Collections.Generic.List[byte]'
    $remaining = $Length
    while ($remaining -gt 0) {
        $lengthBytes.Insert(0, [byte]($remaining -band 0xff))
        $remaining = [int][Math]::Floor($remaining / 256)
    }
    $encodedLength = New-Object 'System.Collections.Generic.List[byte]'
    $encodedLength.Add([byte](0x80 -bor $lengthBytes.Count))
    $encodedLength.AddRange($lengthBytes)
    return ,([byte[]]$encodedLength.ToArray())
}

function ConvertTo-DerElement {
    param(
        [Parameter(Mandatory = $true)][byte]$Tag,
        [Parameter(Mandatory = $true)][byte[]]$Content
    )

    $encoded = New-Object 'System.Collections.Generic.List[byte]'
    $encoded.Add($Tag)
    $encoded.AddRange([byte[]](ConvertTo-DerLength -Length $Content.Length))
    $encoded.AddRange($Content)
    return ,([byte[]]$encoded.ToArray())
}

function ConvertTo-DerInteger {
    param([Parameter(Mandatory = $true)][byte[]]$Value)

    $firstNonzero = 0
    while ($firstNonzero -lt ($Value.Length - 1) -and $Value[$firstNonzero] -eq 0) {
        $firstNonzero++
    }
    $integer = New-Object byte[] ($Value.Length - $firstNonzero)
    [Array]::Copy($Value, $firstNonzero, $integer, 0, $integer.Length)
    if (($integer[0] -band 0x80) -ne 0) {
        $positiveInteger = New-Object byte[] ($integer.Length + 1)
        [Array]::Copy($integer, 0, $positiveInteger, 1, $integer.Length)
        $integer = $positiveInteger
    }
    return ConvertTo-DerElement -Tag 0x02 -Content $integer
}

function ConvertTo-Pkcs1PrivateKey {
    param([Parameter(Mandatory = $true)][System.Security.Cryptography.RSAParameters]$Parameters)

    $components = New-Object 'System.Collections.Generic.List[byte[]]'
    $components.Add([byte[]]@(0))
    foreach ($component in @(
        $Parameters.Modulus,
        $Parameters.Exponent,
        $Parameters.D,
        $Parameters.P,
        $Parameters.Q,
        $Parameters.DP,
        $Parameters.DQ,
        $Parameters.InverseQ
    )) {
        if (-not $component -or $component.Length -eq 0) {
            throw 'The Windows certificate provider did not export a complete RSA private key.'
        }
        $components.Add([byte[]]$component)
    }

    $sequenceContent = New-Object 'System.Collections.Generic.List[byte]'
    foreach ($component in $components) {
        $sequenceContent.AddRange([byte[]](ConvertTo-DerInteger -Value $component))
    }
    return ConvertTo-DerElement -Tag 0x30 -Content ([byte[]]$sequenceContent.ToArray())
}

function ConvertTo-Pem {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][byte[]]$Bytes
    )

    $base64 = [Convert]::ToBase64String($Bytes)
    $lines = [regex]::Matches($base64, '.{1,64}') | ForEach-Object { $_.Value }
    return "-----BEGIN $Label-----`r`n$($lines -join "`r`n")`r`n-----END $Label-----`r`n"
}

function Set-DemoCertificatePrivateKeyAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$ServiceAccount
    )

    $acl = New-Object System.Security.AccessControl.FileSecurity
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetOwner([System.Security.Principal.NTAccount]::new('BUILTIN\Administrators'))
    $identities = @('NT AUTHORITY\SYSTEM', 'BUILTIN\Administrators')
    if (-not [string]::IsNullOrWhiteSpace($ServiceAccount)) {
        switch ($ServiceAccount.ToLowerInvariant()) {
            { $_ -in 'localsystem', 'nt authority\system' } { $ServiceAccount = 'NT AUTHORITY\SYSTEM'; break }
            'nt authority\localservice' { $ServiceAccount = 'NT AUTHORITY\LOCAL SERVICE'; break }
            'nt authority\networkservice' { $ServiceAccount = 'NT AUTHORITY\NETWORK SERVICE'; break }
        }
        if ($identities -notcontains $ServiceAccount) {
            $identities += $ServiceAccount
        }
    }

    foreach ($identity in $identities) {
        $rights = if ($identity -eq 'BUILTIN\Administrators') {
            [System.Security.AccessControl.FileSystemRights]::FullControl
        }
        else {
            [System.Security.AccessControl.FileSystemRights]::Read
        }
        $rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList @(
            $identity,
            $rights,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
}

function New-HMailDemoCertificate {
    param(
        [Parameter(Mandatory = $true)][string]$HostName,
        [Parameter(Mandatory = $true)][string]$CertificateFile,
        [Parameter(Mandatory = $true)][string]$PrivateKeyFile
    )

    $certificateExists = Test-Path -LiteralPath $CertificateFile -PathType Leaf
    $privateKeyExists = Test-Path -LiteralPath $PrivateKeyFile -PathType Leaf
    if ($certificateExists -and $privateKeyExists) {
        Write-Status -Status Existing -Message "Reusing the existing demo TLS certificate and private key for '$HostName'; neither file was replaced."
        return
    }
    if ($certificateExists -or $privateKeyExists) {
        throw 'Only one demo TLS file exists. Refusing to overwrite or pair a partial certificate/key; review both configured paths.'
    }
    if (-not (Get-Command New-SelfSignedCertificate -ErrorAction SilentlyContinue)) {
        throw 'The native New-SelfSignedCertificate cmdlet is unavailable. Use a supported Windows PowerShell installation; no external certificate tool will be downloaded or installed.'
    }

    foreach ($path in @($CertificateFile, $PrivateKeyFile)) {
        $parent = Split-Path -Path $path -Parent
        if ([string]::IsNullOrWhiteSpace($parent)) {
            throw "TLS output path '$path' must include a parent directory."
        }
        if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
            New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop | Out-Null
        }
    }

    $certificateTemporary = "$CertificateFile.$([guid]::NewGuid().ToString('N')).tmp"
    $privateKeyTemporary = "$PrivateKeyFile.$([guid]::NewGuid().ToString('N')).tmp"
    $certificate = $null
    $rsa = $null
    $certificateMoved = $false
    $privateKeyMoved = $false
    try {
        $certificate = New-SelfSignedCertificate -DnsName $HostName -CertStoreLocation 'Cert:\CurrentUser\My' `
            -NotAfter (Get-Date).AddDays(365) -KeyAlgorithm RSA -KeyLength 3072 -KeyExportPolicy Exportable `
            -HashAlgorithm SHA256 -KeyUsage DigitalSignature, KeyEncipherment -Type SSLServerAuthentication `
            -ErrorAction Stop
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
        if (-not $rsa) {
            throw 'Windows did not provide an RSA private key for the generated certificate.'
        }

        $privateKeyDer = ConvertTo-Pkcs1PrivateKey -Parameters $rsa.ExportParameters($true)
        [System.IO.File]::WriteAllText(
            $certificateTemporary,
            (ConvertTo-Pem -Label 'CERTIFICATE' -Bytes $certificate.RawData),
            [System.Text.Encoding]::ASCII
        )
        [System.IO.File]::WriteAllText(
            $privateKeyTemporary,
            (ConvertTo-Pem -Label 'RSA PRIVATE KEY' -Bytes $privateKeyDer),
            [System.Text.Encoding]::ASCII
        )

        $serviceAccount = $null
        $service = Get-CimInstance -ClassName Win32_Service -Filter "Name='hMailServer'" -ErrorAction SilentlyContinue
        if ($service) {
            $serviceAccount = $service.StartName
        }
        Set-DemoCertificatePrivateKeyAcl -Path $privateKeyTemporary -ServiceAccount $serviceAccount
        Move-Item -LiteralPath $certificateTemporary -Destination $CertificateFile -ErrorAction Stop
        $certificateMoved = $true
        Move-Item -LiteralPath $privateKeyTemporary -Destination $PrivateKeyFile -ErrorAction Stop
        $privateKeyMoved = $true
        Set-DemoCertificatePrivateKeyAcl -Path $PrivateKeyFile -ServiceAccount $serviceAccount
        Write-Status -Status Added -Message "Generated a native PowerShell 3072-bit, one-year self-signed demo certificate for '$HostName' in '$CertificateFile'."
        Write-Status -Status Warning -Message "The private key is unencrypted for hMailServer and restricted to SYSTEM, local Administrators, and the hMailServer service identity when known: '$PrivateKeyFile'."
    }
    catch {
        if ($certificateMoved -and (Test-Path -LiteralPath $CertificateFile -PathType Leaf)) {
            Remove-Item -LiteralPath $CertificateFile -ErrorAction Stop
        }
        if ($privateKeyMoved -and (Test-Path -LiteralPath $PrivateKeyFile -PathType Leaf)) {
            Remove-Item -LiteralPath $PrivateKeyFile -ErrorAction Stop
        }
        throw
    }
    finally {
        if ($rsa) {
            $rsa.Dispose()
        }
        if ($certificate) {
            Remove-Item -LiteralPath "Cert:\CurrentUser\My\$($certificate.Thumbprint)" -Force -ErrorAction Stop
        }
        foreach ($temporaryPath in @($certificateTemporary, $privateKeyTemporary)) {
            if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
                Remove-Item -LiteralPath $temporaryPath -ErrorAction Stop
            }
        }
    }
}

function Get-HMailInstallationState {
    $interopPath = Get-HMailInteropPath
    $service = Get-Service -ErrorAction Stop | Where-Object { $_.Name -eq 'hMailServer' } | Select-Object -First 1
    $comRegistered = Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\hMailServer.Application\CLSID' -ErrorAction Stop
    return [pscustomobject]@{
        InteropPath = $interopPath
        Service = $service
        ComRegistered = $comRegistered
        NeedsSetup = -not ($interopPath -and $service -and $comRegistered)
    }
}

function ConvertTo-WindowsCommandLineArgument {
    param([Parameter(Mandatory = $true)][string]$Value)

    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Install-HMailServer {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][Security.SecureString]$AdminPassword
    )

    $installer = Get-HMailInstaller -Path $Path
    Write-Status -Status Warning -Message 'The installer requires the administrator password on its command line; other local users may be able to see it while setup runs.'
    try {
        $plainTextPassword = ConvertTo-PlainText -SecureString $AdminPassword
        $passwordArgument = ConvertTo-WindowsCommandLineArgument -Value "/ADMINPASSWORD=$plainTextPassword"
        $process = Start-Process -FilePath $installer -ArgumentList @(
            '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/useinternaldbms=true', $passwordArgument
        ) -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -ne 0) {
            throw "hMailServer installer failed with exit code $($process.ExitCode). Review setup and correct the problem before rerunning."
        }
    }
    finally {
        $passwordArgument = $null
        $plainTextPassword = $null
    }
    Write-Status -Status Added -Message 'hMailServer setup completed with the built-in database.'
}

function Get-HMailEnumValue {
    param(
        [Parameter(Mandatory = $true)]$Assembly,
        [Parameter(Mandatory = $true)][string]$TypeName,
        [Parameter(Mandatory = $true)][string]$MemberPattern
    )

    $enumType = $Assembly.GetTypes() | Where-Object { $_.IsEnum -and $_.Name -eq $TypeName } | Select-Object -First 1
    if (-not $enumType) {
        throw "Could not find hMailServer enum '$TypeName' in Interop.hMailServer.dll."
    }

    $members = @([Enum]::GetNames($enumType) | Where-Object { $_ -match $MemberPattern })
    if ($members.Count -ne 1) {
        throw "Could not uniquely resolve '$MemberPattern' in enum '$TypeName'. Available values: $([Enum]::GetNames($enumType) -join ', ')"
    }

    return [Enum]::Parse($enumType, [string]$members[0])
}

function Get-ItemByNameOrNull {
    param($Collection, [string]$Name)

    for ($index = 0; $index -lt $Collection.Count; $index++) {
        $item = $Collection.Item($index)
        if ($item.Name -eq $Name) {
            return $item
        }
    }
    return $null
}

function Set-HMailDomain {
    param($Domains, [string]$Name)

    $domain = Get-ItemByNameOrNull -Collection $Domains -Name $Name
    if (-not $domain) {
        $domain = $Domains.Add()
        $domain.Name = $Name
        $domain.Active = $true
        $domain.Save()
        Write-Status -Status Added -Message "Created hMailServer domain '$Name'."
    }
    else {
        Write-Status -Status Existing -Message "hMailServer domain '$Name' already exists; left unchanged."
    }
}

function ConvertFrom-IPv4Cidr {
    param([Parameter(Mandatory = $true)][string]$CIDR)

    $parts = $CIDR -split '/', 2
    $address = $null
    $prefixLength = 0
    if ($parts.Count -ne 2 -or
        -not [System.Net.IPAddress]::TryParse($parts[0], [ref]$address) -or
        $address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork -or
        -not [int]::TryParse($parts[1], [ref]$prefixLength) -or
        $prefixLength -lt 0 -or $prefixLength -gt 32) {
        throw "Invalid IPv4 CIDR range: $CIDR"
    }

    $networkBytes = $address.GetAddressBytes()
    $broadcastBytes = $address.GetAddressBytes()
    $bitsRemaining = $prefixLength
    for ($index = 0; $index -lt 4; $index++) {
        if ($bitsRemaining -ge 8) {
            $mask = 255
            $bitsRemaining -= 8
        }
        elseif ($bitsRemaining -gt 0) {
            $mask = 256 - [int][Math]::Pow(2, 8 - $bitsRemaining)
            $bitsRemaining = 0
        }
        else {
            $mask = 0
        }

        $networkBytes[$index] = [byte]([int]$networkBytes[$index] -band $mask)
        $broadcastBytes[$index] = [byte]([int]$broadcastBytes[$index] -bor (255 -bxor $mask))
    }

    return @{
        LowerIP = ([System.Net.IPAddress]::new($networkBytes)).ToString()
        UpperIP = ([System.Net.IPAddress]::new($broadcastBytes)).ToString()
    }
}

function Test-IsRFC1918Address {
    param([Parameter(Mandatory = $true)][string]$IPAddress)

    $bytes = ([System.Net.IPAddress]::Parse($IPAddress)).GetAddressBytes()
    return ($bytes[0] -eq 10) -or
        ($bytes[0] -eq 172 -and $bytes[1] -ge 16 -and $bytes[1] -le 31) -or
        ($bytes[0] -eq 192 -and $bytes[1] -eq 168)
}

function ConvertTo-IPv4Integer {
    param([Parameter(Mandatory = $true)][string]$IPAddress)

    $bytes = ([System.Net.IPAddress]::Parse($IPAddress)).GetAddressBytes()
    if ($bytes.Length -ne 4) {
        throw "Expected an IPv4 address: $IPAddress"
    }

    return ([uint64]$bytes[0] * [uint64]16777216) +
        ([uint64]$bytes[1] * [uint64]65536) +
        ([uint64]$bytes[2] * [uint64]256) +
        [uint64]$bytes[3]
}

function ConvertFrom-IPv4Integer {
    param([Parameter(Mandatory = $true)][uint64]$Value)

    $octets = [byte[]]@(
        [int][Math]::Floor($Value / 16777216)
        [int][Math]::Floor(($Value % 16777216) / 65536)
        [int][Math]::Floor(($Value % 65536) / 256)
        [int]($Value % 256)
    )
    return ([System.Net.IPAddress]::new($octets)).ToString()
}

function Get-IPv4CidrBlocksForInterval {
    param(
        [Parameter(Mandatory = $true)][uint64]$Start,
        [Parameter(Mandatory = $true)][uint64]$End
    )

    $blocks = @()
    $current = $Start
    while ($current -le $End) {
        $selectedPrefix = 32
        $selectedSize = [uint64]1
        for ($prefix = 0; $prefix -le 32; $prefix++) {
            $blockSize = [uint64][Math]::Pow(2, 32 - $prefix)
            if (($current % $blockSize) -eq 0 -and ($current + $blockSize - 1) -le $End) {
                $selectedPrefix = $prefix
                $selectedSize = $blockSize
                break
            }
        }

        $blocks += '{0}/{1}' -f (ConvertFrom-IPv4Integer -Value $current), $selectedPrefix
        $current += $selectedSize
    }

    return $blocks
}

function Get-UntrustedIPv4Cidrs {
    param([Parameter(Mandatory = $true)]$AllowedRanges)

    $intervals = @(
        foreach ($allowedRange in $AllowedRanges) {
            [pscustomobject]@{
                Start = ConvertTo-IPv4Integer -IPAddress $allowedRange.LowerIP
                End = ConvertTo-IPv4Integer -IPAddress $allowedRange.UpperIP
            }
        }
    ) | Sort-Object Start

    $merged = New-Object System.Collections.Generic.List[object]
    foreach ($interval in $intervals) {
        if ($merged.Count -eq 0) {
            $merged.Add(@{ Start = $interval.Start; End = $interval.End })
            continue
        }

        $last = $merged[$merged.Count - 1]
        if ($interval.Start -le ($last.End + 1)) {
            if ($interval.End -gt $last.End) {
                $last.End = $interval.End
            }
        }
        else {
            $merged.Add(@{ Start = $interval.Start; End = $interval.End })
        }
    }

    $untrustedCidrs = @()
    $cursor = [uint64]0
    $maximumIPv4 = [uint64]4294967295
    foreach ($interval in $merged) {
        if ($cursor -lt $interval.Start) {
            $untrustedCidrs += Get-IPv4CidrBlocksForInterval -Start $cursor -End ($interval.Start - 1)
        }

        if ($interval.End -ge $maximumIPv4) {
            $cursor = $maximumIPv4 + 1
            break
        }
        $cursor = $interval.End + 1
    }

    if ($cursor -le $maximumIPv4) {
        $untrustedCidrs += Get-IPv4CidrBlocksForInterval -Start $cursor -End $maximumIPv4
    }

    return $untrustedCidrs
}

function Set-SmtpFirewallRules {
    param(
        [Parameter(Mandatory = $true)][string[]]$UnauthenticatedRemoteAddresses,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$AuthenticatedRemoteAddresses,
        [Parameter(Mandatory = $true)][string[]]$BlockedSMTP25IPv4Cidrs,
        [Parameter(Mandatory = $true)][string[]]$BlockedSMTP587IPv4Cidrs,
        [Parameter(Mandatory = $true)][bool]$AuthenticatedSubmissionEnabled
    )

    $rules = @(
        @{ Name = 'HMailServer-SMTP25-Untrusted-IPv4-Block'; DisplayName = 'hMailServer SMTP relay - block untrusted IPv4'; Action = 'Block'; Addresses = $BlockedSMTP25IPv4Cidrs; LocalPort = 25 },
        # Windows Firewall rejects ::/0 on some versions; these two prefixes cover all IPv6.
        @{ Name = 'HMailServer-SMTP25-Untrusted-IPv6-Block'; DisplayName = 'hMailServer SMTP relay - block IPv6'; Action = 'Block'; Addresses = @('::/1', '8000::/1'); LocalPort = 25 },
        @{ Name = 'HMailServer-SMTP25-Trusted-Allow'; DisplayName = 'hMailServer SMTP relay - unauthenticated sources'; Action = 'Allow'; Addresses = $UnauthenticatedRemoteAddresses; LocalPort = 25 }
    )
    $blockedSubmissionAddresses = if ($AuthenticatedSubmissionEnabled) {
        $BlockedSMTP587IPv4Cidrs
    }
    else {
        @('0.0.0.0/0')
    }
    $rules += @(
        @{ Name = 'HMailServer-SMTP587-Untrusted-IPv4-Block'; DisplayName = 'hMailServer SMTP submission - block untrusted IPv4'; Action = 'Block'; Addresses = $blockedSubmissionAddresses; LocalPort = 587 },
        @{ Name = 'HMailServer-SMTP587-Untrusted-IPv6-Block'; DisplayName = 'hMailServer SMTP submission - block IPv6'; Action = 'Block'; Addresses = @('::/1', '8000::/1'); LocalPort = 587 }
    )
    if ($AuthenticatedSubmissionEnabled) {
        $rules += @{ Name = 'HMailServer-SMTP587-Trusted-Allow'; DisplayName = 'hMailServer SMTP submission - authenticated sources'; Action = 'Allow'; Addresses = $AuthenticatedRemoteAddresses; LocalPort = 587 }
    }

    $existingRules = @(Get-NetFirewallRule -PolicyStore PersistentStore -ErrorAction Stop)
    $desiredNames = @($rules | ForEach-Object { $_.Name })
    $managedNames = @(
        'HMailServer-SMTP25-Untrusted-IPv4-Block',
        'HMailServer-SMTP25-Untrusted-IPv6-Block',
        'HMailServer-SMTP25-Trusted-Allow',
        'HMailServer-SMTP587-Untrusted-IPv4-Block',
        'HMailServer-SMTP587-Untrusted-IPv6-Block',
        'HMailServer-SMTP587-Trusted-Allow'
    )
    foreach ($stale in @($existingRules | Where-Object {
        $managedNames -contains $_.Name -and $desiredNames -notcontains $_.Name
    })) {
        $stale | Remove-NetFirewallRule -ErrorAction Stop
        Write-Status -Status Updated -Message "Removed obsolete managed firewall rule '$($stale.Name)'."
    }

    foreach ($rule in $rules) {
        $existing = @($existingRules | Where-Object { $_.Name -eq $rule.Name })
        if ($rule.Addresses.Count -eq 0) {
            $existing | Remove-NetFirewallRule -ErrorAction Stop
            Write-Status -Status Missing -Message "No addresses remain for firewall rule '$($rule.Name)'; removed it if present."
            continue
        }
        $properties = @{
            Direction = 'Inbound'
            Action = $rule.Action
            Protocol = 'TCP'
            LocalPort = $rule.LocalPort
            RemoteAddress = $rule.Addresses
            Profile = 'Any'
            Enabled = 'True'
            ErrorAction = 'Stop'
        }
        if ($existing.Count -gt 0) {
            $existing | Set-NetFirewallRule -NewDisplayName $rule.DisplayName @properties
            $status = 'Updated'
        }
        else {
            New-NetFirewallRule -PolicyStore PersistentStore -Name $rule.Name -DisplayName $rule.DisplayName @properties | Out-Null
            $status = 'Added'
        }
        Write-Status -Status $status -Message "Firewall rule '$($rule.Name)' now applies to the configured addresses."
    }
}

function Get-HMailRelayRangeSettings {
    param([bool]$RequireSMTPAuthLocalToExternal = $false)

    return @{
        Priority = 1000
        AllowSMTPConnections = $true
        AllowPOP3Connections = $false
        AllowIMAPConnections = $false
        AllowDeliveryFromLocalToLocal = $false
        AllowDeliveryFromLocalToRemote = $true
        AllowDeliveryFromRemoteToLocal = $false
        AllowDeliveryFromRemoteToRemote = $false
        RequireSMTPAuthLocalToLocal = $true
        RequireSMTPAuthLocalToExternal = $RequireSMTPAuthLocalToExternal
        RequireSMTPAuthExternalToLocal = $true
        RequireSMTPAuthExternalToExternal = $true
    }
}

function Set-HMailIPRange {
    param(
        $Ranges,
        [string]$Name,
        [string]$LowerIP,
        [string]$UpperIP,
        [bool]$RequireSMTPAuthLocalToExternal = $false
    )

    $range = Get-ItemByNameOrNull -Collection $Ranges -Name $name
    if (-not $range) {
        $range = $Ranges.Add()
        $status = 'Added'
    }
    else {
        $status = 'Updated'
    }

    $properties = Get-HMailRelayRangeSettings -RequireSMTPAuthLocalToExternal $RequireSMTPAuthLocalToExternal
    $properties.Name = $Name
    $properties.LowerIP = $LowerIP
    $properties.UpperIP = $UpperIP
    $changed = $false
    foreach ($property in $properties.Keys) {
        if ($range.$property -ne $properties[$property]) {
            $range.$property = $properties[$property]
            $changed = $true
        }
    }
    if ($changed) {
        $range.Save()
    }
    else {
        $status = 'Existing'
    }
    $authText = if ($RequireSMTPAuthLocalToExternal) { 'with SMTP authentication' } else { 'without SMTP authentication' }
    Write-Status -Status $status -Message "hMailServer IP range '$Name' permits local-to-external relay only $authText ($LowerIP-$UpperIP)."
}

function Sync-HMailIPRanges {
    param($Ranges, [object[]]$DesiredRanges)

    $desiredNames = @{}
    $legacyNames = @{}
    $desiredByLegacyName = @{}
    foreach ($desired in $DesiredRanges) {
        $desiredNames[$desired.Name] = $true
        $legacyNames[$desired.LegacyName] = $true
        $desiredByLegacyName[$desired.LegacyName] = $desired
    }
    $obsoleteRanges = @(
        for ($index = 0; $index -lt $Ranges.Count; $index++) {
            $range = $Ranges.Item($index)
            $managed = $range.Name.StartsWith('SendGrid relay ', [StringComparison]::OrdinalIgnoreCase)
            if (-not $managed -and ($legacyNames.ContainsKey($range.Name) -or
                $range.Name -match '^(DataSwitch|StorageSwitch|OutOfBandSwitch|KempLoadmaster|Firewall)-[1-9][0-9]*$')) {
                $managed = $range.LowerIP -eq $range.UpperIP
                $requireSMTPAuth = $false
                if ($desiredByLegacyName.ContainsKey($range.Name)) {
                    $requireSMTPAuth = [bool]$desiredByLegacyName[$range.Name].RequireSMTPAuthLocalToExternal
                }
                $properties = Get-HMailRelayRangeSettings -RequireSMTPAuthLocalToExternal $requireSMTPAuth
                foreach ($property in $properties.Keys) {
                    if ($range.$property -ne $properties[$property]) {
                        $managed = $false
                    }
                }
                if (-not $managed) {
                    Write-Status -Status Warning -Message "Unprefixed IP range '$($range.Name)' does not match script-managed settings and was left unchanged; review it manually."
                }
            }
            if ($managed -and -not $desiredNames.ContainsKey($range.Name)) {
                [pscustomobject]@{ ID = $range.ID; Name = $range.Name }
            }
        }
    )
    foreach ($desired in $DesiredRanges) {
        Set-HMailIPRange -Ranges $Ranges -Name $desired.Name -LowerIP $desired.LowerIP -UpperIP $desired.UpperIP `
            -RequireSMTPAuthLocalToExternal ([bool]$desired.RequireSMTPAuthLocalToExternal)
    }
    # Prune only after every desired entry has been saved successfully.
    foreach ($obsolete in $obsoleteRanges) {
        $Ranges.DeleteByDBID($obsolete.ID)
        Write-Status -Status Updated -Message "Removed obsolete managed hMailServer IP range '$($obsolete.Name)'."
    }
}

function Disable-HMailComputerSMTP {
    param($Ranges)

    $range = Get-ItemByNameOrNull -Collection $Ranges -Name 'My computer'
    if (-not $range) {
        Write-Status -Status Missing -Message "No hMailServer IP range named 'My computer' exists; nothing to disable."
    }
    elseif ($range.AllowSMTPConnections) {
        $range.AllowSMTPConnections = $false
        $range.Save()
        Write-Status -Status Updated -Message "Disabled SMTP access for hMailServer IP range 'My computer'; other settings were left unchanged."
    }
    else {
        Write-Status -Status Existing -Message "SMTP access for hMailServer IP range 'My computer' is already disabled."
    }
}

function Disable-LegacyHMailIPRanges {
    param($Ranges)

    $legacyRanges = @(
        @{ Name = 'DataSwitches'; LowerIP = '172.16.10.0'; UpperIP = '172.16.10.255' }
        @{ Name = 'StorageSwitches'; LowerIP = '172.18.10.0'; UpperIP = '172.18.10.255' }
        @{ Name = 'OutOfBandSwitches'; LowerIP = '172.20.10.0'; UpperIP = '172.20.10.255' }
    )
    foreach ($legacy in $legacyRanges) {
        $range = Get-ItemByNameOrNull -Collection $Ranges -Name $legacy.Name
        if (-not $range) {
            continue
        }
        if ($range.LowerIP -eq $legacy.LowerIP -and $range.UpperIP -eq $legacy.UpperIP -and $range.AllowSMTPConnections) {
            $range.AllowSMTPConnections = $false
            $range.Save()
            Write-Status -Status Updated -Message "Disabled SMTP access for legacy broad IP range '$($legacy.Name)'."
        }
        elseif (($range.LowerIP -ne $legacy.LowerIP -or $range.UpperIP -ne $legacy.UpperIP) -and $range.LowerIP -ne $range.UpperIP) {
            Write-Status -Status Warning -Message "Existing IP range '$($legacy.Name)' differs from the original default and was left unchanged; review it for broad SMTP access."
        }
    }
}

function Reset-HMailCollection {
    param($Collection)

    for ($index = $Collection.Count - 1; $index -ge 0; $index--) {
        $item = $Collection.Item($index)
        $Collection.DeleteByDBID($item.ID)
    }
}

function Set-HMailSenderRule {
    param($Rules, [string]$Name, [string]$Domain, [string]$From, $FromField, $ContainsMatch, $SetHeaderAction)

    $rule = Get-ItemByNameOrNull -Collection $Rules -Name $Name
    $status = 'Added'
    if ($rule) {
        $status = 'Updated'
        if ($rule.Active -and $rule.UseAND -and $rule.Criterias.Count -eq 1 -and $rule.Actions.Count -eq 1) {
            $criteria = $rule.Criterias.Item(0)
            $action = $rule.Actions.Item(0)
            if ($criteria.UsePredefined -and $criteria.PredefinedField -eq $FromField -and
                $criteria.MatchType -eq $ContainsMatch -and $criteria.MatchValue -ceq $Domain -and
                $action.Type -eq $SetHeaderAction -and $action.HeaderName -ceq 'From' -and $action.Value -ceq $From) {
                Write-Status -Status Existing -Message "Global rule '$Name' already matches the configured sender rewrite."
                return
            }
        }
    }
    else {
        $rule = $Rules.Add()
    }
    $rule.Active = $false
    $rule.Name = $Name
    $rule.UseAND = $true
    $rule.Save()
    Reset-HMailCollection -Collection $rule.Criterias
    Reset-HMailCollection -Collection $rule.Actions
    $criteria = $rule.Criterias.Add()
    $criteria.UsePredefined = $true
    $criteria.PredefinedField = $FromField
    $criteria.MatchType = $ContainsMatch
    $criteria.MatchValue = $Domain
    $criteria.Save()
    $action = $rule.Actions.Add()
    $action.Type = $SetHeaderAction
    $action.HeaderName = 'From'
    $action.Value = $From
    $action.Save()
    $rule.Active = $true
    $rule.Save()
    Write-Status -Status $status -Message "Global rule '$Name': From containing '$Domain' becomes '$From'."
}

$config = Read-HMailRelayConfig -Path $ConfigPath
$InternalDomain = $config.internalDomain
$AuthorizedDomain = $config.authorizedDomain
$AuthorizedFrom = $config.authorizedFrom
$SendGridHost = $config.sendGrid.host
$SendGridPort = [int]$config.sendGrid.port
$smtpAuthenticationEnabled = [bool]$config.smtpAuthentication.enabled
$smtpAuthenticationGenerateDemoCertificate = [bool]$config.smtpAuthentication.generateDemoCertificate
$smtpAuthenticationAddress = $config.smtpAuthentication.accountAddress
$smtpAuthenticationPort = [int]$config.smtpAuthentication.port
$allIPs = @('127.0.0.1')

$networkRanges = @(
    foreach ($device in $config.devices) {
        $bounds = ConvertFrom-IPv4Cidr -CIDR $device.cidr
        [pscustomobject]@{
            Name = $device.name
            CIDR = $device.cidr
            LowerIP = $bounds.LowerIP
            UpperIP = $bounds.UpperIP
            RequireSMTPAuthLocalToExternal = [bool]$device.requireSmtpAuth
        }
    }
)
$desiredRanges = @(
    foreach ($ip in $allIPs) {
        [pscustomobject]@{
            Name = "SendGrid relay $ip"
            LegacyName = "SendGrid relay $ip"
            LowerIP = $ip
            UpperIP = $ip
            RequireSMTPAuthLocalToExternal = $false
        }
    }
    foreach ($network in $networkRanges) {
        [pscustomobject]@{
            Name = "SendGrid relay $($network.Name)"
            LegacyName = $network.Name
            LowerIP = $network.LowerIP
            UpperIP = $network.UpperIP
            RequireSMTPAuthLocalToExternal = $network.RequireSMTPAuthLocalToExternal
        }
    }
)
foreach ($network in $networkRanges) {
    if (-not (Test-IsRFC1918Address -IPAddress $network.LowerIP) -or
        -not (Test-IsRFC1918Address -IPAddress $network.UpperIP)) {
        Write-Status -Status Warning -Message "$($network.CIDR) is not wholly within RFC1918 private address space. Use it only if it is the actual source network seen by hMailServer."
    }
}

$unauthenticatedRanges = @(
    foreach ($ip in $allIPs) {
        [pscustomobject]@{ LowerIP = $ip; UpperIP = $ip }
    }
    foreach ($network in $networkRanges) {
        if (-not $network.RequireSMTPAuthLocalToExternal) {
            [pscustomobject]@{ LowerIP = $network.LowerIP; UpperIP = $network.UpperIP }
        }
    }
)
$authenticatedRanges = @(
    foreach ($network in $networkRanges) {
        if ($network.RequireSMTPAuthLocalToExternal) {
            [pscustomobject]@{ LowerIP = $network.LowerIP; UpperIP = $network.UpperIP }
        }
    }
)
$firewallUnauthenticatedAddresses = @($unauthenticatedRanges | ForEach-Object { "$($_.LowerIP)/32" })
$firewallAuthenticatedAddresses = @($networkRanges | Where-Object { $_.RequireSMTPAuthLocalToExternal } | ForEach-Object { $_.CIDR })
$firewallUntrustedSMTP25IPv4Cidrs = @(Get-UntrustedIPv4Cidrs -AllowedRanges $unauthenticatedRanges)
$firewallUntrustedSMTP587IPv4Cidrs = @(Get-UntrustedIPv4Cidrs -AllowedRanges $authenticatedRanges)

$ruleName = 'Rewrite datawisetech.corp sender for SendGrid'
$rangeSummary = @($allIPs | ForEach-Object { "$_-$_" }) + @($networkRanges | ForEach-Object { "$($_.Name)=$($_.CIDR) ($($_.LowerIP)-$($_.UpperIP))" })
$target = "config '$ConfigPath'; domains '$InternalDomain' and '$AuthorizedDomain'; SendGrid smarthost '$SendGridHost`:$SendGridPort'; sender '$AuthorizedFrom'; synchronize managed SMTP relay ranges $($rangeSummary -join ', '), removing obsolete managed entries and migrating legacy device names; disable original broad default ranges; global From rewrite rule '$ruleName'"
$authenticatedNames = @($networkRanges | Where-Object { $_.RequireSMTPAuthLocalToExternal } | ForEach-Object { $_.Name })
$target += "; SMTP authentication account '$smtpAuthenticationAddress' on port $smtpAuthenticationPort"
$target += "; enable TLS 1.2/1.3 only and disable TLS 1.0/1.1; server cipher order: $($config.tls.preferServerCipherOrder); prioritize ChaCha20: $($config.tls.prioritizeChaCha20)"
if ($smtpAuthenticationEnabled) {
    $target += "; configure STARTTLS-required authenticated submission for $($authenticatedNames -join ', ')"
    if ($smtpAuthenticationGenerateDemoCertificate) {
        $target += '; generate a self-signed demo certificate only if both configured TLS files are missing'
    }
}
elseif ($authenticatedNames.Count -gt 0) {
    $target += "; SMTP AUTH clients $($authenticatedNames -join ', ') are pending certificate/key configuration; non-preview runs are blocked"
}
$target += "; if hMailServer is missing or incomplete, run verified hMailServer 5.7.1 Build 3116 setup from '$InstallerPath' (built-in database for new installations)"
$target += "; disable SMTP access on 'My computer', retaining the managed loopback relay"
$target += "; SMTP identity '$($config.smtpHostname)'"
if ($SkipWindowsFirewall) {
    $target += '; skip Windows Firewall changes'
}
else {
    $target += "; Windows Firewall TCP/25 unauthenticated allowlist with $($firewallUntrustedSMTP25IPv4Cidrs.Count) IPv4 block ranges"
    if ($smtpAuthenticationEnabled) {
        $target += "; TCP/$smtpAuthenticationPort authenticated allowlist with $($firewallUntrustedSMTP587IPv4Cidrs.Count) IPv4 block ranges; IPv6 blocked on both ports"
    }
    else {
        $target += '; IPv6 blocked on TCP/25'
    }
}
if ($KeepInternetRange) {
    $target += '; retain Internet SMTP range'
}
else {
    $target += '; disable Internet SMTP range'
}

$installation = Get-HMailInstallationState
$interopPath = $installation.InteropPath

if (-not $WhatIfPreference -and $authenticatedNames.Count -gt 0 -and -not $smtpAuthenticationEnabled) {
    throw 'SMTP AUTH clients are configured, but smtpAuthentication.enabled is false. Install the matching certificate and unencrypted private key, then enable authenticated submission before applying; no changes were made.'
}
if ($smtpAuthenticationEnabled) {
    $certificateFileExists = Test-Path -LiteralPath $config.smtpAuthentication.certificateFile -PathType Leaf
    $privateKeyFileExists = Test-Path -LiteralPath $config.smtpAuthentication.privateKeyFile -PathType Leaf
    if ($certificateFileExists -xor $privateKeyFileExists) {
        throw 'Only one configured TLS file exists. Refusing to apply with an incomplete certificate/key pair.'
    }
    if (-not $certificateFileExists -and -not $privateKeyFileExists) {
        if ($smtpAuthenticationGenerateDemoCertificate) {
            if ($WhatIfPreference) {
                Write-Status -Status Warning -Message "A demo self-signed certificate for '$($config.smtpHostname)' will be generated at apply time; -WhatIf will not create files."
            }
            elseif (-not (Get-Command New-SelfSignedCertificate -ErrorAction SilentlyContinue)) {
                throw 'The native New-SelfSignedCertificate cmdlet is unavailable. Use a supported Windows PowerShell installation; no external certificate tool will be downloaded or installed.'
            }
        }
        elseif ($WhatIfPreference) {
            Write-Status -Status Warning -Message 'Configured TLS certificate and private key were not found; SMTP authentication cannot be applied until both files exist.'
        }
        else {
            throw 'Configured TLS certificate and private key do not exist. No changes were made.'
        }
    }
    if ($SkipWindowsFirewall -and -not $WhatIfPreference) {
        throw 'SMTP AUTH submission requires the managed Windows Firewall allowlists; do not use -SkipWindowsFirewall.'
    }
    if (-not (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue) -and -not $WhatIfPreference) {
        throw 'Get-NetTCPConnection is required to verify the authenticated SMTP listener.'
    }
}

if (-not $WhatIfPreference -and ($installation.NeedsSetup -or -not $SkipWindowsFirewall -or $smtpAuthenticationEnabled)) {
    $isAdministrator = ([Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdministrator) {
        $elevatedParameters = @{}
        foreach ($key in $PSBoundParameters.Keys) {
            $elevatedParameters[$key] = $PSBoundParameters[$key]
        }
        $elevatedParameters.ConfigPath = $ConfigPath
        Start-HMailElevatedSession -ScriptPath $PSCommandPath -Parameters $elevatedParameters
        return
    }
}
if (-not $PSCmdlet.ShouldProcess($target, 'Apply hMailServer relay configuration')) {
    return
}
if (-not $SkipWindowsFirewall) {
    foreach ($commandName in 'Get-NetFirewallRule', 'Set-NetFirewallRule', 'Remove-NetFirewallRule', 'New-NetFirewallRule') {
        if (-not (Get-Command $commandName -ErrorAction SilentlyContinue)) {
            throw "Required Windows Firewall cmdlet not found: $commandName"
        }
    }
}

$app = $null
try {
    if (-not $interopPath -and $installation.Service) {
        throw 'An hMailServer service already exists, but its interop DLL was not found in a standard location. Resolve the existing installation before running setup.'
    }
    if ($smtpAuthenticationEnabled -and $smtpAuthenticationGenerateDemoCertificate -and
        -not (Test-Path -LiteralPath $config.smtpAuthentication.certificateFile -PathType Leaf) -and
        -not (Test-Path -LiteralPath $config.smtpAuthentication.privateKeyFile -PathType Leaf)) {
        New-HMailDemoCertificate -HostName $config.smtpHostname `
            -CertificateFile $config.smtpAuthentication.certificateFile `
            -PrivateKeyFile $config.smtpAuthentication.privateKeyFile
    }
    $adminPasswordSecure = Read-Host 'hMailServer administrator password' -AsSecureString
    $sendGridKeySecure = Read-Host 'SendGrid API key (Mail Send permission)' -AsSecureString
    $adminPassword = ConvertTo-PlainText $adminPasswordSecure
    $sendGridApiKey = ConvertTo-PlainText $sendGridKeySecure
    $smtpAuthenticationPasswordSecure = $null
    $smtpAuthenticationPassword = $null
    if ($smtpAuthenticationEnabled) {
        $smtpAuthenticationPasswordSecure = Read-Host "Password for $smtpAuthenticationAddress" -AsSecureString
        $smtpAuthenticationPassword = ConvertTo-PlainText $smtpAuthenticationPasswordSecure
    }
    if ([string]::IsNullOrWhiteSpace($adminPassword) -or [string]::IsNullOrWhiteSpace($sendGridApiKey)) {
        throw 'Both the hMailServer password and SendGrid API key are required.'
    }
    if ($smtpAuthenticationEnabled -and [string]::IsNullOrWhiteSpace($smtpAuthenticationPassword)) {
        throw 'A nonempty SMTP authentication account password is required.'
    }

    $firewallConfigured = $false
    if ($installation.NeedsSetup) {
        Write-Status -Status Warning -Message "hMailServer installation is missing or incomplete (interop DLL: $([bool]$interopPath); service: $([bool]$installation.Service); COM registration: $($installation.ComRegistered)). Running verified setup; use the existing administrator password if previously installed."
        if ($adminPassword.Length -lt 6) {
            throw 'A new hMailServer administrator password must contain at least 6 characters.'
        }
        if ($adminPassword -match '[\s"]') {
            throw 'The hMailServer 5.7.1 installer cannot safely pass a password containing whitespace or double quotes to its database setup. Choose a password without these characters.'
        }
        if ($SkipWindowsFirewall) {
            Write-Status -Status Warning -Message 'Windows Firewall changes are skipped; ensure hMailServer ports are protected before installing.'
        }
        else {
            Set-SmtpFirewallRules -UnauthenticatedRemoteAddresses $firewallUnauthenticatedAddresses `
                -AuthenticatedRemoteAddresses $firewallAuthenticatedAddresses `
                -BlockedSMTP25IPv4Cidrs $firewallUntrustedSMTP25IPv4Cidrs `
                -BlockedSMTP587IPv4Cidrs $firewallUntrustedSMTP587IPv4Cidrs `
                -AuthenticatedSubmissionEnabled $smtpAuthenticationEnabled
            $firewallConfigured = $true
            Write-Status -Status Updated -Message 'Applied the SMTP firewall allowlist before installation; these rules remain if setup fails.'
        }
        Install-HMailServer -Path $InstallerPath -AdminPassword $adminPasswordSecure
        $installation = Get-HMailInstallationState
        $interopPath = $installation.InteropPath
        if ($installation.NeedsSetup) {
            throw "Setup completed, but hMailServer is still incomplete (interop DLL: $([bool]$interopPath); service: $([bool]$installation.Service); COM registration: $($installation.ComRegistered)). Configuration was stopped."
        }
        $service = $installation.Service
        if ($service.Status -ne 'Running') {
            throw "hMailServer setup completed, but the service is $($service.Status), not running."
        }
    }

    if (-not $SkipWindowsFirewall -and -not $firewallConfigured) {
        Set-SmtpFirewallRules -UnauthenticatedRemoteAddresses $firewallUnauthenticatedAddresses `
            -AuthenticatedRemoteAddresses $firewallAuthenticatedAddresses `
            -BlockedSMTP25IPv4Cidrs $firewallUntrustedSMTP25IPv4Cidrs `
            -BlockedSMTP587IPv4Cidrs $firewallUntrustedSMTP587IPv4Cidrs `
            -AuthenticatedSubmissionEnabled $smtpAuthenticationEnabled
        $firewallConfigured = $true
        Write-Status -Status Updated -Message 'Applied Windows Firewall SMTP allowlists before changing hMailServer ranges.'
    }

    $interop = [System.Reflection.Assembly]::LoadFrom($interopPath)
    $startTlsRequired = Get-HMailEnumValue -Assembly $interop -TypeName 'eConnectionSecurity' -MemberPattern '^eCSSTARTTLSRequired$'
    if ($smtpAuthenticationEnabled) {
        $smtpProtocol = Get-HMailEnumValue -Assembly $interop -TypeName 'eSessionType' -MemberPattern '^eSTSMTP$'
    }
    $fromField = Get-HMailEnumValue -Assembly $interop -TypeName 'eRulePredefinedField' -MemberPattern '^eFTFrom$'
    $containsMatch = Get-HMailEnumValue -Assembly $interop -TypeName 'eRuleMatchType' -MemberPattern '^eMTContains$'
    $setHeaderAction = Get-HMailEnumValue -Assembly $interop -TypeName 'eRuleActionType' -MemberPattern '^eRASetHeaderValue$'

    $app = New-Object -ComObject hMailServer.Application -ErrorAction Stop
    if (-not $app.Authenticate($HMailAdminUser, $adminPassword)) {
        throw 'hMailServer authentication failed.'
    }

    $settings = $app.Settings
    $tlsSettingsChanged = Set-HMailTLSSettings -Settings $settings `
        -PreferServerCipherOrder $config.tls.preferServerCipherOrder `
        -PrioritizeChaCha20 $config.tls.prioritizeChaCha20
    if ($settings.HostName -cne $config.smtpHostname) {
        $settings.HostName = $config.smtpHostname
        Write-Status -Status Updated -Message "Set SMTP hostname to '$($config.smtpHostname)'."
    }

    Set-HMailDomain -Domains $app.Domains -Name $InternalDomain
    Set-HMailDomain -Domains $app.Domains -Name $AuthorizedDomain

    $settings.SMTPRelayer = $SendGridHost
    $settings.SMTPRelayerPort = $SendGridPort
    $settings.SMTPRelayerRequiresAuthentication = $true
    $settings.SMTPRelayerUsername = 'apikey'
    $settings.SMTPRelayerUseSSL = $false
    $settings.SMTPRelayerConnectionSecurity = $startTlsRequired
    $settings.SetSMTPRelayerPassword($sendGridApiKey)

    Sync-HMailIPRanges -Ranges $settings.SecurityRanges -DesiredRanges $desiredRanges
    Disable-HMailComputerSMTP -Ranges $settings.SecurityRanges
    Disable-LegacyHMailIPRanges -Ranges $settings.SecurityRanges

    if (-not $KeepInternetRange) {
        $internetRange = Get-ItemByNameOrNull -Collection $settings.SecurityRanges -Name 'Internet'
        if ($internetRange) {
            $internetRange.AllowSMTPConnections = $false
            $internetRange.AllowDeliveryFromLocalToLocal = $false
            $internetRange.AllowDeliveryFromLocalToRemote = $false
            $internetRange.AllowDeliveryFromRemoteToLocal = $false
            $internetRange.AllowDeliveryFromRemoteToRemote = $false
            $internetRange.Save()
            Write-Status -Status Updated -Message "Disabled SMTP access for the existing broad hMailServer IP range 'Internet'."
        }
        else {
            Write-Status -Status Missing -Message "No hMailServer IP range named 'Internet' exists; no broad range was changed."
        }
    }
    else {
        Write-Status -Status Existing -Message "Kept the hMailServer 'Internet' IP range unchanged by request."
    }

    Set-HMailSenderRule -Rules $app.Rules -Name $ruleName -Domain $InternalDomain -From $AuthorizedFrom `
        -FromField $fromField -ContainsMatch $containsMatch -SetHeaderAction $setHeaderAction

    if ($smtpAuthenticationEnabled) {
        $domain = Get-ItemByNameOrNull -Collection $app.Domains -Name $AuthorizedDomain
        if (-not $domain) {
            throw "Could not find hMailServer domain '$AuthorizedDomain' after domain provisioning."
        }
        Set-HMailSMTPAuthAccount -Domain $domain -Address $smtpAuthenticationAddress -Password $smtpAuthenticationPassword
        $certificate = Set-HMailSSLCertificate -Certificates $settings.SSLCertificates `
            -Name $config.smtpAuthentication.certificateName `
            -CertificateFile $config.smtpAuthentication.certificateFile `
            -PrivateKeyFile $config.smtpAuthentication.privateKeyFile
        Set-HMailSMTPSubmissionPort -Ports $settings.TCPIPPorts `
            -PortNumber $smtpAuthenticationPort -CertificateID $certificate.ID `
            -SMTPProtocol $smtpProtocol -StartTlsRequired $startTlsRequired
    }
    if ($smtpAuthenticationEnabled -or $tlsSettingsChanged) {
        Restart-Service -Name 'hMailServer' -Force -ErrorAction Stop
        $service = Get-Service -Name 'hMailServer' -ErrorAction Stop
        if ($service.Status -ne 'Running') {
            throw "hMailServer service is $($service.Status) after applying SMTP/TLS settings."
        }
        if ($smtpAuthenticationEnabled) {
            $hMailService = Get-CimInstance -ClassName Win32_Service -Filter "Name='hMailServer'" -ErrorAction Stop
            if (-not $hMailService -or $hMailService.ProcessId -eq 0) {
                throw 'Could not determine the running hMailServer process ID to verify the submission listener.'
            }
            $listener = Get-NetTCPConnection -State Listen -LocalPort $smtpAuthenticationPort `
                -OwningProcess $hMailService.ProcessId -ErrorAction SilentlyContinue
            if (-not $listener) {
                throw "hMailServer is running, but no TCP listener is active on port $smtpAuthenticationPort."
            }
        }
    }

    if ($SkipWindowsFirewall) {
        Write-Status -Status Warning -Message 'Skipped Windows Firewall changes; existing firewall rules were left untouched.'
    }

}
finally {
    $adminPassword = $null
    $sendGridApiKey = $null
    $adminPasswordSecure = $null
    $sendGridKeySecure = $null
    $smtpAuthenticationPassword = $null
    $smtpAuthenticationPasswordSecure = $null
    if ($app) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}