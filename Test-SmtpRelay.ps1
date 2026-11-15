#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$From = 'sender@example.com',

    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$To = 'recipient@example.net',

    [string]$SmtpServer = '127.0.0.1',

    [ValidateRange(1, 65535)]
    [int]$Port = 25
)

$tcpClient = $null
$smtpClient = $null
$message = $null

try {
    Write-Host '[1/3] Checking relay connection...' -ForegroundColor Cyan
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    $connectResult = $tcpClient.BeginConnect($SmtpServer, $Port, $null, $null)
    if (-not $connectResult.AsyncWaitHandle.WaitOne(10000, $false)) {
        throw "Timed out connecting to $SmtpServer`:$Port. Check the relay address, port, and firewall."
    }
    $tcpClient.EndConnect($connectResult)
    $connectResult.AsyncWaitHandle.Close()
    $tcpClient.Close()
    $tcpClient = $null
    Write-Host '[+] Relay is reachable.' -ForegroundColor Green

    Write-Host '[2/3] Submitting a test message without SMTP authentication or TLS...' -ForegroundColor Cyan
    $subject = 'SMTP relay test ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $body = @"
This is a test message submitted through the internal SMTP relay.

Relay: $SmtpServer`:$Port
From: $From
To: $To
Submitted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')
"@

    $message = New-Object System.Net.Mail.MailMessage($From, $To, $subject, $body)
    $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
    $smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
    $smtpClient.EnableSsl = $false
    $smtpClient.UseDefaultCredentials = $false
    $smtpClient.Credentials = $null
    $smtpClient.Timeout = 15000
    $smtpClient.Send($message)

    Write-Host '[+] Relay accepted the message.' -ForegroundColor Green
    Write-Host '[3/3] Check the hMailServer queue/logs and the recipient inbox; acceptance does not confirm final delivery.' -ForegroundColor Yellow
}
catch {
    Write-Host ('[X] Relay test failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
    throw
}
finally {
    if ($message) { $message.Dispose() }
    if ($smtpClient) { $smtpClient.Dispose() }
    if ($tcpClient) { $tcpClient.Close() }
}