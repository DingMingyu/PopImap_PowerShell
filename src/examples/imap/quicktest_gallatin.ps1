

# P@ssroad01!
# IrxHBQB~P6uh_5g0Q5IiUN52F_.3KfWT6q

$clientId = "9d174a7b-ca8d-4594-b596-5c39a6a513d5"
$tenantId = "4d9f9865-f5e1-41c9-bfae-19d16aa6a2f4"
$scopes = @("https://partner.outlook.cn/IMAP.AccessAsUser.All")
$token = Get-MsalToken -TenantId $tenantId -ClientId $clientId -AzureCloudInstance AzureChina -Interactive -Scopes $scopes

$server="partner.outlook.cn"
$port = 993
$mailbox = "yuanliang@yuliantest.partner.onmschina.cn"
$logFile = "d:\logs\imap_{0:yyyyMMdd}.log" -f (Get-Date)

$imap = Get-ImapClient -Server $server -Port $port -OutputPath $logFile
$imap.Connect()
$imap.O365Authenticate($token.AccessToken, $mailbox)
$imap.ExecuteCommand('list "" *')
$imap.Close()

################################

$clientId = "9d174a7b-ca8d-4594-b596-5c39a6a513d5"
$tenantId = "4d9f9865-f5e1-41c9-bfae-19d16aa6a2f4"
$clientSecret = "IrxHBQB~P6uh_5g0Q5IiUN52F_.3KfWT6q"
$clientSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$scopes = @("https://partner.outlook.cn/.default")
$token = Get-MsalToken -TenantId $tenantId -ClientId $clientId -AzureCloudInstance AzureChina -Scopes $scopes -ClientSecret $clientSecret

$server="partner.outlook.cn"
$port = 993
$mailbox = "yuanliang@yuliantest.partner.onmschina.cn"
$logFile = "d:\logs\imap_{0:yyyyMMdd}.log" -f (Get-Date)

$imap = Get-ImapClient -Server $server -Port $port -OutputPath $logFile
$imap.Connect()
$imap.O365Authenticate($token.AccessToken, $mailbox)
$imap.ExecuteCommand('list "" *')
$imap.Close()
