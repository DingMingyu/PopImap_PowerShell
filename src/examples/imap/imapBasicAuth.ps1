# Set-Location D:\dev\work\Exchange\PopImap
Set-Location (Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\")

# MyUserData.psm1 is not included in the project source control. Its content is similar to that of .\Data\UserData.psm1.
Import-Module .\Data\MyUserData.psm1
$userData = Get-UserData

Import-Module .\src\PopImap\PopImap.psd1

$server = "outlook.office365.com"
$port = 993
$logFile = "logs\imap_{0:yyyyMMdd}.log" -f (Get-Date)

$imap = Get-ImapClient -Server $server -Port $port -OutputPath $logFile

$imap.Connect()
$success = $imap.Logon($userData.user, $userData.password)
if ($success)
{
  $success = $imap.ExecuteCommand("list `"INBOX/`" *")
}
$imap.Close()
