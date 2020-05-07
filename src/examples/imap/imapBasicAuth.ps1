# Set-Location D:\dev\work\Exchange\PopImap
Set-Location (Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\")

# MyUserData.psm1 is not included in the project source control. Its content is similar to that of .\Data\UserData.psm1.
Import-Module .\Data\MyUserData.psm1
$userData = Get-UserData

Import-Module .\src\PopImap\PopImap.psd1

$logFile = "logs\imap.log"
if (Test-Path $logFile)
{
  Remove-Item $logFile
}
$receiver = Get-FileMessageReceiver -Path $logFile

$server = "outlook.office365.com"
$port = 993
$imap = Get-ImapClient -Server $server -Port $port
$imap.MessageReceivers.Add($receiver)

$imap.Connect()
$success = $imap.Logon($userData.user, $userData.password)
if ($success)
{
  $success = $imap.ExecuteCommand("list `"INBOX/`" *")
}
$imap.Close()
