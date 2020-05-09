# Set-Location D:\dev\work\Exchange\PopImap
Set-Location (Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\")

# MyUserData.psm1 is not included in the project source control. Its content is similar to that of .\Data\UserData.psm1.
Import-Module .\Data\MyUserData.psm1
$userData = Get-UserData

Import-Module .\src\PopImap\PopImap.psd1

$server = "outlook.office365.com"
$port = 995
$logFile = "logs\pop3_{0:yyyyMMdd}.log" -f (Get-Date)

$pop3 = Get-Pop3Client -Server $server -Port $port -OutputPath $logFile

$pop3.Connect()
$success = $pop3.Logon($userData.user, $userData.password)
if ($success)
{
  $pop3.ExecuteCommand("LIST")
}
$pop3.Close()
