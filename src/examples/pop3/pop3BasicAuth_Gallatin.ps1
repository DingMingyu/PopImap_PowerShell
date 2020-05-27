Set-Location D:\dev\work\Exchange\PopImap_PowerShell
# Set-Location (Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\")

# MyUserData.psm1 is not included in the project source control. Its content is similar to that of .\Data\UserData.psm1.
Import-Module .\Data\MyUserData2.psm1
$userData = Get-UserData

Import-Module .\src\PopImap\PopImap.psd1

$server = "partner.outlook.cn"
$port = 995

$pop3 = Get-Pop3Client -Server $server -Port $port
$pop3.Connect()
$null = $pop3.Logon($userData.user, $userData.password)
$null = $pop3.ExecuteCommand("STAT")
$null = $pop3.ExecuteCommand("RETR 36")
$pop3.Close()
