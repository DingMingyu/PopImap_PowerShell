$version = "0.1"
$modulePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "release\$version"))

$userdataModulePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "..\userdata\MyUserData.psm1"))
Import-Module $userdataModulePath
$userData = Get-UserData$nuGetApiKey = 

Publish-Module -Path $modulePath -NuGetApiKey $userData.nuGetApiKey
