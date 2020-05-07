$version = "0.1"
$modulePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "release\$version"))

if (Test-Path $modulePath)
{
  Remove-Item "$modulePath\*" -Force -Recurse
}

$srcPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "..\src\PopImap"))
Copy-Item -Path $srcPath -Destination $modulePath -Recurse