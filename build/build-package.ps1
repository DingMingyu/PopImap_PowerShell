$modulePath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "release\PopImap"))

if (Test-Path $modulePath)
{
  Remove-Item $modulePath -Force -Recurse
}
New-Item -Path $modulePath -ItemType Directory -Force

$srcPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "..\src\PopImap")) + "\*"
Copy-Item -Path $srcPath -Destination $modulePath -Recurse