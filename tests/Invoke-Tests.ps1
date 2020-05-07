$rootDir = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "..\"))
$testFiles = [System.IO.Directory]::GetFiles($rootDir, "*.tests.ps1", "AllDirectories")                                                                                         
foreach ($testFile in $testFiles)
{
  "Testing $testFile"
  Invoke-Pester $testFile
}
