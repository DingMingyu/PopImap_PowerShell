class UserData
{
  [string]$user = "xxxxn@xxxxx.onmicrosoft.com"
  [string]$password = "xxxxx"
  [guid]$clientId = "xxxxxxxxxxxxxxxxxxxxxx"
  [guid]$tenantId = "xxxxxxxxxxxxxxxxxxxxxx"
  [string]$nuGetApiKey = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
}

function Get-UserData
{ 
  return [UserData]::new()
}

Export-ModuleMember -Function Get-UserData