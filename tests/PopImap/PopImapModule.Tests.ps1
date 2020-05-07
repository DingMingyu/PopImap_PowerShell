$rootDir = [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath "..\..\"))
. (Join-Path -Path $rootDir -ChildPath "src\PopImap\PopImap.ps1")

[datetime]$time = (Get-Date)
$msg = [Message]@{
  text = "A test message"
  sender = "C"
  timeStamp = $time
}

Describe "Message Class" -Tag "Unit" {
  It "Convert to String" {
    $str = "{0} {1} {2}" -f $time.ToString("O"), $msg.sender, $msg.text
    $msg.ToString() | Should Be $str
  }
}

Describe "MessageReceiver Class" -Tag "Unit" {
  Mock Write-Host {}
  It "Write Message To Host" {
    $receiver = [MessageReceiver]::new()
    $receiver.Receive($msg)
    Assert-MockCalled Write-Host -Exactly 1 -ParameterFilter {$Object -eq $msg} -Scope It
  }
}

Describe "FileMessageReceiver Class" -Tag "Unit" {
  Mock Add-Content {}
  It "Write Message To File" {
    $receiver = [FileMessageReceiver]::new("d:\mock.txt")
    $receiver.Receive($msg)
    Assert-MockCalled Add-Content -Exactly 1 -ParameterFilter {$Value -eq $msg} -Scope It
  }
}

Describe "Get-O365Token Function" -Tag "Unit" {
  It "Transform Token" {
    $user="TestUser@contoso.com"
    $token="abcdefghijklmnopqrstuvwxyz"
    $o365Token = Get-O365Token -accessToken $token -upn $user
    $o365Token | Should Be "dXNlcj1UZXN0VXNlckBjb250b3NvLmNvbQFhdXRoPUJlYXJlciBhYmNkZWZnaGlqa2xtbm9wcXJzdHV2d3h5egEB"
  }
}

Import-Module (Join-Path -Path $rootDir -ChildPath "Data\MyUserData.psm1")
$userData = Get-UserData

function Get-Imap
{
  $server = "outlook.office365.com"
  $port = 993
  return Get-ImapClient -Server $server -Port $port -WithDefaultOutput $false
}

Describe "ImapClient Class" -Tag "Integration" {
  It "Connect" {
    $imap = Get-Imap
    $receiver = [MemoryMessageReceiver]::new()
    $imap.MessageReceivers.Add($receiver)
    $imap.Connect()
    $receiver.Store.Count| Should Be 1
    $receiver.Store[0].text.StartsWith("* OK The Microsoft Exchange IMAP4 service is ready.") | Should Be $true
  }
  It "Logon with correct user pass" {
    $imap = Get-Imap
    $imap.Connect()
    $receiver = [MemoryMessageReceiver]::new()
    $imap.MessageReceivers.Add($receiver)
    $success = $imap.Logon($userData.user, $userData.password)
    $success | Should Be $true
    $receiver.Store.Count| Should Be 2
    $text = "0001 login {0} ****" -f $userData.user
    $receiver.Store[0].text | Should Be $text
    $receiver.Store[1].text | Should Be "0001 OK LOGIN completed.`r`n"
  }
  It "Logon with incorrect user pass" {
    $imap = Get-Imap
    $imap.Connect()
    $receiver = [MemoryMessageReceiver]::new()
    $imap.MessageReceivers.Add($receiver)
    $success = $imap.Logon($userData.user, "bad password")
    $success | Should Be $false
    $receiver.Store.Count| Should Be 2
    $receiver.Store[1].text | Should Be "0001 BAD Command Argument Error. 12`r`n"
  }
  It "Close" {
    $imap = Get-Imap
    $imap.Connect()
    $receiver = [MemoryMessageReceiver]::new()
    $imap.MessageReceivers.Add($receiver)
    $imap.Close()
    $receiver.Store.Count| Should Be 1
    $receiver.Store[0].text | Should Be "Connection is closed."
  }
}

function Get-Pop3
{
  $server = "outlook.office365.com"
  $port = 995
  return Get-Pop3Client -Server $server -Port $port -WithDefaultOutput $false
}

Describe "Pop3Client Class" -Tag "Integration" {
  It "Connect" {
    $pop3 = Get-Pop3
    $receiver = [MemoryMessageReceiver]::new()
    $pop3.MessageReceivers.Add($receiver)
    $pop3.Connect()
    $receiver.Store.Count| Should Be 1
    $receiver.Store[0].text.StartsWith("+OK The Microsoft Exchange POP3 service is ready.") | Should Be $true
  }
  It "Logon with correct user pass" {
    $pop3 = Get-Pop3
    $pop3.Connect()
    $receiver = [MemoryMessageReceiver]::new()
    $pop3.MessageReceivers.Add($receiver)
    $success = $pop3.Logon($userData.user, $userData.password)
    $success | Should Be $true
    $receiver.Store.Count| Should Be 4
    $receiver.Store[2].text | Should Be "PASS ****"
    $receiver.Store[3].text | Should Be "+OK User successfully logged on.`r`n"
  }
  It "Logon with incorrect user pass" {
    $pop3 = Get-Pop3
    $pop3.Connect()
    $receiver = [MemoryMessageReceiver]::new()
    $pop3.MessageReceivers.Add($receiver)
    $success = $pop3.Logon($userData.user, "bad password")
    $success | Should Be $false
    $receiver.Store.Count| Should Be 4
    $receiver.Store[3].text | Should Be "-ERR Logon failure: unknown user name or bad password.`r`n"
  }
  It "Close" {
    $pop3 = Get-Pop3
    $pop3.Connect()
    $receiver = [MemoryMessageReceiver]::new()
    $pop3.MessageReceivers.Add($receiver)
    $pop3.Close()
    $receiver.Store.Count| Should Be 1
    $receiver.Store[0].text | Should Be "Connection is closed."
  }
}
