class Message
{
  [string]$sender
  [string]$text
  [datetime]$timeStamp = [datetime]::Now

  [string]ToString()
  {
    return "{0} {1} {2}" -f $this.timeStamp.ToString("O"), $this.sender, $this.text
  }
}

class MessageReceiver
{
  [void]Receive([Message]$msg)
  {
    Write-Host $msg
  }
}

class FileMessageReceiver : MessageReceiver
{
  [string]$path

  FileMessageReceiver([string]$path)
  {
    $this.path = $path
    $folder = Split-Path -Path $path
    if ($folder -and !(Test-Path -Path $folder))
    {
      $null = New-Item -ItemType Directory -Path $folder
    }
  }

  [void]Receive([Message]$msg)
  {
    Add-Content -Path $this.path -Value $msg
  }
}

class MemoryMessageReceiver : MessageReceiver
{
  $Store = [System.Collections.Generic.List[Message]]::new()

  MemoryMessageReceiver()
  {
  }

  [void]Receive([Message]$msg)
  {
    $this.Store.Add($msg)
  }
}

class TcpClient
{
  [string]$server
  [int]$port
  [System.IO.StreamReader]$reader
  [System.IO.StreamWriter]$writer
  [System.Net.Sockets.TcpClient]$client
  $MessageReceivers = [System.Collections.Generic.List[MessageReceiver]]::new()

  TcpClient([string]$server, [int]$port)
  {
    $this.server = $server
    $this.port = $port
  }

  [void]Close()
  {
    if ($this.client)
    {
      $this.client.Close()
      $this.SendMessage("C", "Connection is closed.")
    }
  }

  [void]Connect()
  {
    $this.client = New-Object System.Net.Sockets.TcpClient($this.server, $this.port)
    $this.client.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]"Socket", [System.Net.Sockets.SocketOptionName]"KeepAlive", $true)
    $strm = $this.client.GetStream()
    [System.Net.Security.RemoteCertificateValidationCallback]$c={return $true}
    $strm = New-Object System.Net.Security.SslStream($strm, $false, $c)
    $strm.AuthenticateAsClient($this.server)
    $this.reader = New-Object System.IO.StreamReader($strm, [System.Text.Encoding]::ASCII)
    $this.writer = New-Object System.IO.StreamWriter($strm, [System.Text.Encoding]::ASCII)
    $this.writer.NewLine = "`r`n"
    $this.writer.AutoFlush = $true
    $line = $this.reader.ReadLine()
    $this.SendMessage("S", $line)
  }

  [void]SendMessage([string]$from, [string]$text)
  {
    [Message]$msg = @{
      text = $text
      sender = $from
    }
    foreach($receiver in $this.MessageReceivers)
    {
      $receiver.Receive($msg)
    }
  }
}

function Get-O365Token([string]$accessToken, [string]$upn)
{
  [char]$ctrlA = 1
  $token = "user=" + $upn + $ctrlA + "auth=Bearer " + $accessToken + $ctrlA + $ctrlA
  $bytes = [System.Text.Encoding]::ASCII.GetBytes($token)
  $encodedToken = [Convert]::ToBase64String($bytes)
  return $encodedToken
}

class ImapClient : TcpClient
{
  [int]$tag = 0

  ImapClient([string]$server, [int]$port) : base($server, $port)
  {
  }

  [bool]ExecuteCommand([string]$cmd)
  {
    $tagText = $this.getNextTagText()
    $cmdText = [string]::Format("{0} {1}", $tagText, $cmd)
    $this.ExecutePartial($cmdText)
    $sb = [System.Text.StringBuilder]::new()
    do 
    {
      $line = $this.reader.ReadLine()
      $sb.AppendLine($line)
    }
    while(!$line.StartsWith($tagText + " "))
    $this.SendMessage("S", $sb.ToString())
    $parts = $line.Split(" ")
    return $parts.Length -gt 1 -and $parts[1] -eq "OK"
  }

  [void]ExecutePartial($cmd)
  {
    $this.SendMessage("C", $this.redact($cmd))
    $this.writer.WriteLine($cmd)
  }

  [string]ReadPartial()
  {
    $result = $this.reader.ReadLine()
    $this.SendMessage("S", $result)
    return $result
  }

  [bool]SaveEmail([string]$folder, [string]$content)
  {
    $t = $this.getNextTagText()
    $cmd = "$t APPEND $folder {" + $content.Length + "}"
    $this.ExecutePartial($cmd)
    $line = $this.ReadPartial()
    if ($line -and $line.StartsWith("+ Ready for additional command text."))
    {
      $this.ExecutePartial($content)
      $line = $this.ReadPartial()
      $parts = $line.Split(" ")
      return $parts.Length -gt 1 -and $parts[1] -eq "OK"
    }
    return $false
  }

  [bool]Logon([string]$user, [string]$pass)
  {
    return $this.ExecuteCommand("login $user $pass")
  }

  [bool]XOauth2Authenticate([string]$oAuthToken)
  {
    $cmd = "AUTHENTICATE XOAUTH2 $oAuthToken"
	  return $this.ExecuteCommand($cmd)
  }

  [bool]O365Authenticate([string]$accessToken, [string]$upn)
  {
    $token = Get-O365Token -accessToken $accessToken -upn $upn
    return $this.XOauth2Authenticate($token)
  }

  [string]redact([string]$text)
  {
    $parts = $text.Split(" ")
    if ($parts.Length -ge 4 -and $parts[1] -eq "login")
    {
      $parts[3] = "****"
      return [string]::Join(" ", $parts[0..3])
    }
    return $text
  }

  [string]getNextTagText()
  {
    return (++$this.tag).ToString("D4")
  }
}

class Pop3Client : TcpClient
{
  Pop3Client([string]$server, [int]$port) : base($server, $port)
  {
  }

  [bool]ExecuteCommand([string]$cmd)
  {
    $result = $false
    $this.SendMessage("C", $this.redact($cmd))
    $this.writer.WriteLine($cmd)
    $sb = [System.Text.StringBuilder]::new()
    $line = $this.reader.ReadLine()
    $sb.AppendLine($line)

    $parts = $cmd.Split(" ")
    $c1 = $parts[0].ToUpper()
    $hasMore = ($c1 -eq "RETR") `
      -or ($c1 -eq "TOP") `
      -or ($c1 -eq "CAPA") `
      -or (($c1 -eq "LIST") -and ($parts.Length -eq 1)) `
      -or (($c1 -eq "UIDL") -and ($parts.Length -eq 1))
    $result = $line.StartsWith("+")
    if ($hasMore -and $line.StartsWith("+OK"))
    {
      do
      {
        $line = $this.reader.ReadLine()
        $sb.AppendLine($line)
      }
      while($line -ne ".")
    }
    $this.SendMessage("S", $sb.ToString())
    return $result
  }

  [bool]Logon([string]$user, [string]$pass)
  {
    $result = $this.ExecuteCommand("USER $user")
    if ($result)
    {
      $result = $this.ExecuteCommand("PASS $pass")
    }
    return $result
  }

  [bool]XOauth2Authenticate([string]$oAuthToken)
  {
    $result = $this.ExecuteCommand("AUTH XOAUTH2")
    if ($result)
    {
      $result = $this.ExecuteCommand($oAuthToken)
    }
    return $result
  }

  [bool]O365Authenticate([string]$accessToken, [string]$upn)
  {
    $token = Get-O365Token -accessToken $accessToken -upn $upn
    return $this.XOauth2Authenticate($token)
  }

  [string]redact([string]$text)
  {
    if ($text.StartsWith("PASS "))
    {
      return "PASS ****"
    }
    return $text
  }
}

function Add-Output {
  [OutputType([TcpClient])]
  param 
  (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [TcpClient]$client,
    [Parameter(Mandatory=$true)] [bool]$WithDefaultOutput,
    [Parameter(Mandatory=$true)] [bool]$WithFileOutput,
    [Parameter(Mandatory=$true)] [string]$OutputPath
  )
  process
  {
    if ($WithDefaultOutput)
    {
      $client.MessageReceivers.Add([MessageReceiver]::new())
    }
    if ($WithFileOutput)
    {
      $client.MessageReceivers.Add([FileMessageReceiver]::new($OutputPath))
    }
    return $client
  }
}

<#
.SYNOPSIS
    Get IMAP client to communicate with an IMAP server.
.DESCRIPTION
    This cmdlet returns an ImapClient object.
.EXAMPLE
    PS C:\>$imap = Get-ImapClient -Server "outlook.office365.com" -Port 993
    PS C:\>$imap.Connect()
    PS C:\>$imap.Logon("user@contoso.com", "<password>")
    PS C:\>$imap.ExecuteCommand("list `"INBOX/`" *")
    PS C:\>$imap.Close()
#>
function Get-ImapClient
{
  [OutputType([ImapClient])]
  param
  (
    [Parameter(Mandatory=$true)] [string]$Server,
    [Parameter(Mandatory=$true)] [int]$Port,
    [Parameter(Mandatory=$false)] [bool]$WithDefaultOutput = $true,
    [Parameter(Mandatory=$false)] [bool]$WithFileOutput = $true,
    [Parameter(Mandatory=$false)] [string]$OutputPath = "logs\imap.log"
  )  
  process
  {
    $client = [ImapClient]::new($Server, $Port)
    return $client | Add-Output -WithDefaultOutput $WithDefaultOutput -WithFileOutput $WithFileOutput -OutputPath $OutputPath
  }  
}

<#
.SYNOPSIS
    Get POP3 client to communicate with an POP3 server.
.DESCRIPTION
    This cmdlet returns an Pop3Client object.
.EXAMPLE
    PS C:\>$pop3 = Get-Pop3Client -Server "outlook.office365.com" -Port 995
    PS C:\>$pop3.Connect()
    PS C:\>$pop3.Logon("user@contoso.com", "<password>")
    PS C:\>$pop3.ExecuteCommand("LIST")
    PS C:\>$pop3.Close()
#>
function Get-Pop3Client
{
  [OutputType([Pop3Client])]
  param
  (
    [Parameter(Mandatory=$true)] [string]$Server,
    [Parameter(Mandatory=$true)] [int]$Port,
    [Parameter(Mandatory=$false)] [bool]$WithDefaultOutput = $true,
    [Parameter(Mandatory=$false)] [bool]$WithFileOutput = $true,
    [Parameter(Mandatory=$false)] [string]$OutputPath = "logs\pop3.log"
  )  
  process
  {
    $client = [Pop3Client]::new($Server, $Port)
    return $client | Add-Output -WithDefaultOutput $WithDefaultOutput -WithFileOutput $WithFileOutput -OutputPath $OutputPath
  }  
}
