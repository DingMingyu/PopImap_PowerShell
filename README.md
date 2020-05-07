## PopImap
The project creates a PowerShell module that helps user communicate with a mail server with IMAP(https://tools.ietf.org/html/rfc3501) server or POP3(https://tools.ietf.org/html/rfc1081).

The module supports both Basic Authentication(user/password) and OAuth authentication, with special support for Microsft's Office 365. 
In order to apply OAuth authentication, you need to fetch a valid AccessToken, e.g. from Azure. The AccessToken fetching feature is not included in this module, but you can find some 3rd party modules for the purpose, from the example scripts in the project.

## Getting Started
Please go through the scripts under .\scr\examples and you'll know how to communicate with a mail server through IMAP or POP3.

## Test
.\tests\Invoke-Tests.ps1

## Build
.\build\Build-Package.ps1
