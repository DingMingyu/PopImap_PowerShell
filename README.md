## PopImap
The project creates a PowerShell module that helps user communicate with a mail server with IMAP(https://tools.ietf.org/html/rfc3501) server or POP3(https://tools.ietf.org/html/rfc1081).

The module supports both Basic Authentication(user/password) and OAuth authentication, with special support for Microsft's Office 365. 
In order to apply OAuth authentication, you need to fetch a valid AccessToken, e.g. from Azure. The AccessToken fetching feature is not included in this module, but you can find some 3rd party modules for the purpose, from the example scripts in the project.

## Getting Started
Run the below command to install the module.
  Install-Module -Name PopImap
Please go through the scripts under .\scr\examples and you'll know how to communicate with a mail server through IMAP or POP3.
To make the examples and tests work, you need to create .\data\myUserData.psm1 in your local project folder. The file content is similar to .\data\userData.psm1. The file cannot be not included in the source control because it contains user info, e.g. password. 

## Test
.\tests\Invoke-Tests.ps1

## Build
.\build\Build-Package.ps1
