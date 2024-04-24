# Introduction

This is a collection of scripts for creating and managing Microsoft Teams Rooms. This script has been written to follow the best practices, as documented on Mircosoft's Learn portal.  [Create resource accounts for rooms and shared Teams devices - Microsoft Teams | Microsoft Learn](https://learn.microsoft.com/en-us/microsoftteams/rooms/create-resource-account?tabs=exchange-online%2Cgraph-powershell-password)

## Requirements

This script has been designed to be compatible with PowerShell 5.1 and 7.4.

Running these scripts requires the *ExchangeOnlineManagement* and *Microsoft.Graph* PowerShell modules to be installed.  They can be installed for the local user with the following commands:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

The PowerShell Execution Policy will need to be configured to allow the execution of these scripts.  Please reference Microsoft's [Set-ExecutionPolicy](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4) documention for options, including the [Unblock-File](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4#example-7-unblock-a-script-to-run-it-without-changing-the-execution-policy) option to unblock a single script.

## Creating a single Microsoft Teams Room

To create a single room, start by reviewing the "Usage Location" settings (Default: US) and the Calendar Processing privacy options.  These values should be set to either `$true` or `$false`.  For more information about the processing rules, please see the [Set-CalendarProcessing](https://learn.microsoft.com/en-us/powershell/module/exchange/set-calendarprocessing?view=exchange-ps) cmdlet documentation. The settings are in lines 28-52 of the New-MTR.ps1 script.

```powershell
# 
# Microsoft Requires the UsageLocation to be set on the user to assign
# MTR licenses
#
$usageLocation = "US"
#
# Calendar Privacy Settings
# 
# The meeting organizer will not be added to the subject of the meeting request.
$AddOrganizerToSubject = $false
# The meeting will automatically be accepted if the room is available.
$AutomateProcessing = "AutoAccept"
# Inform the meeting organizer if the meeting is declined due to conflict
$OrganizerInfo = $true
# Comments must not be deleted, otherwise Zoom, Teams and Webex Meeting
# connection details may be removed preventing One Touch Join
$DeleteComments = $false
# Remove the subject of the meeting to anonymize the meeting details
$DeleteSubject = $false
# Remove the Private flag for incoming meetings, if specified by the
# meeting organizer
$RemovePrivateProperty = $false
# Allow meetings from external senders to be processed.  It is recommended
# that this be False, due to the potential abuse from external senders.
$ProcessExternalMeetings = $false
```

After verifying that the processing rules are set appropriately, execute the New-MTR.ps1 script in a PowerShell terminal and follow the prompts to create the room.

```
PS C:\Source\MTR-Scripts> .\New-MTR.ps1
```

### New-MRT.ps1 Example:

```
PS C:\Source\MTR-Scripts> . 'C:\Source\repos\MTR-Scripts\New-MTR.ps1'
Please choose room license type (Pro or Basic): Basic
Enter the account id (ConferenceRoom@example.net): : sampleroom@example.net
Enter the Room Name: : Sample Conference Room
Enter the Room Alias: : sampleroom
Enter the Room Password: : **************
New Room Summary:
=================
Account Name:      sampleroom@example.net
Room Description:  Sample Conference Room
Room Alias:        sampleroom
Room License Type: basic
Do you wish to continue? (Y/N): y
Logging into ExchangeOnline
Welcome To Microsoft Graph!
Logging into Microsoft Graph
Getting Microsoft
Creating mailbox on ExchangeOnline
Created mailbox Sample Conference Room - sampleroom@example.net
Waiting 30 seconds for directory to synchronize before updating.
Updating mailbox calendar processing rules
Querying the AzureAD account
Updating Usage Location and password expiration per Mircrosoft recommendations
True
Assigning the Microsoft License
Done creating Microsoft Teams Room for sampleroom@example.net
```

## License

Copyright 2024 New York Technology Company

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
