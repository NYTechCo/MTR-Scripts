# Copyright 2024 New York Technology Company
#
# Permission is hereby granted, free of charge, to an1y person obtaining a 
# copy of this software and associated documentation files (the “Software”),
# to deal in the Software without restriction, including without limitation 
# the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the 
# Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
# IN THE SOFTWARE.
#

# This module requires the ExchangeOnline and Microsoft Graph PowerShell 
# modules.  They can be installed with the commands below
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Install-Module Microsoft.Graph -Scope CurrentUser
# Install-Module ExchangeOnlineManagement -Scope CurrentUser

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


# Collect Teams Room Account Details
$licenseTypePrompt = "Please choose room license type (Pro or Basic)"
$liceseTypeResponse = ""

while (!(($liceseTypeResponse.ToLower() -eq "pro") -xor ($liceseTypeResponse.ToLower() -eq "basic")))
{
    $liceseTypeResponse = Read-Host -Prompt $licenseTypePrompt
}

$accountNamePrompt = "Enter the account id (ConferenceRoom@example.net): "
$accountName = Read-Host -Prompt $accountNamePrompt

$namePrompt = "Enter the Room Name: "
$name = Read-Host -Prompt $namePrompt

$aliasPrompt = "Enter the Room Alias: "
$alias = Read-Host -Prompt $aliasPrompt

$passwordPrompt = "Enter the Room Password: "
$password = Read-Host -Prompt $passwordPrompt -AsSecureString 

# Confirm Room Summary
Write-Host "New Room Summary:"
Write-Host "================="
Write-Host "Account Name:      $accountName"
Write-Host "Room Description:  $name"
Write-Host "Room Alias:        $alias"
Write-Host "Room License Type: $liceseTypeResponse"

$proceedPrompt = "Do you wish to continue? (Y/N)"
$proceed = Read-Host -Prompt $proceedPrompt

if (!($proceed.ToUpper() -eq "Y") -or ($proceed.ToUpper() -eq "YES"))
{
    Write-Host "Exiting..."
    Exit
}
# Connect and authenticate to ExchangeOnline
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
{
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "Logging into ExchangeOnline"
}
else 
{
    Write-Output "Missing ExchangeOnline Powershell Module"
    Exit 
}

# Connect and authenticate to Microsoft Graph
if (Get-Module -ListAvailable -Name Microsoft.Graph)
{
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.Read.All" 
    Write-Host "Logging into Microsoft Graph"
}
else
{
    Write-Output "Missing Microsoft Graph Powershell Module"
    Exit 
}

Write-Host "Getting Microsoft "
# Microsoft Teams Rooms SKU Part Number for MS Graph API
$mtrBasic = "Microsoft_Teams_Rooms_Basic"
$mtrPro = "Microsoft_Teams_Rooms_Pro"


if ($liceseTypeResponse.ToLower() -eq "basic" )
{
    $mtrLicenseType = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $mtrBasic
}
elseif ($liceseTypeResponse.ToLower() -eq "pro") {
    $mtrLicenseType = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $mtrPro
}

if (!$mtrLicenseType)
{
    Write-Host "No Microsoft Teams Room Licenses Found ($licenseTypeResponse)"
    Exit
}

if ($mtrLicenseType.consumedunits -ge $mtrLicenseType.prepaidunits.enabled)
{
    Write-Host "No $($mtrLicenseType.SkuPartNumber) licenses available." 
    Write-Host "Used Licenses: $($mtrLicenseType.consumedunits)"
    Write-Host "Available Licenses: $($mtrLicenseType.prepaidunits.enabled)"
    Exit
}

Write-Host "Creating mailbox on ExchangeOnline"
# Add new Room Resource Mailbox

$mailbox = New-Mailbox -MicrosoftOnlineServicesID $accountName -Name $name `
                       -Alias $alias -Room -EnableRoomMailboxAccount $true  `
                       -RoomMailboxPassword $password



Write-Host "Created mailbox $($mailbox.Name) - $($mailbox.MicrosoftOnlineServicesID)"
Write-Host "Waiting 30 seconds for directory to synchronize before updating."
# Waiting for 30 seconds to let the directory syncronize before updating 
# the account details
Start-Sleep 30

Write-Host "Updating mailbox calendar processing rules"
# Set the Calendar Processing Rules based on the processing and privacy 
# rules defined above
Set-CalendarProcessing -Identity $accountName `
                       -AutomateProcessing $AutomateProcessing `
                       -AddOrganizerToSubject $AddOrganizerToSubject `
                       -OrganizerInfo $OrganizerInfo `
                       -DeleteComments $DeleteComments `
                       -DeleteSubject $DeleteSubject `
                       -RemovePrivateProperty $RemovePrivateProperty `
                       -ProcessExternalMeetingMessages $ProcessExternalMeetings
                    
Write-Host "Querying the AzureAD account"
Get-MgUser -Search "mail:$accountName" -ConsistencyLevel eventual

Write-Host "Updating Usage Location and password expiration per Mircrosoft recommendations"
Update-MgUser -UserId $account.Id -PasswordPolicies DisablePasswordExpiration `
              -UsageLocation $usageLocation -PassThru

Write-Host "Assigning the Microsoft License"
Set-MgUserLicense -UserId $account.Id `
                              -AddLicenses @{ SkuId = $mtrLicenseType.SkuId } `
                              -RemoveLicenses @()

Write-Host "Done creating Microsoft Teams Room for $accountName"
