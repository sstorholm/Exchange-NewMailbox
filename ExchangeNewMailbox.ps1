##########################################################################
# ExchangeNewMailbox Script
# Creates a new mailbox and a new user and sets parameters accordingly
# Sebastian Storholm 08.03.2020
#
# Updated version 2 by Sebastian Storholm 12.05.2020
# - Added automation for distirbution groups
##########################################################################

# Specify specific Domain Contoller
$DomainController = "dc.costoso.com"

# Get credentials and create a PS session to the Exchange server
$UserCredential = Get-Credential
$SessionExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.costoso.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $SessionExchange -DisableNameChecking

# Take first and last name as input and generate necessary variations for creating the mailbox

Write-Host "================ Costoso Mail - New User ================"
$FirstName = Read-Host -Prompt 'User First Name?';
$LastName = Read-Host -Prompt 'User Last Name?';
$Alias = $FirstName.ToLower() + '.' + $LastName.ToLower()
$FullName = "$FirstName $LastName"
$UPN = $Alias + "@contoso.com"

# To avoid having longer aliases truncated and thus creating discrepancy between usernames and email aliases (aka. UPN will be 20 first chars of alias)
# Check for longer aliases than 20 chars, and if that's the case, have the admin truncate the UPN manually

if ($Alias.length -gt 20) {
    Write-Output "ERROR: Alias greater than 20 characters - Please enter user principal name manually!"
    $UPN = Read-Host "Enter User UPN"
}

# Logic for interactive menu for selecting mailbox database
# Use your mailbox database name in place of MBDB1 and MBDB2

Write-Host "================ Mailbox Database ================"
Write-Host "Select 1 for MBDB1 (mailbox size limited to 4GB)"
Write-Host "Select 2 for MBDB2 (mailbox size unlimited)"
$selectionMBDB = Read-Host

switch ($selectionMBDB)
 {
     '3' {
         $MBDB = 'MBDB1'
     } '4' {
         $MBDB = 'MBDB2'
     }
 }

# Logic for interactive menu for selecting which company a user belongs to and setting parameters accordingly

Write-Host "====================== User Company ======================"
Write-Host "Select 1 for Costoso"
Write-Host "Select 2 for Northwind Traders"
Write-Host "Select 3 for Blue Yonder Airlines"

$selectionCompany = Read-Host

switch ($selectionCompany)
 {
     '1' {
         $Company = 'Costoso'
         $Office = 'Los Angeles'
         $DistSelect = '1'
     } '2' {
         $Company = 'Northwind Traders'
         $Office = 'Helsinki'
     } '3' {
         $Company = 'Blue Yonder Airlines'
         $Office = 'London'
            }
 }

# Logic for selecting user organizational unit if it can't be determined earlier

Write-Host "================ User Organizational Unit ================"
Write-Host "Select 1 for United States of America"
Write-Host "Select 2 for Finland"
Write-Host "Select 3 for United Kingdom"
Write-Host "Select 4 for Sweden"
$selectionOU = Read-Host

switch ($selectionOU)
 {
     '1' {
         $OU = 'costoso.com/Costoso/US/Users/'
     } '2' {
         $OU = 'costoso.com/Costoso/FI/Users/'
     } '3' {
         $OU = 'costoso.com/Costoso/UK/Users/'
     } '4' {
         $OU = 'costoso.com/Costoso/SE/Users/'
     }
 }

# Display all the settings generated so far

Write-Host "================ Final Settings ================"
Write-Host "User Full Name:     $FullName"
Write-Host "User E-mail Alias:  $Alias"
Write-Host "User UPN:           $UPN"
Write-Host "Company:            $Company"
Write-Host "Office:             $Office"
Write-Host "User Mailbox DB:    $MBDB"
Write-Host "User OrgUnit CN:    $OU"

if ($DistSelect -eq 1) {
    Write-Host "User is stationed in Los Angeles"
    Write-Host "Adding user to distribution group staff_la@contoso.com"
}

###############################
# TODO: Add confirmation dialog
###############################

# Create the mailbox using the default password ChangeMe123 and all the parameters generated

New-Mailbox -Name "$FullName" -UserPrincipalName $UPN -Alias $Alias -Database $MBDB -OrganizationalUnit $OU -Password (ConvertTo-SecureString -String 'ChangeMe123' -AsPlainText -Force) -FirstName $FirstName -LastName $LastName -DomainController $DomainController

# Set the user office and company parameters since they can't be done at mailbox creation.
# ALTERNATIVE METHOD BELOW THROUGH AD INSTEAD OF EXCHANGE
# Set-User -Identity $UPN -Office $Office -Company $Company

# Logic for adding non-standard primary SMTP addresses for certian companies

if ($Company.StartsWith("North")) {
    $NWTEmail = $Alias + "@northwindtraders.com"
    Set-Mailbox -Identity $UPN -EmailAddressPolicyEnabled $false -DomainController $DomainController
    Set-Mailbox -Identity $UPN -EmailAddresses @{add="SMTP:$NWTEmail"} -DomainController $DomainController
}

# User is stationed at HQ
if ($DistSelect -eq 1) {
    Add-DistributionGroupMember -Identity "staff_la" -Member $Alias -DomainController $DomainController
}

# Create a new session to a domain controller for setting the Log On script and home directory stuff since it can't be done through Exchange

$SessionAD = New-PSSession -ComputerName $DomainController -Credential $UserCredential
Invoke-Command $SessionAD -Scriptblock { Import-Module ActiveDirectory }
Import-PSSession -Session $SessionAD -module ActiveDirectory

Set-ADUser -Identity $UPN.Replace("@contoso.com","") -ScriptPath $LogOnScript -HomeDrive 'H:' -HomeDirectory $HomeDirectory

#Set the Office and Company parameters for the user using the Exchange Powershell CMDlet

Set-ADUser -Identity $UPN -Office $Office -Company $Company

# Clean up sessions upon exit

Remove-PSSession $SessionExchange
Remove-PSSession $SessionAD
