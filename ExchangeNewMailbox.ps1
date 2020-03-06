$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.costoso.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking


Write-Host "================ Costoso Mail - New User ================"
$FirstName = Read-Host -Prompt 'User First Name?';
$LastName = Read-Host -Prompt 'User Last Name?';
$Alias = $FirstName.ToLower() + '.' + $LastName.ToLower()
$FullName = "$FirstName $LastName"
$UPN = $Alias + "@contoso.com"
if ($Alias.length -gt 20) {
    Write-Output "ERROR: Alias greater than 20 characters - Please enter user principal name manually!"
    $UPN = Read-Host "Enter User UPN"
}
Write-Host "================ Mailbox Database ================"
Write-Host "Select 1 for MBDB1 (mailbox size limited to 4GB)"
Write-Host "Select 2 for MBDB2 (mailbox size unlimited)"
$selectionMBDB = Read-Host 

switch ($selectionMBDB)
 {
     '3' {
         $MBDB = 'MBDB1'
     } '4' {
         $MBDB = 'MBDB1'
     }
 }
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
Write-Host "================ Final Settings ================"
Write-Host "User Full Name: $FullName"
Write-Host "User E-mail Alias: $Alias"
Write-Host "User UPN: $UPN"
Write-Host "User Mailbox DB: $MBDB"
Write-Host "User OrgUnit CN: $OU"

New-Mailbox -Name "$FullName" -UserPrincipalName $UPN -Alias $Alias -Database $MBDB -OrganizationalUnit $OU -Password (ConvertTo-SecureString -String 'ChangeMe123' -AsPlainText -Force) -FirstName $FirstName -LastName $LastName

Remove-PSSession $Session
