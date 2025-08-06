<# Author : Shane Wallace 
# Student Id: 010895068
#>

# Adds the path of the new OU to add

$newOU = "OU=Finance,DC=consultingfirm,DC=com"


# check if AD OU exists, deletes and then recreates new OU if found

if ([adsi]::Exists("LDAP://$newOU")) {

Write-Host "$newOU already exists. Deleting and recreating a new OU"

Remove-ADOrganizationalUnit -Identity $newOU -Recursive -confirm:$false

}


New-ADOrganizationalUnit -Name Finance -ProtectedFromAccidentalDeletion $False

Write-Host "created $newOU"

$NewAD = Import-CSV $PSScriptRoot\financePersonnel.csv


foreach ($ADuser in $NewAD)

{


$First = $ADUser.First_Name

$Last = $ADUser.Last_Name

$Name = $First + " " + $Last

$SamName = $ADUser.samAccount

$Postal = $ADUser.PostalCode

$Office = $ADUser.OfficePhone

$Mobile = $ADUser.MobilePhone


New-AdUser -GivenName $First -Surname $Last -Name $Name -SamAccountName $SamName -DisplayName $Name -PostalCode $Postal -MobilePhone $Mobile -OfficePhone $Office -Path $newOU

}


Get-ADUser -Filter * -SearchBase “ou=Finance,dc=consultingfirm,dc=com” -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > $psscriptroot\AdResults.txt

