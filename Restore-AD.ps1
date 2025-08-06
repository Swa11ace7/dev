# path of the new OU to add

$newOU = "OU=Finance,DC=consultingfirm,DC=com"


# check if AD OU exists, delete and recreate if so

if ([adsi]::Exists("LDAP://$newOU")) {

Write-Host "$newOU already exists. Deleting and recreating"

Remove-ADOrganizationalUnit -Identity $newOU -Recursive -confirm:$false

}


New-ADOrganizationalUnit -Name Finance -ProtectedFromAccidentalDeletion $False

Write-Host "created $newOU"


# use $PSScriptRoot instead of absolute (full) path to file so that the script will work on other systems and locations

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
