<# Author : Shane Wallace 
# Student Id: 010895068
#>

$domain = Get-ADDomain
# Creating a New User Account
$newUser = @{
    SamAccountName = "mcribb"
    Name = "Mark Cribb"
    GivenName = "Mark"
    Surname = "Cribb"
    DisplayName = "Mark Cribb"
    UserPrincipalName = "mrcibb@consultingfirm.com"
    Path = $domain.UserContainer
    AccountPassword = (ConvertTo-SecureString "P@ssword12345!" -AsPlainText -Force)
    Enabled = $true
}
New-ADUser @newUser

# Resetting a User Password
$user = "mcribb"
$newPassword = (ConvertTo-SecureString "N3WP@ssword12345!" -AsPlainText -Force)
Set-ADAccountPassword -Identity $user -NewPassword $newPassword -Reset

# Unlocking a User Account
Unlock-ADAccount -Identity $user

# Adding a User Account
$group = "CN=FinanceGroup, $($domain.DistinguishedName)"
Add-ADGroupMember -Identity $group -Members $user

# Checks for the existence of the OU named "Finance" and removes it if found
$AdRoot = $domain.DistinguishedName
$DnsRoot = $domain.DnsRoot
$OUCanonicalName = "Finance"
$OUDisplayName = "Finance"
$AdPath = "OU=$($OUCanonicalName),$($AdRoot)"

try {
    $OU =  Get-ADOrganizationalUnit -Identity $AdPath
    Write-Host "OU not found"
}
catch {
    $OU = $null
}

If ($OU) {
    Set-ADOrganizationUnit -Identity $AdPath -ProtectedFromAccidentalDeletion:$false
    Remove-ADOrganizationalUnit -Identity $AdPath -Recursive -Confirm:$false
    Write-Host "OU found "
    Write-Host "OU removed"
}

New-ADOrganizationalUnit -Path $AdRoot -Name $OUCanonicalName -DisplayName $OUDisplayName 

# Searching for Inactive User Accounts
$dataFilter = (Get-Date).AddDays(-90)
$inactiveUsers = Get-ADUser -Filter {LastLogonDate -lt $dataFilter} -Properties LastLogonDate

$inactiveUsers | Select-Object Name, LastLogonDate


# Import the CSV file into the OU
$NewADUsers = Import-Csv "c:\Source\Requirements2\financePersonnel.csv"


ForEach ($ADUser in $NewADUsers) {
    $Attributes = @{
        GivenName           = $($ADUser.First_Name)
        Surname             = $($ADUser.Last_Name)
        Name                = "$($ADUser.First_Name) $($ADUser.Last_Name)"
        DisplayName         = "$($ADUser.First_Name) $($ADUser.Last_Name)"
        SamAccountName      = $($ADUser.samAccountName)
        UserPrincipalName   = "$($ADUser.samAccountName)@$($DnsRoot)"
        PostalCode          = $($ADUser.PostalCode)
        OfficePhone         = $($ADUser.OfficePhone)
        MobilePhone         = $($ADUser.MobilePhone)
        Path                = $AdPath
    }
    New-ADUser @Attributes
}
# Used to create a ADResults file
Get-ADUser -Filter * -SearchBase "ou=Finance,dc=consultingfirm,dc=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt
