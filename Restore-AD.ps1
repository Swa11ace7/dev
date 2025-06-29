<# Author : Shane Wallace 
# Student Id: 010895068
#>


# Checks for the existence of the OU named "Finance" and removes it if found
$domain = Get-ADDomain
$AdRoot = $domain.DistinguishedName
$DnsRoot = $domain.DnsRoot
$OUCanonicalName = "Finance"
$AdPath = "OU=$($OUCanonicalName),$($AdRoot)"

try {
    $OU =  Get-ADOrganizationalUnit -Identity $AdPath
    Write-Host "OU not found"
}
catch {
    $OU = $null
}

if ($OU) {
    Set-ADOrganiztionUnit -Identity $AdPath -ProtectedFromAccidentalDeletion:$false
    Remove-ADOrganizationalUnit -Identity $AdPath -Recursive -Confirm:$false
    Write-Host "OU found "
    Write-Host "OU removed"
}

New-ADOrganizationalUnit -Path $AdRoot -Name $OUCanonicalName -DisplayName $OUDisplayName 


# Import the CSV file into the OU
$csvImport = Import-Csv "$PSScriptRoot\financePersonnel.csv"


foreach ($ADUser in $csvImport) {
     $Attributes = @{
        GivenName = $($ADUser.First_Name)
        Surname = $($ADUser.Last_Name)
        SamAccountName = $($ADUser.samAccountName)
        UserPrincipalName = "$($ADUser.samAccountName)@$($DnsRoot)"
        DisplayName = "$($ADUser.First_Name) $($ADUser.Last_Name)"
        PostalCode = $($ADUser.PostalCode)
        OfficePhone = $($ADUser.OfficePhone)
        MobilePhone = $($ADUser.MobilePhone)
        Path = $AdPath
    }
    New-ADUser @Attributes
}
# Used to create a ADResults file
Get-ADUser -Filter * -SearchBase "ou=Finance,dc=consultingfirm,dc=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt
