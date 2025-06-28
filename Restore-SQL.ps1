<# Author : Shane Wallace 
# Student Id: 010895068
#>


if (Get-Module -Name sqlps) { Remove-Module sqlps}
Import-Module -Name SqlServer

Write-Host -ForegroundColor Cyan "[SQL]: Staring SQL Tasks"

$sqlServerInstanceName = "SRV19-PRIMARY\SQLEXPRESS"

#Sets the database name with this variable
$databaseName = "ClientDB"

$sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstanceName

$databaseObject = Get-SqlDAtabase -ServerInstance $sqlServerInstanceName -Name $databaseName -ErrorAction SilentlyContinue
if ($databaseObject) {
    Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) database has been found and removed"
    
    # Kills all processes in to the database
    $sqlServerObject.killAllProcesses($databaseName)

    #Sets dabase object to Single user mode
    $databaseObject.UserAccess = "Single"

    # Deletes the database
    $databaseObject.drop()
}
else {
    White-Host "[SQL]: $($databaseName) database not found"
}


# Creates a new database
$databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName
$databaseObject.Create()
White-Host -ForegroundColor Cyan "[SQL]: Database Created:[$($sqlServerInstanceName)].[$($databaseName)]"


#Creates a new database table
$schema = "dbo"
$tableName = "Client_A_Contacts"
Invoke-Sqlcmd -ServerInstance $sqlServerInstanceName -Database $databaseName -InputFile $PSScriptRoot\CreateTable_Client_A_Contact.sql
Write-Host -ForegroundColor Cyan "[SQL]: Table Created:[$($sqlServerInstanceName)].[$($databaseName)].[$($schema)].[$($tableName)]"


# Inserts new data into the table
$InsertQuery = "INSERT INTO [$($schema)].[$($tableName)] (FirstName, LastName, DisplayName, PostalCode, OfficePhone, MobilePhone)"
$NewClient = Import-Csv -Path $PSScriptRoot\NewClientData.csv

Write-Host -ForegroundColor Cyan "[SQL]: Inserting Data"
foreach ($NewClient in $NewClient) {
    $Values = " VALUES ('$(NewClient.FirstName)',
     '$($NewClient.LastName)', 
     '$($NewClient.DisplayName)', 
     '$($NewClient.PostalCode)', 
     '$($NewClient.OfficePhone)', 
     '$($NewClient.MobilePhone)')"
     $query = $InsertQuery + $Values
     Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $query
}
Invoke-Sqlcmd -Database ClientDB -ServerInstance .\SQLEXPRESS -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SqlResults.txt