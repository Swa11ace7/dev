<# Author : Shane Wallace 
# Student Id: 010895068
#>


if (Get-Module -Name sqlps) { Remove-Module sqlps}
Import-Module -Name SqlServer

Write-Host -ForegroundColor Cyan "[SQL]: Staring SQL Tasks"

$sqlServerInstanceName = "SRV19-PRIMARY\SQLEXPRESS"

#Sets the database name with this variable
$databaseName = "ClientDB"


# Create Connection to Server
$connectionString = "Data Source=$($sqlServerInstanceName);Initial Catalog=master;Inegrated Security=True;"
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = $connectionString
$sqlConnection.Open()

# Create query to detect if database alreadt exists
$detectQuery = "SELECT name FROM sys.databases WHERE name = '$($databaseName)'"

# Create Command object to run query 
$sqlCommand = New-Object System.Data.SqlClient.SqlCommand $detectQuery, $sqlConnection
$sqlReader = $sqlCommand.ExecuteReader()

if($sqlReader.Read())
{
    $sqlReader.Close()
    Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) Database Found. Deleting"
    # Create query and command to dorp database
    $dropQuery = "ALTER DATABASE [$($databaseName)] SET SINGLE_USER WITH ROLLBACK IMMEDIATE DROP DATABASE [$($databaseName)]"
    $sqlCommand.CommandText = $dropQuery
    $sqlCommand.ExecuteNonQuery() | Out-Null
}
$sqlReader.Close()
$sqlReader.Dispose()

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
    Write-Host "[SQL]: $($databaseName) database not found"
}


# Creates a new database
$databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName
$databaseObject.Create()
Write-Host -ForegroundColor Cyan "[SQL]: Database Created:[$($sqlServerInstanceName)].[$($databaseName)]"


#Creates a new database table
$schema = "dbo"
$tableName = "Client_A_Contacts"
Invoke-Sqlcmd -ServerInstance $sqlServerInstanceName -Database $databaseName -InputFile $PSScriptRoot\CreateTable_Client_A_Contact.sql
Write-Host -ForegroundColor Cyan "[SQL]: Table Created:[$($sqlServerInstanceName)].[$($databaseName)].[$($schema)].[$($tableName)]"


# Inserts new data into the table
$InsertQuery = "INSERT INTO [$($schema)].[$($tableName)] (FirstName, LastName, DisplayName, PostalCode, OfficePhone, MobilePhone)"
$NewClients = Import-Csv $PSScriptRoot\NewClientData.csv

Write-Host -ForegroundColor Cyan "[SQL]: Inserting Data"
foreach ($NewClient in $NewClients) {
    $Values = " VALUES ('$($NewClient.FirstName)',
     '$($NewClient.LastName)', 
     '$($NewClient.DisplayName)', 
     '$($NewClient.PostalCode)', 
     '$($NewClient.OfficePhone)', 
     '$($NewClient.MobilePhone)')"
     $query = $InsertQuery + $Values
     Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $query

}


# Read Data
Write-Host -ForegroundColor Cyan "[SQL]: Reading Data"
$selectQuery = "SELCT * FROM $($tableName)"
$Clients = Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $query
foreach ($Client in $Clients)  {
    Write-Host "Client Name: $($Client.FirstName) $($Client.LastName)"
    Write-Host "Display Name: $($Client.DisplayName)"
    Write-Host "Postal Code: $($Client.PostalCode)"
    Write-Host "Office Phone: $($Client.OfficePhone)"
    Write-Host "Mobile Phone: $($Client.MobilePhone)"
    Write-Host "-------------"
}
Write-Host -ForegroundColor Cyan "[SQL]: SQL Tasks Complete"
Invoke-Sqlcmd -Database ClientDB -ServerInstance .\SQLEXPRESS -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SqlResults.txt
