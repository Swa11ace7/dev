<# Author : Shane Wallace 
# Student Id: 010895068
#>

if (Get-Module -Name sqlps) { Remove-Module sqlps}
Import-Module -Name SqlServer

Write-Host -ForegroundColor Cyan "[SQL]: Staring SQL Tasks"

# Set a string variable equal to the name SQL instance 
$sqlServerInstanceName = "SRV19-PRIMARY\SQLEXPRESS"

#Sets the database name with this variable
$databaseName = "ClientDB-InvokeSqlcmd"

# Create an Object reference to the SQL Server 
$sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstanceName

# Create an object referennce to the Database and try to detect if it exists
$databaseObject = Get-SqlDatabase -ServerInstance $sqlServerInstanceName -Name $databaseName -ErrorAction SilentlyContinue
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


# Call the Create method on the database object to create it
$databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName
$databaseObject.Create()

# Update user on the progress
White-Host -ForegroundColor Cyan "[SQL]: Database Created:[$($sqlServerInstanceName)].[$($databaseName)]"

<# Create Table #>

# Invoke a SQL Command against the SQL Instance bt reading in the contents of the Client_A_Contacts.sql
$schema = "dbo"
$tableName = 'Client_A_Contacts'
Invoke-Sqlcmd -ServerInstance $sqlServerInstanceName -Database $databaseName -InputFile $PSScriptRoot\Client_A_Contact.sql

# Update user on the progress
Write-Host -ForegroundColor Cyan "[SQL]: Table Created:[$($sqlServerInstanceName)].[$($databaseName)].[$($schema)].[$($tableName)]"

# Import the rows from the CSV file and iterate over each one calling Invoke-Sqlcmd 
# and passing a well-formatted INSERT statement

$InsertQuery = "INSERT INTO [$($schema)].[$($tableName)] (First_Name, Last_Name, Display_Name, Postal_Code, Office_Phone, Mobile_Phone)"
$NewClients = Import-Csv -Path $PSScriptRoot\NewClientData.csv

Write-Host -ForegroundColor Cyan "[SQL]: Inserting Data"
foreach ($NewClient in $NewClients) {
    $Values = " VALUES ('$(NewClient.First_Name)', `
                        '$($NewClient.Last_Name)', `
                        '$($NewClient.Display_Name)', `
                        '$($NewClient.Postal_Code)', `
                        '$($NewClient.Office_Phone)', `
                        '$($NewClient.Mobile_Phone)')"
    $query = $InsertQuery + $Values
    Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $query
}

# Read Data
Write-Host -ForegroundColor Cyan "[SQL]: Reading Data"
$selectQuery = "SELECT * FROM $($tableName)"
$Clients = Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstanceName -Query $selectQuery
foreach ($Client in $Clients) {
    Write-Host -ForegroundColor Cyan "Client Name: $($Client.First_Name), $($Client.Last_Name)"
    Write-Host -ForegroundColor Cyan "Display Name: $($Client.Display_Name)"
    Write-Host -ForegroundColor Cyan "Postal Code: $($Client.Postal_Code)"
    Write-Host -ForegroundColor Cyan "Phone: $($Client.Office_Phone), $($Client.Mobile_Phone)"
    Write-Host  "----------------"
}

# Used to create a SqlResults file
Invoke-Sqlcmd -Database ClientDB -ServerInstance .\SQLEXPRESS -Query 'SELECT * FROM dbo.Client_A_Contacts' > .\SqlResults.txt
