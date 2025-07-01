<# Author : Shane Wallace 
# Student Id: 010895068
#>

Try
{
    <# Create Database #>
    Write-Host -ForegroundColor Cyan "[SQL]: Staring SQL Tasks"
    
    #Import the SqlServer module
    if (Get-Module -Name sqlps) { Remove-Module sqlps }
    Import-Module -Name SqlServer

    # Set a string variable equal to the name SQL instance
    $sqlServerInstanceName = "SRV19-PRIMARY\SQLEXPRESS"

    # Sets the database name with this variable
    $databaseName = "ClientDB"

    # Create Connection to the SQL Server
    $connectionString = "Data Source=$($sqlServerInstanceName);Initial Catalog=master;Integrated Security=True;"
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $sqlConnection.ConnectionString = $connectionString
    $sqlConnection.Open()

    # Create Query to check if the database already exists
    $detectQuery = "SELECT database_id FROM sys.databases WHERE name = '$($databaseName)'"

    #Create Command to execute the query
    $sqlCommand = New-Object System.Data.SqlClient.SqlCommand $detectQuery, $sqlConnection
    $sqlReader = $sqlCommand.ExecuteReader()

    if($sqlReader.Read())
    {
        $sqlReader.Close()
        Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) database has been found and removed"
        # Create query command to drop the database
        $dropQuery = "ALTER DATABASE [$($databaseName)] SET SINGLE_USER WITH ROLLBACK IMMEDIATE DROP DATABASE [$($databaseName)]"
        $sqlCommand.CommandText = $dropQuery
        $sqlCommand.ExecuteNonQuery() | Out-Null
    }
    $sqlReader.Close()
    $sqlReader.Dispose()

    # Create the database
    $createQuery = "CREATE DATABASE [$($databaseName)]" 
    $sqlCommand.CommandText = $createQuery
    $sqlCommand.ExecuteNonQuery() | Out-Null
    Write-Host -ForegroundColor Cyan "[SQL]: Database Created:[$($sqlServerInstanceName)].[$($databaseName)]"

    # Create Table
    $schema = "dbo"
    $tableName = 'Client_A_Contacts'

    $createTableQuery = "
        Use [$($databaseName)]
        CREATE TABLE Client_A_Contacts
        (
            first_name varchar(40),
            last_name varchar(40),
            city varchar(40),
            county varchar(40),
            zip_code varchar(40),
            office_phone varchar(40),
            mobile_phone varchar(40),
        )"
    
    $sqlCommand.CommandText = $createTableQuery
    $sqlCommand.ExecuteNonQuery() | Out-Null
    Write-Host -ForegroundColor Cyan "[SQL]: Table Created:[$($sqlServerInstanceName)].[$($databaseName)].[$($schema)].[$($tableName)]"

    # Get Data from CSV file
    $NewClients = Import-Csv -Path "$PSScriptRoot\NewClientData.csv"

    # Insert Data into SQL
    Write-Host -ForegroundColor Cyan "[SQL]: Inserting Data"
    $insert = "INSERT INTO [$($tableName)] (first_name, last_name, city, county, zip_code, office_phone, mobile_phone)"

    foreach($NewClient in $NewClients)
    {
        $values = " VALUES ('$($NewClient.first_name)',
                            '$($NewClient.last_name)', 
                            '$($NewClient.city)', 
                            '$($NewClient.county)', 
                            '$($NewClient.zip_code)', 
                            '$($NewClient.office_phone)', 
                            '$($NewClient.mobile_phone)')"
        $sqlCommand.CommandText = $insert + $values
        $sqlCommand.ExecuteNonQuery() | Out-Null
    }

    # Read Data
    Write-Host -ForegroundColor Cyan "[SQL]: Reading Data"
    $dataTable = New-Object System.Data.DataTable
    $sqldataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $sqlCommand.CommandText = "SELECT * FROM $($tableName)"
    $sqldataAdapter.SelectCommand = $sqlCommand
    $sqldataAdapter.Fill($dataTable)
    
    foreach ($row in $dataTable.Rows)
    {
        Write-Host $row["first_name"]
        Write-Host "Client Name: $($row["first_name"]), $($row["last_name"])"
        Write-Host "Address: $($row["county"]) County, City of $($row["city"]), Zip - $($row["zip"])"
        Write-Host "Phone: Office $($row["office_phone"]), Mobile $($row["mobile_phone"])"
        Write-Host "-----------"
    }
    Write-Host -ForegroundColor Cyan "[SQL]: SQL Tasks Complete"

}
# Catch any Errors
Catch{
    Write-Host -ForegroundColor Red "An EXCEPTION occurred: $($_.Exception.Message)"
    Write-Host -ForegroundColor Red "Stack Trace: $($_.Exception.StackTrace)"
}
Finally{
    $sqlConnection.Close()  
}
