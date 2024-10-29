<#
Christian Etienne
011715153
05/21/2024
D411 Task 2
#>

Try {

    ### Create database

    Write-Host -ForegroundColor Cyan "[SQL]: Starting SQL Server Tasks..."
    # Import SQL Server Module
    if (Get-Module -Name sqlps) { Remove-Module sqlps }

    # I think this is what I was missing on the first attempt
    Import-Module -Name SqlServer

    # Set string variable for the SQL instance name
    $sqlServerInstance = "SRV19-PRIMARY\SQLEXPRESS"

    # Set string variable for the Database name
    $databaseName = "ClientDB"

    # Create object reference to the sql server
    $sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstance

    # Create onject reference to the database/ check if it exists
    $databaseObject = Get-SqlDatabase -ServerInstance $sqlServerInstance -Name $databaseName -ErrorAction SilentlyContinue
    if($databaseObject) {
        Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) database exists. Deleting..."

        # Kill running processes on database
        $sqlServerObject.KillAllProcesses($databaseName)

        # Set database to single user access mode to limit access during database recreation
        $databaseObject.UserAccess = "Single"

        # Delete the database 
        $databaseObject.Drop()
    }
    else {
        Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) database not found..."
    }

    # Create database object
    $databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName
    $databaseObject.Create()
    Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) database created: [$($sqlServerInstance)].[$($databaseName)]"

    ### Create table
    # invoke sql command against the instance
    $schema = "dbo"
    $tableName = "Client_A_Contacts"
    Invoke-Sqlcmd -ServerInstance $sqlServerInstance -Database $databaseName -InputFile $PSScriptRoot\CreateTable_ClientA.sql

    # Update user on progress
    Write-Host -ForegroundColor Cyan "[SQL]: $($tableName) table created: [$($sqlServerInstance)].[$($databaseName)].[$($schema)].[$($tableName)]"

    # Import records from csv file and insert each into table
    $InsertQuery = "INSERT INTO [$($schema)].[$($tableName)] (first_name, last_name, city, county, zip, officePhone, mobilePhone) "
    $NewClients = Import-Csv $PSScriptRoot\NewClientData.csv

    # Loop through records and insert into table
    Write-Host -ForegroundColor Cyan "[SQL]: Inserting data into $($tableName) table..."
    foreach($Client in $NewClients) {
        $Values = "VALUES ('$($Client.first_name)', `
                           '$($Client.last_name)', `
                           '$($Client.city)', `
                           '$($Client.county)', `
                           '$($Client.zip)', `
                           '$($Client.officePhone)', `
                           '$($Client.mobilePhone)')"
                           
        $query = $InsertQuery + $Values
        Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstance -Query $query
    }

Write-Host -ForegroundColor Cyan "[SQL]: SQL Tasks Complete."

# Generate output file
Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstance -Query 'SELECT * FROM dbo.Client_A_Contacts' > $PSScriptRoot\SqlResults.txt
Write-Host -ForegroundColor Cyan "[SQL]: Generating output file..."
}
Catch {
    # Catch any exceptions
    Write-Host -ForegroundColor Red "An exception has occured"
    Write-Host -ForegroundColor Red "$($PSItem.ToString())`n`n$($PSItem.ScriptStackTrace)"
}

Break