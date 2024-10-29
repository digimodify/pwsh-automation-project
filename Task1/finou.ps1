<#
Christian Etienne
011715153
05/21/2024
D411 Task 2
#>

Try {
    ### Create a database ###
    
    Write-Host -ForegroundColor Cyan "[SQL]: Starting SQL Server Tasks"
    # Import SQLServer Module
    if (Get-Module -Name sqlps) { Remove-Module sqlps }
    Import-Module -Name SqlsServer

    # Set a string variable for the name of the SQL Instance
    $sqlServerInstance = "SRV10-PRIMARY\SQLEXPRESS"

    # Set a string variable for the Database name
    $databaseName = 'MyDatabase-InvokeSqlCmd'

    # Create object reference to the sql server
    $sqlServerObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $sqlServerInstance

    # Create object reference to the database and check if it exists
    $databaseObject = Get-SqlDatabase -sqlServerInstance $sqlServerInstance -Name $databaseName -ErrorAction SilentlyContinue
    if($databaseObject) {
        Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) Database exists. Deleting..."
        
        # Kill running processes in database
        $sqlServerObject.KillAllProcesses($databaseName)

        # Set database to single user access mode
        $databaseObject.UserAccess = "Single"

        # Delete the database
        $databaseObject.Drop()
    }
    else {
        Write-Host -ForegroundColor Cyan "[SQL]: $($databaseName) Not Found..."
    }

    # Create the database object with the create method
    $databaseObject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Database -ArgumentList $sqlServerObject, $databaseName
    $databaseObject.Create()
    Write-Host -ForegroundColor Cyan "[SQL]: Database Created: [$($sqlServerInstance)].[$($databaseName)]"

    ### Create table ###

    # Invoke a sql command against the sql instance
    $schema = "dbo"
    $tableName = "Client_A_Contacts"
    Invoke-Sqlcmd -ServerInstance $sqlServerInstance -Database $databaseName -InputFile $PSScriptRoot\CreateTable_ClientA.sql

    # Update user on progress
    Write-Host -ForegroundColor Cyan "[SQL]: Table created: [$($sqlServerInstance)].[$($databaseName)].[$($schema)].[$($tableName)]"

    # Import records from the csv file and iterate over each; Insert each into table
    $InsertQuery = "INSERT INTO [$($schema)].[$($tableName)] (first_)name, last_name, city, county, zip, officePhone, mobilePhone)"
    $NewClients = Import-Csv $PSScriptRoot\NewClientData.csv

    # Loop through each record and insert into table
    Write-Host -ForegroundColor Cyan "[SQL]: Inserting Data..."
    foreach($client in $NewClients) {
        $Values = "VALUES ('$($Client.first_name)', `
                           '$($Client.last_name)', `
                           '$($Client.city)', `
                           '$($Client.county)', `
                           '$($Client.zip)', `
                           '$($Client.officePhone)'
                           '$($Client.mobilePhone)'
        )"
        $query = $InsertQuery +$Values
        Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstance -Query $query
    }
# Read data
Write-Host -ForegroundColor Cyan "[SQL]:Reading Data..."
$selectQuery = "SELECT * FROM $($tableName)"
$Clients = Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstance -Query $selectQuery
foreach($Client in $Clients) {
    Write-Host "Client Name: $($Client.first_name) $($Client.last_name)"
    Write-Host "Address: $($Client.county) County, City of $($Client.city). Zip - $($Client.zip)"
    Write-Host "Phone: Office $($Client.officePhone), Mobile $($Client.mobilePhone)"
    Write-Host "----------"
}
Write-Host -ForegroundColor Cyan "[SQL]: SQL Tasks Complete"

# Generate output file with the contact data from the table
Invoke-Sqlcmd -Database $databaseName -ServerInstance $sqlServerInstance -Query "SELECT * FROM dbo.NewClients" > $PSScriptRoot\SqlResults.txt

}
# Catch any errors
Catch {
    Write-Host -ForegroundColor Red "An exception has occured."
    Write-Host -ForegroundColor Red "$($PSItem.ToString())`n`n$($PSItem.ScriptStackTrace)"
}