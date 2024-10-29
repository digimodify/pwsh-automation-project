<#
Christian Etienne
011715153
05/21/2024
D411 Task 2
#>

Try 
{
    Write-Host -ForegroundColor Cyan "[AD]: Starting Active Directory Tasks..."
    # Check if Active Directory module is loaded, import.
    if (Get-Module -Name ActiveDirectory) { Remove-Module ActiveDirectory }
    Import-Module -Name ActiveDirectory

    # Assign variables for creating the Finance OU
    $ADRoot = (Get-ADDomain).DistinguishedName
    $DnsRoot = (Get-ADDomain).$DnsRoot
    $OUCanonicalName = "Finance"
    $OUDisplayName = "Finance"
    $ADPath = "OU=$($OUCanonicalName),$($ADRoot)"

    # Check if the Finance OU exists in Active Directory
    if (-not([ADSI]::Exists("LDAP://$($ADPath)"))) {
        Write-Host "[AD]: $($OUCanonicalName) not found..."
    } 
    else {
        # If the OU is located, delete so that it can be recreated.
        Write-Host -ForegroundColor Cyan "[AD]: $($OUCanonicalName) already exists. Deleting..."
        Remove-ADOrganizationalUnit -Identity $ADPath -Recursive -Confirm:$false
        Write-Host -ForegroundColor Cyan "[AD]: $($OUCanonicalName) has been deleted..."       
    }

    # Create the OU
    Write-Host -ForegroundColor Cyan "[AD]: Creating the Finance Organizational Unit..."
    New-ADOrganizationalUnit -Path $ADRoot -Name $OUCanonicalName -DisplayName $OUDisplayName -ProtectedFromAccidentalDeletion $false

    # Import records from csv file
    $newADUsers = Import-csv $PSScriptRoot\financePersonnel.csv

    # Counter to keep track of added users/progress
    $newUserCount = $newADUsers.Count
    $count = 1

    # Iterate through records and add to OU
    foreach($User in $newADUsers) {
        $firstName = $User.First_Name
        $lastName = $User.Last_Name
        $displayName = $firstName + " " + $lastName
        $postalCode = $User.PostalCode
        $officePhone = $User.OfficePhone
        $mobilePhone = $User.MobilePhone 

        # Display status of current users added to the OU
        $status = "[AD]: Adding user: $($displayName) ($($count) of $($newUserCount))"
        Write-Progress -Activity 'D411 Task 2 - Restore Finance OU' -Status $status -PercentComplete (($count/$newUserCount) * 100)

        # Create the users in AD with New-ADUser cmdlet with necessary columns
        New-ADuser -GivenName $firstName `
                   -Surname $lastName `
                   -Name $displayName `
                   -PostalCode $postalCode `
                   -OfficePhone $officePhone `
                   -MobilePhone $mobilePhone `
                   -Path $ADPath
        # Increment counter
        $count++ 
    }
    Write-Host -ForegroundColor Cyan "[AD]: Active Directory Tasks Complete."

# Generate output file
Get-ADUser -Filter * -SearchBase "ou=Finance,dc=consultingfirm,dc=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\ADResults.txt
Write-Host -ForegroundColor Cyan "[AD]: Generating output file..."
}
catch {

# Catch any exceptions
Write-Host -ForegroundColor Red "An exception has occured"
Write-Host -ForegroundColor Red "$($PSItem.ToString())`n`n$($PSItem.ScriptStackTrace)"
}

Break