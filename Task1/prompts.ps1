<#
Name: Christian Etienne 
ID: 011755153
Date: 05/16/2024
D411 Scripting and Utomation Applications Task 1
#>

Function Get-UserInput {
    <#
    .SYNOPSIS 
        Display a list of choices to the user
    .DESCRIPTION 
        This function holds the list of selections a user can choose to perform the corresponding action.
    #>

    # Create a StringBuilder object to construct the list of options
    $Options = New-Object -TypeName System.Text.StringBuilder
    [void]$Options.AppendLine(" ")
    [void]$Options.AppendLine("Prompts")
    [void]$Options.AppendLine("--------------------")
    [void]$Options.AppendLine("1.) Daily Log")
    [void]$Options.AppendLine("2.) Files")
    [void]$Options.AppendLine("3.) CPU and Memory Usage")
    [void]$Options.AppendLine("4.) Processes")
    [void]$Options.AppendLine("5.) Exit")
    [void]$Options.AppendLine("--------------------")
    
    # Display the options and prompt the user for input
    Write-Host -ForegroundColor White $Options.ToString()
    return $(Write-Host -ForegroundColor White "Please make a selection (1 - 5): "; Read-Host)
}
# Variable that will hold the user selection
$UserInput = 1

# Error Handling: Try/Catch for errors that occur during the script execution
try {
    # Loop until user exits
    while ($UserInput -ne 5) 
    {
        # Prompt user for selection and run corresponding code
        $UserInput = Get-UserInput
        # Switch statement will handle selected options
        switch ($UserInput) 
        {
            # User selects option 1:  Prints all log files in root folder and also appends list of files and current date to text file.
            1 {
                Write-Host "Daily Log"
                Write-Host "---------------------"
                Write-Host "Appending results to DailyLog.txt..."
                Get-Date | Out-File -FilePath "$PSScriptRoot\DailyLog.txt" -Append
                Get-ChildItem -Path $PSScriptRoot -Filter *.log | Out-File -FilePath "$PSScriptRoot\DailyLog.txt" -Append                
            }
            # Option 2: Prints list of all files in root folder and also appends list of files to contents file in tabular format.
            2 {
                Write-Host "Files"
                Write-Host "---------------------"
                Write-Host "Appending results to D411Contents.txt..."
                Get-ChildItem -Path $PSScriptRoot | Sort-Object Name| Format-Table -AutoSize -Wrap | Out-File -FilePath "$PSScriptRoot\D411contents.txt" -Append  
            }
            # Option 3: Prints list of CPU and Memory Ussage. Shows 4 samples in 3 second intervals.
            3 {
                Write-Host "CPU and Memory"
                Write-Host "---------------------"
                Write-Host "Displaying CPU Usage (%) and Memory available in MB..."
                Write-Host " "
                $CounterList = $("\Processor(_Total)\% Processor Time", "\Memory\Available MBytes")
                Get-Counter -Counter $CounterList -SampleInterval 5 -MaxSamples 4
                Write-Host "CPU and Memory sample complete."
             
            } 
            # Option 4: Prints list of processes sorted by Virtual Memory usage in descending order. Will list ID, Name, and Virtual Memory in Out-Grid view (another window).
            4 {
                Write-Host "Processes"
                Write-Host "---------------------"
                Write-Host "Launching graphical representation of current running processes..."
                Write-Host " "
                Get-Process | Select-Object ID, Name, VM | Sort-Object VM | Out-GridView
            }
            # Option 5: Exits the script
            5 {
                Write-Host "Exiting gracefully..."
                Break
            }
            # This will show when an invalid selection is entered.
            default { Write-Host "Invalid selection. Please select 1 -5." }
        }
    }
}
# Prints a message for the specific event of the suystem running out of memory/ other unhandled exception.
catch [System.OutOfMemoryException]
{
    Write-Host -ForegroundColor Red "A system out of memory exception has occured:"
    # String representation of the exception and it's call stack
    Write-Host -ForegroundColor Red "$($PSItem.ToString())`n`n$($PSItem.ScriptStackTrace)"
} 
catch {
    Write-Host -ForegroundColor Red "An unhandled exception has occured:"
    # String representation of the exception and it's call stack
    Write-Host -ForegroundColor Red "$($PSItem.ToString())`n`n$($PSItem.ScriptStackTrace)"
}
finally {
    # Close any open resources if applicable
}
