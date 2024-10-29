# Powershell Automation Project

## Task 1 - Automate commonly run processes

Script 1: User Prompt and Logging Script

This script provides users with a menu-driven interface to perform various system tasks and log information efficiently. It includes options for:
- Daily Log: Lists and appends entries to log files.
- Files: Lists and appends information for all files in the specified directory.
- CPU/Memory Usage: Displays real-time CPU and memory usage metrics.
- Processes: Shows currently active processes in a graphical interface.


Script 2: SQL Database Management Script

This script automates SQL Server database tasks, including:
- Database and Table Creation: Creates a new database and table if they don’t already exist.
- Data Import: Reads from a CSV file and inserts the data into the SQL table.
- Data Export: Retrieves data from the table and writes it to an output file for validation.
- Built-in error handling provides reliable performance, and sample data with test results is stored in the repository directory for reference.

<br>

## Task 2 - Automate AD and SQL Server tasks

Script 1: Active Directory (AD) Tasks

This script accomplishes the following:
- Checks for and imports the ActiveDirectory module.
- Defines and checks for an Organizational Unit (OU) called Finance.
- Deletes and recreates the OU as needed.
- Imports user data from a CSV file, creating users in the Finance OU.
- Generates an output file with details of the users in the Finance OU.

Script 2: SQL Server Tasks

This script includes:
- Checking for and importing the SqlServer module.
- Defining the SQL Server instance and database name.
- Checking for the existence of a database (ClientDB), dropping it if present, and then recreating it.
- Creating a table and inserting data from a CSV file into the table.
- Generating an output file with all data from the table.
