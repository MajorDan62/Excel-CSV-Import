# Excel-CSV-Import

The code is a PowerShell script that involves reading data from an Excel file, performing operations based on the data, and includes additional features such as email functionality and launching another script named "_dnss.ps1" with specific parameters.

Here's a breakdown of the script's main components and features:

Functions:

Last_reboot(): Retrieves the last boot-up time of the Windows operating system.
find_col(): Searches for a specified column in the Excel sheet.
display_col(): Displays the columns of the Excel sheet.
Variables and Configuration:

$excelfile: Specifies the path and name of the Excel file to be processed.
$_row_: Specifies the starting row for processing data in the Excel sheet.
$genesis: Stores the current date and time when the script starts.
$smtpServer, $smtpPort, $fromAddress, $toAddress: Email configuration details.
$subject: Specifies the subject of the email.
Command-line Arguments:

The script accepts command-line arguments, such as -xls (to specify an Excel file), -csv (to specify a CSV file), -query (to specify a column to search for), -l (to launch the "_dnss.ps1" script), and -em (to send an email to the specified address).
Excel File Processing:

If the -xls argument is provided, the script opens the specified Excel file, reads data from the specified column, and processes the records.
CSV File Processing:

If the -csv argument is provided, the script opens the specified CSV file, retrieves the specified column, and processes the records.
Output and Reporting:

The script provides progress updates during processing and outputs the summary of processed records.
It calculates the script's execution time and displays it.
It generates a body for an email containing the number of processed records and script execution time.
Additional Functionality:

If the -em argument is provided, the script sends an email using the specified email configuration.
If the -l argument is provided, it launches the "_dnss.ps1" script with specific parameters.
Please note that the script's functionality and behavior may depend on other files or scripts referenced, such as "_dnss.ps1" and any associated modules or dependencies.

If you have any specific questions or need further assistance with this script, please let me know.
