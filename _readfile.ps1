#############################	
# GET DATA FROM EXCEL FILE  #
# VERSION 1_230511          #
# DATED 11th-05-2023        #
# DESIGNER MA JORDAN        #
#############################
# VERSION 1_20230522        #
# ADDED EMAiL FUNCTION      #
# ADDED LAUNCH DNSS CODE    #
#############################

function Last_reboot(){return  (gcim Win32_OperatingSystem).LastBootUpTime} 
function find_col{for ($counter=$_row_; $counter -le $colmax; $counter++){if (($sheet.Cells.Item(1,$counter).text) -eq $query){$_col=$counter}}return $_col}
function display_col
{
    for ($counter=1; $counter -le ($colmax -1 ); $counter++)
    {
        $COLproperties  =   [ordered]@{ COL  = $sheet.Cells.Item(1,$counter).text}
        $COL_obj        =   New-Object psobject -Property $colproperties                                    #create collection of data
        $COLproperties +=   $COL_obj  
        # $_columnID_[$counter] = $sheet.Cells.Item(1,$counter).text
        # $_columnID_[$counter]
    }
}

clear
$cr             =   "`r`n"
$_curdir        =   get-location 
$excelfile      =   "$_curdir\data2.xlsx"
$_row_          =   2
$genesis        =   Get-Date
$_date 			= 	Get-Date -format yyyyMMdd_HH_mm
$_lr_	        =	last_reboot
$_input_xls_    =   $false
$_input_csv_    =   $false
$_launch_dnss_  =   $false
$query          =   "Host (Impacted)"
# Define the email details
$smtpServer = "imrpool.fs.fujitsu.com"
$smtpPort = 25
$fromAddress = "do.not.reply@powershell.fujitsu.com"
$toAddress = "mike.jordan@fujitsu.com"
$smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtpClient.UseDefaultCredentials = $true
$subject = "Test Generated @ $_date"

if ($args.count -gt 0)
{
	for ($i=0;$i -lt $args.count;$i++)
	{
		if ($args[$i].ToUpper() -eq "-xls")     {$excelfile	    =   "$_curdir\"+$args[($i+1)];$_input_xls_=$true}
		if ($args[$i].ToUpper() -eq "-csv")     {$csvfile	    =   "$_curdir\"+$args[($i+1)];$_input_csv_=$true}
		if ($args[$i].ToUpper() -eq "-query")   {$query	        =   $args[($i+1)]}
		if ($args[$i].ToUpper() -eq "-l")       {$_launch_dnss_=$true}
		if ($args[$i].ToUpper() -eq "-em")       {$_launch_email_=$true}
    }
}
else {$_display=$True} 

if ($_display)
{
	write-host "===================================="
	write-host "| Name    : _readfile              |"
	write-host "| Version : 20230511_v1            |"
	write-host "| Dated   : 12th March 2023        |"
	write-host "| Author  : " -nonewline
	write-host -f green "Mike Jordan" -nonewline
	write-host -f white "            |"
	write-host "| Last Reboot $_lr_  |"
	write-host "===================================="

	write-host " -csv       Select CSV input source"
	write-host " -query     Search for this column "
	write-host " -xls       Select EXCEL input source"
	write-host " -l         Launch DNSS Script"
	write-host " -em        Send Eamil to Admin"
	write-host
	exit
}

IF ($_input_xls_)
{
    write-host -b green -f white "--  PROCESS EXCEL SPREADSHEETS  --"
    write-host  "Last Rebooted  $_lr_ "
    $excelfile      =   $excelfile.replace('\.\','\')
    if (Test-Path $excelfile -PathType Leaf){write-host "Processing file : "$excelfile -f green}else{write-host -f white -b red  " ** ERROR: NO DATAFILE LOCATED !! ** $cr    Please retype a valid filename   ";exit}

    $objExcel       =   New-Object -ComObject Excel.Application         #create excel object
    $workbook       =   $objExcel.Workbooks.Open($excelfile)		    #Open excel workbook
    $sheet          =   $workbook.Worksheets.Item(1)
    $rowmax         =   ($sheet.UsedRange.Rows).count			        #Count #Rows in the worksheet
    $colmax         =   ($sheet.UsedRange.Columns).count                #Count #Cols in the worksheet
    $wsmax          =   ($workbook.Worksheets).count                    #Count #tabs in the worksheet   
    $RulesfromEXCEL =   New-Object string[] $rowmax                     #Create arrays in the worksheet
    $properties     =   New-Object string[] $rowmax                     #Create arrays in the worksheet
    $_col           =   find_col
    $_columnID_     =   New-Object string[] $colmax   
    $COLproperties  =   New-Object string[] $colmax 

    write-host "File Summary    : #rows=$rowmax, #cols=$colmax, ColID=$_col,  #Sheets=$wsmax" -f green
    # $x=display_cols
    for ($counter=$_row_; $counter -le $rowmax; $counter++)
        {
            $_PercentFree_  =  (( $counter *100 ) / $rowmax )
            $_PercentFree_  =   [math]::round($_PercentFree_ ,2)
            Write-Progress "Processing Complete"  -perc $_PercentFree_

            if ($sheet.Cells.Item($counter,1).text)
            { 
                $properties = [ordered]@{ DestinationIP  = $sheet.Cells.Item($counter,$_col).text}
                $NET_obj = New-Object psobject -Property $properties                                    #create collection of data
                $RulesfromEXCEL +=   $NET_obj                                                           #compile the collection  into a string
            } 
        }
    $workbook.close($True)

    $_formoutput_   =   $RulesfromEXCEL.DestinationIP  | Sort-Object | Get-Unique 
    $num           =   $_formoutput_.count 
    $_formoutput_ | clip

    write-host -f white -b green "Summary : $num records copied to Clipboard"
}
elseif ($_input_csv_)
{
    write-host -b green -f white "-- PROCESSING CSV FILE --"
    write-host -b blue -f white "Seaching for '$query'"
    write-host  "Last Rebooted  $_lr_ "
    if (Test-Path $csvfile -PathType Leaf){write-host "Processing file : "$csvfile -f green}else{write-host -f white -b red  " ** ERROR: NO DATAFILE LOCATED !! ** $cr    Please retype a valid filename   ";exit}

    $csv    =   Import-Csv -Path $csvfile | select -ExpandProperty $query| Sort-Object | Get-Unique 
    $num    =   ($csv|measure).count
    $csv|clip
    write-host -f white -b green "Summary : $num records copied to Clipboard"
}

$exodus=Get-Date
$delta=$exodus-$genesis;$h="{0:D2}" -f [int]$delta.hours;$m="{0:D2}" -f [int]$delta.minutes;$s="{0:D2}" -f [int]$delta.seconds
write-host -b blue -f white "[_GDFESS] Script Active for"$h"h "$m"m "$s"s ";
write-host

$body = "Powershell Script '_readfile.ps1' processed " + $num +" records `r`n" +   "Script active for  " + $delta + "`r`n" 
# Create a new email message
$mailMessage = New-Object System.Net.Mail.MailMessage($fromAddress, $toAddress, $subject, $body)

# Send the email
if ($_launch_email_){$smtpClient.Send($mailMessage)}

if ($_launch_dnss_){$param1 = "-mem";$param2 = "-r";Invoke-Expression -Command "C:\DNSS\_dnss.ps1 -Param1 $param1 -Param2 $param2"}   

