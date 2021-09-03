# DM - 27.08.2021
# Script that extracts specific errors by pooling a list of CSV files

# List of CSV files to open (in the same folder as the Pwsh script)
$CSV_ORIG_path = ".\CSV_ORIG\"
$CSV_ERR_path = ".\CSV_ERR\"
$outputpath = "C:\PKI Service\DANIEL\JANUS\csv_error\allErrors.xlsx"
# List of files to process
$files = "Janus-17201.csv","Janus-17203.csv","Janus-17204.csv","Janus-17207.csv","Janus-17208.csv","Janus-17209.csv",
         "Janus-17211.csv","Janus-17213.csv","Janus-17215.csv","Janus-17216.csv","Janus-17218.csv","Janus-17220.csv",
         "Janus-17224.csv","Janus-17225.csv","Janus-17226.csv","Janus-17253.csv","Janus-17254.csv","Janus-17255.csv","Janus-17256.csv"
# List of errors to extract
$errors = "65","66","119","122","271","313"

$excel = New-Object -ComObject excel.application 
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$wksht= $workbook.Worksheets.Item(1) 
$wksht.Name = 'The name you choose'

$wksht.Cells.Item(1,2) = '65' 
$wksht.Cells.Item(1,3) = '66' 
$wksht.Cells.Item(1,4) = '119'
$wksht.Cells.Item(1,5) = '122' 
$wksht.Cells.Item(1,6) = '271' 
$wksht.Cells.Item(1,7) = '313'
# counter for instrument number
$i = 2
# counter for error number
$j = 2

# The loop for each file in the list
foreach ($path in $files)
{
    Write-Host $path
    # write instrument number in the Excel file
    # create a new name for the file
    $name = $path.Replace('Janus-','')
    $name_new = $name.Replace('.csv',' ')

    $wksht.Cells.Item($i,1) = $name_new


    # import the CSV original files (generated from the DB)
    $CSV_file1 = $CSV_ORIG_path+$path
    $A = Import-Csv $CSV_file1 -Delimiter ','

    # ... and for each error
    foreach ($error in $errors)
    {
        #TODO - compile also some statistics based on the sum of extracted errors

        # count the number of specific errors
        $ErrorSum = @($A | Where-Object -Property ErrorCode -Like $error).Count
        "Number of errors {$error}: $ErrorSum"
        
        # save the value in the Excel file
        $wksht.Cells.Item($i,$j) = $ErrorSum

        # create a new name for the file
        $name1 = $path.Replace('Janus-','')
        $file_new = $name1.Replace('.csv',' ')

        # Save the result for each error in its own file
        $CSV_file2 = $CSV_ERR_path+$file_new+"_Error_"+$error+".csv"
        $A | Where-Object -Property ErrorCode -Like $error | select ErrorCode, ErrorText, TestDatetime | Export-Csv -Path $CSV_file2 -Delimiter ';' -NoTypeInformation

        $j++
    }

    $j = 2
    $i++

}


$workbook.SaveAs($outputpath) 
$excel.Quit()
