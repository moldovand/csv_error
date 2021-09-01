# DM - 27.08.2021
# Script that extracts specific errors by pooling a list of CSV files

# List of CSV files to open (in the same folder as the Pwsh script)
$CSV_ORIG_path = ".\CSV_ORIG\"
$CSV_ERR_path = ".\CSV_ERR\"
$files = "Janus-17253.csv","Janus-17254.csv","Janus-17255.csv","Janus-17256.csv"

# List of errors to extract
# $errors = "65","66","119","122","271","313"
$errors = "66"

# The loop for each file in the list
foreach ($path in $files)
{
    # import the CSV original files (generated from the DB)
    $CSV_file1 = $CSV_ORIG_path+$path
    $A = Import-Csv $CSV_file1 -Delimiter ';'

    # ... and for each error
    foreach ($error in $errors)
    {
        #TODO - compile also some statistics based on the sum of extracted errors

        # count the number of specific errors
        $ErrorSum = @($A | Where-Object -Property ErrorCode -Like $error).Count
        "Number of errors {$error}: $ErrorSum"

        # Save the result for each error in its own file
        $CSV_file2 = $CSV_ERR_path+$path+"_Error_"+$error+".csv"
        $A | Where-Object -Property ErrorCode -Like $error | select ErrorCode, ErrorText, TestDatetime | Export-Csv -Path $CSV_file2 -Delimiter ';' -NoTypeInformation
    }

}
