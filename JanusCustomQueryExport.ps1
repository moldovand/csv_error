<# 
Script to iterate over all .mdb files in $localdbpath, open an MS Access Object
and export $QueryName to a combined Csv file. 
#>

$hostname = $env:COMPUTERNAME    # TODO Make sure environmental variable matches machine name

$localdbpath = 'C:\Packard\Janus\database\'
$serverpath = '\\10.90.1.9\Janus_PE\PerkinElmer\databases\' + $hostname ## TODO Make sure PC-Name $hostname matches Janus

# for testing 
#$localdbpath = 'C:\Users\TSE\JanusDbExport\databases'
#$serverpath = 'C:\Users\TSE\JanusDbExport\server\PE_admin\PerkinElmer\databases\' + $hostname

#Custom query corresponds to "AllOperationsErrorsQuery" from Multiprobe.mdb 
# here-string @"..."@ for multiline string

$query = @"
         SELECT TestTbl.TestName,
               TestTbl.TestDateTime,
               OperationErrorTbl.TestId,
               OperationErrorTbl.ErrorCode,
               ProcedureTbl.ProcedureId,
               ProcedureTbl.ProcedureName,
               OperationErrorTbl.UserResponse,
               ResponseTbl.[Response Code],
               OperationErrorTbl.ElapsedTime,
               OperationErrorTbl.ErrorText,
               ResponseTbl.[Response Text],
               ApplicationTbl.UserName,
               `"$hostname`" as JanusMachineName
             FROM OperationErrorTbl,
                  ResponseTbl,
                  ProcedureTbl,
                  ApplicationTbl,
                  TestTbl
WHERE (((ProcedureTbl.ProcedureId)=[OperationErrorTbl].[ProcedureId]) 
      AND ((OperationErrorTbl.[ErrorCode])<>0) 
      AND ((TestTbl.TestId)=[OperationErrorTbl].[TestId]) 
      AND ((TestTbl.ApplicationId)=[ApplicationTbl].[ApplicationId]) 
      AND ((ProcedureTbl.TestId)=[OperationErrorTbl].[TestId]) 
      AND ((OperationErrorTbl.[UserResponse])=[Response Code]));
"@




function ExportQuery {
    param( [string]$query,
           [string]$dbpath
            )
    
        $connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=`"$dbpath`""
        $conn = new-object System.Data.OleDb.OleDbConnection($connString)
        $conn.open()

        $cmd = new-object System.Data.OleDb.OleDbCommand($query,$conn) 

        $da = new-object System.Data.OleDb.OleDbDataAdapter($cmd) 

        $dt = new-object System.Data.dataTable 
        [void]$da.fill($dt)

        $conn.close()

        $dt | export-csv $outfile -NoTypeInformation -Append
   
}


$mdbfiles = Get-ChildItem ( $localdbpath + '\*.mdb')

foreach ($mdbfile in $mdbfiles) {

    if ((Get-item $mdbfile).Length/1MB -gt 10) {              # SKIP new files
        ExportQuery -dbpath $mdbfile -query $query                       # run our function
    
    }

}

# move generated file to Server
New-item -Path $serverpath -ItemType Directory -Force
Move-Item $outfile -Destination $serverpath -Force