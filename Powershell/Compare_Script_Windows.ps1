<# 

READ ME
    script path is generated based on the location where the script is placed
    Both the input files & plink.exe must be placed in the same path as the script
    Update the serverName in getFromUnix function

#>

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

$Global:linesMismatchTable = @()

#region FUNCTIONS

#Function to create HTML file
function createHTMLFile($dataTable, $outFilePath ,$outFileName) {

    $dataTable | Format-Table

    #HTML Output table - Timestamp
    $reportCreationTime = " Report created: " + (Get-Date)

    #HTML table component
    $html = $dataTable | ConvertTo-Html -As Table -Title 'dataTable' -PreContent "Executed by $env:username" -PostContent $reportCreationTime

    #attribute to adjust the HTML Table border
    $html = $html -replace '<table>' , '<table border=1>'

    #To remove the issue with Convertto-HTML cmdlet changing link structure
    $html  = $html -replace '&lt;', '<'; $html = $html -replace '&quot;' , '"'; $html = $html -replace '&gt;', '>'

    #name for the HTML file
    $outFile = $outFilePath + '\' + $outFileName+ '.html'

    #save the HTML file
    $html | Out-File $outFile

    #To open the HTML file : Invoke-Item $outFile

}

#Primary function used here
function captureMismatch ($file1, $file2)
{
    
    #extract headers from the csv files
    $headers1 = $file1[0].psobject.Properties.Name

    $headers2 = $file2[0].psobject.Properties.Name

    $differences = Compare-Object -ReferenceObject $headers1 -DifferenceObject $headers2

    #Array to store compare failed record/fields
    $fieldsMismatch = @()

    #Check number of Records and header between both files and record type
    #if(($file1.Length -eq $file2.Length) -and ($headers1.ToString() -eq $headers2.ToString()) -and ($headers1[0] -eq $headers2[0])){
    #edited 09/28/2023
    if(($file1.Length -eq $file2.Length) -and ($headers1.length -eq $headers2.length) -and ($differences.InputObject -eq $null)){
        
        #Variable to capture number of compare failures for curret record type
        $failedLine = 0

        $totalHeaders = ($headers1 + $headers2 | Sort-Object -Unique)

        for($i=0; $i -lt $file1.length; $i++){
            
            #Check each line matches between 2 files
            if($file1[$i] -eq $file2[$i]){

                #Success record - no updates required
                Write-host "Lines match"

            }
            else{ 
                #compare failed for the line -add count-
                $failedLine+=1
                
                #compare line failed- identify fields failing compare
                for($j=0 ; $j -lt $totalHeaders.Length ; $j++){

                    if($file1[$i].($totalHeaders[$j].ToString()) -eq $file2[$i].($totalHeaders[$j].ToString())){
                        
                        <#Field matches - no need to add error#>

                    }
                    elseif($totalHeaders[$j] -notmatch 'Run Time'){
                    
                        #field mismatch - create a error row
                        $failRow = [PSCustomObject]@{
                            RecordType = $headers1[0]
                            $headers1[1] = $file1[$i].($headers1[1].ToString())
                            $headers1[2] = $file1[$i].($headers1[2].ToString())
                            $headers1[3] = $file1[$i].($headers1[3].ToString())
                            Mismatch_header = $totalHeaders[$j]
                            Test_File_Value = $file1[$i].($totalHeaders[$j].ToString())
                            Prod_File_Value = $file2[$i].($totalHeaders[$j].ToString())
                        }

                        $fieldsMismatch+= $failRow

                        $fieldsMismatch | format-table

                    }
                }

            }

        }

    }
    Elseif($headers1.length -ne $headers2.length){

        <#Place Holder to update mismatch between record type#>
        #added 09/28/2023
        Write-Host 'Header Count Mismatch'

        $totalHeaders = ($headers1 + $headers2 | Sort-Object -Unique)

        for($i=0; $i -lt $file1.length; $i++){
            
            #Check each line matches between 2 files
            if($file1[$i] -eq $file2[$i]){

                #Success record - no updates required
                Write-host "Lines match"

            }
            else{ 

                #compare failed for the line -add count-
                $failedLine+=1
                
                #compare line failed- identify fields failing compare
                for($j=0 ; $j -lt $totalHeaders.Length ; $j++){

                    if($file1[$i].($totalHeaders[$j].ToString()) -eq $file2[$i].($totalHeaders[$j].ToString())){

                        <#Field matches - no need to add error#>

                    }
                    elseif($totalHeaders[$j] -notmatch 'Run Time'){
                    
                        #field mismatch - create a error row
                        $failRow = [PSCustomObject]@{
                            RecordType = $headers1[0]
                            $headers1[1] = $file1[$i].($headers1[1].ToString())
                            $headers1[2] = $file1[$i].($headers1[2].ToString())
                            $headers1[3] = $file1[$i].($headers1[3].ToString())
                            Mismatch_header = $totalHeaders[$j]
                            Test_File_Value = $file1[$i].($totalHeaders[$j].ToString())
                            Prod_File_Value = $file2[$i].($totalHeaders[$j].ToString())
                        }

                        $fieldsMismatch+= $failRow

                        $fieldsMismatch | format-table

                    }
                }

            }

        }

    }
    Elseif($file1.Length -ne $file2.Length){

        <#Place Holder to update mismatch between record count#>
        #added 09/28/2023
        Write-Host 'Number of lines between files are not equal'

    }
    Elseif($differences.InputObject -ne $null){

        <#Place Holder to update mismatch headers between files#>
        #added 09/28/2023
        Write-Host 'Headers are not same'

        $totalHeaders = ($headers1 + $headers2 | Sort-Object -Unique)

        for($i=0; $i -lt $file1.length; $i++){
            
            #Check each line matches between 2 files
            if($file1[$i] -eq $file2[$i]){

                #Success record - no updates required
                Write-host "Lines match"

            }
            else{ 
                #compare failed for the line -add count-
                $failedLine+=1
                
                #compare line failed- identify fields failing compare
                for($j=0 ; $j -lt $totalHeaders.Length ; $j++){

                    if($file1[$i].($totalHeaders[$j].ToString()) -eq $file2[$i].($totalHeaders[$j].ToString())){
                        
                        <#Field matches - no need to add error#>
                        
                    }
                    elseif($totalHeaders[$j] -notmatch 'Run Time'){
                    
                        #field mismatch - create a error row
                        $failRow = [PSCustomObject]@{
                            RecordType = $headers1[0]
                            $headers1[1] = $file1[$i].($headers1[1].ToString())
                            $headers1[2] = $file1[$i].($headers1[2].ToString())
                            $headers1[3] = $file1[$i].($headers1[3].ToString())
                            Mismatch_header = $totalHeaders[$j]
                            Test_File_Value = $file1[$i].($totalHeaders[$j].ToString())
                            Prod_File_Value = $file2[$i].($totalHeaders[$j].ToString())
                        }

                        $fieldsMismatch+= $failRow

                        $fieldsMismatch | format-table
                    }

                }

            }

        }

    }
    
    if($fieldsMismatch.length -gt 0){

        createHTMLFile -dataTable $fieldsMismatch -outFilePath $htmlPath -outFileName $headers1[0]

    }

    return $fieldsMismatch
}

#endregion FUNCTIONS

#region script flow - sequence of steps

#Date time param for folder creation
$folderName = get-date -Format MMddyyyy_HHmm

$filepath = $scriptPath + '\OMNIRecTypes.csv'

$filenames = Import-Csv -Path $filepath

$fileCount = ($filenames | Measure-Object).Count

$htmlPath = $scriptPath + '\' + $folderName

#check if folder exist, else add folder
if( Test-Path -Path $htmlPath){}
else{ New-Item -ItemType Directory -Path $htmlPath}

#initialize array variables
$dashboard = @()

$recordMismatchTable = @()

##$Global:linesMismatchTable = @()

$recordType = @()

$linesMismatchTable = @()

##Set-Variable -Name linesMismatchTable -Scope 1

$fileCount = ($filenames | Measure-Object).Count

for($cnt=0; $cnt -lt $fileCount; $cnt++){
#for($cnt=0; $cnt -lt 1; $cnt++){

    Write-Host "$cnt"

    $testfilename = $scriptPath + '\' + $filenames[$cnt].TestFile

    $prodfilename = $scriptPath + '\' + $filenames[$cnt].ProdFile

    $hash1 = Get-FileHash -Path $testFileName

    $hash2 = Get-FileHash -Path $prodfilename

    if($hash1.Hash -eq $hash2.Hash){

        #Files are equal - No need to report

    }
    else{

        #Read the saved files into variables for comparison
        $testFile = Import-csv -Path $testfilename -Delimiter '|'

        $prodFile = Import-csv -Path $prodfilename -Delimiter '|'

        #Call function to extract the mismatches
        $mismatchedData = captureMismatch -file1 $testFile -file2 $prodFile


        if($mismatchedData.length -gt 0){

            $linesMismatched = [PSCustomObject]@{
                RecordType = $mismatchedData.RecordType[0]
                NumberOfFieldMismatches = $mismatchedData.length
                LinkToResult  = '<a href = " '+ $htmlPath +'\'+ $mismatchedData.RecordType[0]+'.html">Click here</a>'
            }

            $linesMismatchTable += $linesMismatched
        }

        Write-Host "Line mismatch table"
        $linesMismatchTable | format-table

        #create record mismatch html file
        createHTMLFile -dataTable $linesMismatchTable -outFilePath $htmlPath -outFileName 'linesMismatchTable'

    }
    
    clear-variable -Name mismatchedData

}

#calculate  the total number of mismatched rows/records across all record types
$totalRecordMismatch = 0

if($linesMismatchTable.length -gt 0){

    foreach ($line in $linesMismatchTable) { $totalRecordMismatch += $line.NumberOfFieldMismatches }

    $autoregfile = $scriptPath + '/INPUT.AUTOREG.txt'

    $File = $scriptPath + '/NGINT.OUTPUT.DATA.AUTOREG.txt'

    $inputAutoReg = Get-Content -Path $autoregfile

    $inputAutoRegLength = ($inputAutoReg | Measure-Object).Count

    $testOrProdFile = Get-Content -Path $File

    $testOrProdFileLength = ($testOrProdFile | Measure-Object).Count

    #Create a Dashboard Object table
    $dashboard = [PSCustomObject]@{
        #Report_Digit = $fset.FSET[$fcnt]
        Total_Test_Scenarios_Validated = $inputAutoRegLength
        Total_Input_records_Validated = $testOrProdFileLength
        Total_Records_Mismatch = $totalRecordMismatch
        Link_to_Mismatches = '<a href = " '+ $htmlPath+ '\'+ 'linesMismatchTable'+ '.html">Click here</a>'
    }

    #Create Html for the dashboard
    createHTMLFile -dataTable $dashboard -outFilePath $htmlPath -outFileName 'Dashboard' -Charset "UTF-8"

    $dashpath = $htmlPath+ '\Dashboard.html'

    Invoke-Item -Path $dashpath

}

#endregion script flow - sequence of steps 
