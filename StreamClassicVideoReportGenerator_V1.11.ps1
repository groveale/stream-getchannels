#
# Copyright (C) Microsoft Corporation.  All rights reserved.
#

<#
.SYNOPSIS
    Script for fetching all Stream Classic videos and exporting to a CSV
.AADTENANTID
    Aad Tenant Id of the customer.
.INPUTFILE
    File Path to import the Stream token from. EX: "C:\Users\Username\Desktop\token.txt"
.OUTDIR
    Folder Path where CSV will be exported. EX: "C:\Users\Username\Desktop"
.RESUMELASTRUN
    True/False. Whether execution should be resumed from last run or scratch. Default value is true.
.PUBLISHEDDATELE
    yyyy-mm-dd. Optional Parameter. Fetches video entries for which PublishedDate less than value. Default filter not applied. EX: "2021-02-15"
.PUBLISHEDDATEGE
    yyyy-mm-dd. Optional Parameter. Fetches video entries for which PublishedDate greater than value. Default filter not applied. EX: "2021-02-15"
.CREATEDESTINATIONPATHMAPPINGFORM365GROUPCONTAINERS
    True/False. Optional Parameter. If set true, the script will create a destination path mapping for M365Group containers. Inventory report generation will not be done.
.MIGRATIONDESTINATIONCSVFILEPATH
    File path to import details of M365Group containers, for which destination path mapping is required. Optional Parameter.
.CWCCREATORDETAILMAPPING
    True/False. Optional Parameter. If set true, the script will create a mapping of CWC creator details.
.CWCCREATORCSVFILEPATH
    File path to import details of CWC containers, for which creator details mapping is required. Optional Parameter.

Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true -PublishedDateLe "2022-02-15"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true -PublishedDateGe "2022-02-15"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -ResumeLastRun true -PublishedDateLe "2022-02-15" -PublishedDateGe "2021-02-15"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -CreateDestinationPathMappingForM365GroupContainers true -MigrationDestinationCsvFilePath "C:\Users\Username\Desktop\MigrationDestinations.csv"
Example:
.\StreamClassicVideoReportGenerator.ps1 -AadTenantId "00000000-0000-0000-0000-000000000000" -InputFile "C:\Users\Username\Desktop\token.txt" -OutDir "C:\Users\Username\Desktop" -CWCCreatorDetailMapping true -CWCCreatorCsvFilePath C:\Users\Username\Desktop\CWC\CWCContainerIds.csv
#>

[CmdletBinding()]
Param(
  [Parameter(Mandatory = $true)]
  [string]$AadTenantId,

  [Parameter(Mandatory = $true)]
  [string]$InputFile,

  [Parameter(Mandatory = $true)]
  [string]$OutDir,

  [Parameter(Mandatory = $false)]
  [string]$ResumeLastRun = 'true',

  [Parameter(Mandatory = $false)]
  [string]$PublishedDateLe,

  [Parameter(Mandatory = $false)]
  [string]$PublishedDateGe,

  [Parameter(Mandatory = $false)]
  [string]$CreateDestinationPathMappingForM365GroupContainers = 'false',

  [Parameter(Mandatory = $false)]
  [string]$MigrationDestinationCsvFilePath = '',

  [Parameter(Mandatory = $false)]
  [string]$CWCCreatorDetailMapping = 'true',

  [Parameter(Mandatory = $false)]
  [string]$CWCCreatorCsvFilePath
)

Function GetBaseUrl{
    $tenantPatchUri = "https://api.microsoftstream.com/api/tenants/"+$AadTenantId+"?api-version=1.4-private"

    $headers = @{
        Authorization = "Bearer $token"
    }

    $body = "{}"

    ((Get-Date).tostring() + ' TenantPatch URI: ' + $tenantPatchUri + "`n") | Out-File $logFilePath -Append

    try
    {
        $response = Invoke-RestMethod -Uri $tenantPatchUri -Method Patch -Body $body -Headers $headers -ContentType "application/json"
    }
    catch
    {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append

        #Stop execution if Unauthorized(401).
        if($_.Exception.Response.StatusCode.value__ -eq 401)
        {
            Write-Host "========Enter new token and start the script again======="
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n" -ForegroundColor Red
        exit
    }

    return $response.apiEndpoint
}

Function ReportOrchestration{
    Param(
        $baseUrl,
        $offsetId
    )

    $orchestrationUri = $baseUrl+'migrationReports/orchestration?api-version=1.0-odsp-migration'

    if($filter.Length -ne 0)
    {
        $orchestrationUri += '&$filter='+$filter
    }

    if($offsetId)
    {
        $orchestrationUri += '&$skiptoken=offsetId:'+$offsetId
    }

    $headers = @{
        Authorization = "Bearer $token"
    }

    ((Get-Date).tostring() + ' ReportOrchestration URI: ' + $orchestrationUri + "`n") | Out-File $logFilePath -Append

    try
    {
        $response = Invoke-RestMethod -Uri $orchestrationUri -Method Get -Headers $headers -ContentType "application/json"
    }
    catch
    {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append

        #Stop execution if Unauthorized(401).
        if($_.Exception.Response.StatusCode.value__ -eq 401)
        {
            Write-Host "========Enter new token and start the script again======="
        }

        #Stop execution if BadRequest(400)
        if($_.Exception.Response.StatusCode.value__ -eq 400)
        {
            $errorMessage = $Error[0].ErrorDetails
            Write-Host "$errorMessage"
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n" -ForegroundColor Red
        exit
    }

    return $response
}

Function CreateDirectory{
    Param(
        $category
    )

    $directoryPath = $OutDir+'\'+$category

    Write-Host "Checking if $category directory exists or not..."

    #Create directory if it does not exist.
    if(!(Test-Path $directoryPath))
    {
        Write-Host "Not found. Creating directory..."

        New-Item -Path $directoryPath -ItemType Directory | Out-Null

        Write-Host "Created directory. Path: $directoryPath."
    }
    else
    {
        Write-Host "Directory found."
    }

    return $directoryPath
}

Function MergeAndDedupe{
    $finalReportName = $reportPath + '\Report_' + $timeStamp;

    #Fetch all the intermediate CSVs, merge, sorting based on a unique Id - VideoId, remove duplicate entries
    $csvfile = Get-ChildItem -Filter *.csv -Path $reportPath -Recurse | Select-Object -ExpandProperty FullName | Import-Csv | Sort-Object VideoId -Unique 
    $totalVideoCount = $csvfile.Count - 1;
    $csvfile | Export-CSV "$finalReportName.csv" -NoTypeInformation

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $finalReportName.csv" -ForegroundColor Green
    Write-Host "******************************************************************************`n"

    Write-Host "Number of videos discovered: $totalVideoCount"
    ((Get-Date).tostring() + " Number of videos discovered: $totalVideoCount`n") | Out-File $logFilePath -Append
}

Function Dedupe{

    Param(
        $csvPath,
        $dateTimeNameSuffix
    )

    Write-Host "`nPreparing output directory..."

    $reportPath = CreateDirectory 'report'

    $reportFullName = $reportPath+'\StreamClassicVideoReport_'+$dateTimeNameSuffix+".csv"

    #Fetch all the temporary file(temp CSVs), merge, sorting based on a unique Id - VideoId, remove duplicate entries and keep the last entry it found in the merged file.
    $csvfile = Get-ChildItem -Filter *.csv -Path $csvPath | Select-Object -ExpandProperty FullName | Import-Csv | Sort-Object VideoId -Unique

    $finalReportName = 'StreamClassicVideoReport_'+$dateTimeNameSuffix

    $reportsFolderPath = CreateDirectory ('report\'+$finalReportName)

    $fileName = $reportsFolderPath+'\'+$finalReportName+'_'

    #SPLIT the deduped file to multiple CSVs of 10k records each.

    # variable used to advance the number of the row from which the export starts.
    $startrow = 0

    # counter used in names of resulting CSV files
    $counter = 1

    while ($startrow -lt $csvfile.Count)
    {
        #pick 10k records starting from the $startrow position and export content to a new file.
        $csvfile | Select-Object -skip $startrow -first 10000 | Export-CSV "$fileName$($counter).csv" -NoTypeInformation

        # Increment the number of the rows from which the export starts.
        $startrow += 10000

        # incrementing the $counter variable.
        $counter++

    }

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $reportsFolderPath"
    Write-Host "******************************************************************************`n"

    return $csvfile.Count
}

$functions = {

    function Get-ContainerVideoCount {
        param(
            [string]$baseUrl,
            [String]$containerId,
            [String]$containerType,
            [hashtable]$Headers = @{},
            [ref]$log

        )
        try {
            if($containerType -eq "CompanywideChannel")
            {
                $url = $baseUrl +'channels/'+ $containerId + "?adminmode=true&api-version=1.4-private"
            }
            elseif ($containerType -eq "M365Group" -or $containerType -eq "StreamOnlyGroup")
            {
                $url = $baseUrl +'groups/'+ $containerId + "?adminmode=true&api-version=1.4-private"
            } 
            else
            {
                # User container case  
                return "NA";
            }  
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $Headers
            return $response.metrics.videos
        }
        catch {
            $log.Value += ((Get-Date).tostring() + " Error Occured for Container Id :  $containerId , ContainerType $containerType  $_`n")
            return $null
        }
    }  
    
    Function WriteToCsv{
        Param(
            $reportData,
            $csvPath,
            $csvName, 
            $cwcFilePath,
            $cwcCsvName,
            [ref]$log
        )

        $ReportFilePath = $csvPath+'\'+$csvName
        $cwcContainerFile = $cwcFilePath+'\'+$cwcCsvName
        $csvHeaders = 'VideoId','Name','State','Description','PublishedDate','LastViewDate','Size (in Bytes)','Views','Likes','ContentType','PrivacyMode','Creator','Owners','ContainerId','ContainerName','ContainerType','ContainerEmailId','ContainerAadId','MigratedDestination','ContainerVideosInClassicUI','IsEligibleForMigration';

        $ActionDelegate = {
            param($video)
            $row = '';
            $row += '"' + $video.id + '",';
            $row += '"' + $video.name.Replace('"','""') + '",';
            $row += '"' + $video.state + '",';
            if ($video.description)
            {
                $row += '"' + $video.description.Replace('"','""') + '",';
            }
            else
            {
                $row += '"",';
            }
            if($video.publishedDate -eq "9999-12-31T23:59:59.9999999Z")
            {
                $row += '"",'
            }
            else
            {
                if($video.publishedDate) {
                    $video.publishedDate  = Get-Date -Date $video.publishedDate -format "MM/dd/yyyy HH:mm:ss";
                }
                $row += '"' + $video.publishedDate + '",';
            }
            if($video.lastViewDate) {
                $video.lastViewDate  = Get-Date -Date $video.lastViewDate -format "MM/dd/yyyy HH:mm:ss";
            }
            $row += '"' + $video.lastViewDate + '",';
            if($video.size)
            {
                $row += '"' + $video.size.tostring() + '",';
            }
            else
            {
                $row += '"",'
            }
            if($video.container.containerType -eq 'CompanywideChannel'){
                if (!(Test-Path $cwcContainerFile -PathType leaf))
                {
                    New-Item -Path $cwcContainerFile -ItemType File | Out-Null
                }
                $cwcContainerobject = [PSCustomObject]@{
                    'ContainerId' = $video.container.id
                }
                # Adding the CWC containerId to the CSV file
                $cwcContainerobject | Export-Csv -Path $cwcContainerFile -Append -NoTypeInformation
            }
            $row += '"' + $video.viewCount.tostring() + '",';
            $row += '"' + $video.likeCount.tostring() + '",';
            $row += '"' + $video.contentType + '",';
            $row += '"' + $video.privacyMode + '",';
            $row += '"' + $video.creator + '",';
            $row += '"' + $video.owners + '",';
            $row += '"' + $video.container.id + '",';
            $row += '"' + $video.container.name + '",';
            $row += '"' + $video.container.containerType + '",';
            $row += '"' + $video.container.emailId + '",';
			$row += '"' + $video.container.containerAadId + '",';
            $row += '"' + $video.destinationUrl + '",';
            if ($null -ne $video.container.id -and $containerVideoCountMap.ContainsKey($video.container.id)) {
                $row += '"' + $containerVideoCountMap[$video.container.id] + '",'
            }
            else {
                $row += '"NA",'
            }
            if ($video.state -in @("Processing", "Completed") -and $video.publishedDate -ne "9999-12-31T23:59:59.9999999Z" -and ($video.container.id -ne "" -and $video.container.id -ne "00000000-0000-0000-0000-000000000000")) {
                $isEligible = "Yes"
            }
            else {
                $isEligible = "No"
            }
            $row += '"' + $isEligible + '"'
            return $row
        } 
        # Create file if doesn't exist
        if (!(Test-Path $ReportFilePath -PathType leaf))
        {
            New-Item -Path $ReportFilePath -ItemType File | Out-Null;
        }
        $lengthOfHeadersRow = (Get-Content $ReportFilePath | Select-Object -First 1).Length;
        if($lengthOfHeadersRow -eq 0)
        {
            #It will be 0, in case of newly created file
            #Put the headers in the file first
            Add-Content -Path $ReportFilePath -Value (($csvHeaders| ForEach-Object { return $_}) -join ',');
        }
        #Writing content to temp file
        foreach($video in $reportData)
        {
            $row = Invoke-Command $ActionDelegate -ArgumentList $video
            $row | Out-File $ReportFilePath -encoding utf8 -Append
        }
    }

    Function ReportDetailsAndWriteToCsv{
        Param(
            $baseUrl,
            $offsetId,
            $csvPath,
            $csvName,
            $token,
            $filter,
            $cwcFilePath,
            $cwcCsvName
        )

        $log = ''
        $status = ''
        $statusCode = ''

        $reportUri = $baseUrl+'migrationReports?api-version=1.0-odsp-migration'

        if($filter.Length -ne 0)
        {
            $reportUri += '&$filter='+$filter
        }

        if($offsetId)
        {
            $reportUri += '&$skiptoken=offsetId:'+$offsetId
        }

        $headers = @{
            Authorization = "Bearer $token"
        }

        $log += ((Get-Date).tostring() + ' ReportDetails URI: ' + $reportUri + "`n")

        try
        {
            $global:containerVideoCountMap = @{}
            #ReportDetails API call
            $response = Invoke-RestMethod -Uri $reportUri -Method Get -Headers $headers -ContentType "application/json"

            foreach($video in $response.value)
            {
                if ($null -ne $video.container.id -and !$containerVideoCountMap.ContainsKey($video.container.id)) 
                {
                    $videoCount = Get-ContainerVideoCount -baseUrl $baseUrl -containerId $video.container.id -containerType $video.container.containerType -Headers $headers -log ([ref]$log)
                    if($null -ne  $videoCount){
                        $containerVideoCountMap.Add($video.container.id, $videoCount)  
                    }   
                }
                
            }
            
            #Write content returned from ReportDetails API call to a temp file(CSV).
            WriteToCsv $response.value $csvPath $csvName $cwcFilePath $cwcCsvName ([ref]$log)

            $status = "Success"
        }
        catch
        {
            $log += ((Get-Date).tostring() + ' ' + $_.Exception + "`n")
            
            $log += ((Get-Date).tostring() + ' Error: ' + $Error + "`n")

            #If API or write to CSV fails then status should be written as FAILED in State.csv for this OffsetId.
            $status = "Failed"

            if($_.Exception.Response.StatusCode.value__ -eq 401)
            {
                $statusCode = $_.Exception.Response.StatusCode.value__

                Write-Host "========Enter new token and start the script again======="
            }
                
            if($_.Exception.Response.StatusCode.value__ -eq 400)
            {
                $statusCode = $_.Exception.Response.StatusCode.value__
                $errorMessage = $Error[0].ErrorDetails
                Write-Host "$errorMessage"
            }

            Write-Host "`nSome error occurred. Check logs for more info.`n"
        }

        return @($status, $log, $statusCode)
    }
}

Function AddHeadersToStateFile{
    Param(
        $stateFilePath
    )

    $csvHeaders = 'OffsetId','RetryCount','Status';

    $lengthOfHeadersRow = (Get-Content $stateFilePath | Select-Object -First 1).Length;

    if($lengthOfHeadersRow -eq 0)
    {   # It will be 0, in case of newly created file
        # Put the headers in the file first
        Add-Content -Path $stateFilePath -Value (($csvHeaders| ForEach-Object { return $_}) -join ',');
    }
}

Function CreateFile{
    Param(
        $filePath
    )

    if((Test-Path $filePath -PathType leaf))
    {
        Remove-Item ($filePath)
    }

    New-Item -Path $filePath -ItemType File | Out-Null

    Write-Host "Created $filePath file."
}

Function RetryFailedOffsetIds{
    Param(
        $stateFilePath,
        $csvPath,
        $csvNamePrefix,
        $baseUrl,
        $parallelCalls,
        $cwcFilePath
    )

    $i = 0

    ((Get-Date).tostring() + " Retrying failed offset Ids, if any.`n") | Out-File $logFilePath -Append

    $file = Import-csv $stateFilePath

    #Fetching those records from state.csv which are FAILED and have not been retried 5 times yet.
    $file = @($file | Where-Object {($_.Status -eq "Failed") -and ([int]$_.RetryCount -lt 5)})

    $NumberOfFailedEntries = $file.Length

    while($i -lt $NumberOfFailedEntries)
    {
        $j = 1

        #Retrying the failed Ids parallely in the batch of degreeOfParallelism.
        while(($j -le $parallelCalls) -and ($i -lt $NumberOfFailedEntries))
        {
           $csvName = $timeStamp + '_' + $j.ToString() +'.csv'
           $cwcCsvName ='cwcContainerId'+ '_' + $timeStamp + '_' + $i.ToString() +'.csv' 
           Start-Job -Name $file[$i].OffsetId -InitializationScript $functions -ScriptBlock {

                  param($baseUrl, $offsetId, $csvPath, $csvName, $token, $cwcFilePath, $cwcCsvName)

                  #Fetches the content and write into individual temporary files.
                  ReportDetailsAndWriteToCsv -baseUrl $baseUrl -offsetId $offsetId -csvPath $csvPath -csvName $csvName -token $token -cwcFilePath $cwcFilePath -cwcCsvName $cwcCsvName

                } -ArgumentList $baseUrl, $file[$i].OffsetId, $csvPath, $csvName, $token , $cwcFilePath, $cwcCsvName| Out-Null

            $j++
            $i++
        }

        Get-Job | Wait-Job | Out-Null

        $jobs = Get-job | Sort-Object -Property {[int]$_.Name}

        foreach($job in $jobs)
        {
            $result = $job | Receive-job

            #If Job failes to complete then status should be 'Failed' for it to be retried on next run.
            #If Job succeeds, then status will be returned from the execution code.
            if($job.JobStateInfo.State -ne 'Completed')
            {
                $result[0] = 'Failed'
            }

            $result[1] | Out-File $logFilePath -Append

            #UPDATE the state.csv file for previously failed OffsetId.

            $stateFile = $null

            $stateFile = Import-Csv -Path $stateFilePath

            if(@($stateFile).Length -gt 0)
            {
                $stateFile | ForEach-Object{
                    #update the status and retryCount for the matching OffsetId.
                    if($job.Name -eq $_.OffsetId)
                    {
                        if($_.Status -eq "Failed")
                        {
                            $_.Status = $result[0]
                            $_.RetryCount = ([int]$_.RetryCount+1)
                        }
                    }
                }

                #Export the locally written data back to state.csv.
                $stateFile | Export-Csv -Path $stateFilePath -NoTypeInformation
            }

            if($result[2] -eq 401)
            {
                exit
            }
        }

        Get-Job | Remove-Job
        Merge-CSVFiles -rootFolder $cwcFilePath
    }
}

function Merge-CSVFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$rootFolder
    )

    $outputFilePath = $rootFolder + "\CWCContainerIds.csv"
    # Get all CSV files in the root folder
    if(-not (Test-Path $rootFolder)){
        return
    }
    $csvFiles = Get-ChildItem -Path $rootFolder -Filter "*.csv"

    # Create an empty array to store the merged data
    $mergedData = @()

    # Iterate through each CSV file
    foreach ($csvFile in $csvFiles) {
        # Import the CSV file and append its data to the merged data array
        $data = Import-Csv -Path $csvFile.FullName | Select-Object -Property * -Unique
        $mergedData += $data
    }

    # Export the merged data to a new CSV file
    $mergedData | Export-Csv -Path $outputFilePath -NoTypeInformation

    #cleanup unused files
    $excludedFiles = @("CWCContainerIds.csv")
    foreach ($csvFile in $csvFiles) {
        if ($excludedFiles -notcontains $csvFile.Name) {
            Remove-Item -Path $csvFile.FullName -Force
        }
    }
}
    
Function GenerateFilterString{
    Param([ref]$filter)

    if ($PublishedDateLe.Length -ne 0)
    {
        $filter.Value = 'publishedDate le ' + $PublishedDateLe
    }

    if ($PublishedDateGe.Length -ne 0)
    {
        if ($filter.Value.Length -ne 0)
        {
            $filter.Value += ' and '
        }

        $filter.Value += 'publishedDate ge ' + $PublishedDateGe
    }
}

Function StartReportGeneration{

    $startDateTime = Get-Date
    Get-Job | Remove-Job

    $stateDirectoryPath = CreateDirectory 'state'

    Write-Host "------------------------------------------------------------`n"
    $reportFolderName = 'StreamClassicVideoReport'
    $reportPath = CreateDirectory $reportFolderName

    $stateFilePath = $stateDirectoryPath+'\state.csv'
    $lastRunFilePath = $stateDirectoryPath+'\lastRunFolder.txt'
    $cwcFilePath = CreateDirectory 'CWC'
    $offsetIdsList = @()

    # Initializing the report output directory filename
    $timeStamp = Get-Date -Format FileDateTime
    $csvPathFromOutDir = $reportFolderName + '\' + $timeStamp

    # If execution is being started from scratch, create the required files.
    if($ResumeLastRun -eq 'false')
    {
        CreateFile $stateFilePath

        AddHeadersToStateFile $stateFilePath

        CreateFile $logFilePath
      
        CreateFile $lastRunFilePath
        if(Test-Path $cwcFilePath -PathType leaf){
            Remove-Item -Path $cwcFilePath -Recurse | Out-Null
        }
        $cwcFilePath = CreateDirectory 'CWC'
        #In case of fresh run, create new output directory and save that path in the lastRunFolder file
        $csvPath = CreateDirectory $csvPathFromOutDir
        Set-Content -Path $lastRunFilePath -Value $csvPath

    }
    else
    {
        #Checking for state.csv for resumption.
        Write-Host "`nSearching for last saved state..."

        if (!(Test-Path $logFilePath -PathType leaf))
        {
            CreateFile $logFilePath
        }

        if (!(Test-Path $stateFilePath -PathType leaf))
        {
            Write-Host "No saved state found. Preparing to fetch from scratch...`n"
            CreateFile $stateFilePath

            AddHeadersToStateFile $stateFilePath
        }
        else
        {
            #resume state
            $stateData = Import-Csv $stateFilePath
            $offsetIdsList = $stateData | Select-Object -ExpandProperty OffsetId
        }

        if (!(Test-Path $lastRunFilePath -PathType leaf)) {
            #For backward compatibility, as script executions prior to this version will not have the new lastRunFolder file
            #so creating new output directory and save that path in the lastRunFolder file
            CreateFile $lastRunFilePath
            $csvPath = CreateDirectory $csvPathFromOutDir
            Set-Content -Path $lastRunFilePath -Value $csvPath
        }
        else {
            #Fetch the LastRunFolder incase of resumed script run
            $csvPath = Get-Content $lastRunFilePath
        }
    }

    ((Get-Date).tostring() + " ------START------`nScript version: $scriptVersion`n") | Out-File $logFilePath -Append

    $baseUrl = GetBaseUrl
    $csvNamePrefix = ""

    Write-Host "------------------------------------------------------------`n"

    Do{
        $offsetId = $offsetIdsList | Select-Object -Last 1

        $orchestrationResponse = ReportOrchestration $baseUrl $offsetId
        $orchestrationResponse | Out-File $logFilePath -Append

        $offsetIdsList = $orchestrationResponse.offsetIds
        $parallelThreads = $orchestrationResponse.degreeOfParallelism
        #Check if no offset ids were returned for the first orchestration call in the script execution (whether resuming last run or starting new run)
        if($firstOrchestration -and ($offsetIdsList.Length -eq 0)) {
            if($ResumeLastRun -eq 'true') {
                Write-Host "`nNo new videos found" -ForegroundColor Yellow| Out-File $logFilePath -Append
                break #break from the loop as we need to retry for failed offsets, if any, incase it was a resume run
            }
            else {
                #remove report folder created before
                Remove-Item ($csvPathFromOutDir) -Confirm:$false -Force -Recurse
                Write-Host "`nNo videos found" -ForegroundColor Yellow| Out-File $logFilePath -Append
                exit
            }
        }   
        $firstOrchestration = $false;

        Write-Host "`nExtracting data..."

        $i = 1

        #Fetching ReportDetails in parallel.
        #Number of jobs = Number of Ids in offsetIdsList(fetched from Orchestration API) OR degree of parallelism.
        foreach($Id in $offsetIdsList)
        {
           $csvName = $timeStamp + '_' + $i.ToString() +'.csv'
           $cwcCsvName ='cwcContainerId'+ '_' + $timeStamp + '_' + $i.ToString() +'.csv'

           Start-Job -Name $Id -InitializationScript $functions -ScriptBlock {

                  param($baseUrl, $offsetId, $csvPath, $csvName, $token, $filter, $cwcFilePath, $cwcCsvName)

                  #Fetches the content and write into individual temporary files.
                  ReportDetailsAndWriteToCsv -baseUrl $baseUrl -offsetId $offsetId -csvPath $csvPath -csvName $csvName -token $token -filter $filter -cwcFilePath $cwcFilePath -cwcCsvName $cwcCsvName

                } -ArgumentList $baseUrl, $Id, $csvPath, $csvName, $token, $filter, $cwcFilePath, $cwcCsvName| Out-Null

            $i++
        }

        Get-Job | Wait-Job | Out-Null

        $jobs = Get-job | Sort-Object -Property {[int]$_.Name}

        foreach($j in $jobs)
        {
            $result = $j | Receive-job

            #If Job failes to complete then status should be 'Failed' for it to be retried on next run.
            #If Job succeeds, then status will be returned from the execution code.
            if($j.JobStateInfo.State -ne 'Completed')
            {
                $result[0] = 'Failed'
            }

            #Write logs, returned from the job, into the file.
            $result[1] | Out-File $logFilePath -Append

            # Write to state.csv
            ('"' + $j.Name.ToString() + '","0","' + $result[0] + '"') | Out-File $stateFilePath -encoding ASCII -Append

            #If API call failed with 401, then there is no need to continue and exit immediately as token is not valid anymore.
            if($result[2] -eq 401)
            {
                exit
            }
        }

        Get-Job | Remove-Job
        Merge-CSVFiles -rootFolder $cwcFilePath

    }while($offsetIdsList)

    #Once all OffsetIds have been processed, scan state.csv to retry failed ones.
    RetryFailedOffsetIds $stateFilePath $csvPath $csvNamePrefix $baseUrl $parallelThreads $cwcFilePath

    Write-Host "`n******************************************************************************"
    Write-Host "Reports are available at this location: $csvPath" -ForegroundColor Green
    Write-Host "******************************************************************************`n"

    $endDateTime = Get-Date

    Write-Host "Time elapsed: "($endDateTime - $startDateTime)"`n"
    ((Get-Date).tostring() + " Time elapsed: "+($endDateTime - $startDateTime)+"`n") | Out-File $logFilePath -Append

    if (Test-Path $stateFilePath -PathType leaf)
    {
        #check for failed offsets in state even after retrying
        $stateCsvData = Import-csv $stateFilePath
        $failedOffsets = @($stateCsvData | Where-Object {($_.Status -eq "Failed")})
    
        if($failedOffsets.Length -gt 0){
            # saving the failed offsets in logs
            $failedOffsets | Out-File $logFilePath -Append
            Write-Host "Failed to fetch details for some videos." -ForegroundColor Yellow
            Write-Host "Please re-run the script. If the error persists, kindly reach out to Customer Support and share the log file: $logFilePath"
        }
    }
}

Function GetM365GroupSharepointDocumentUrl{
    Param(
        $M365GroupEmailId
    )

    try
    {
        # Get sharepoint destination url of default site lib, for the given M365 group email id
        $output= Get-UnifiedGroup -Identity $M365GroupEmailId | Select SharepointDocumentsUrl
        return $output
    }
    catch
    {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append
    }
}

Function CreateMigrationDestinationPathMapping{
    Write-Host "Started process for creating Auto-Mapping of M365 group containers to Migration destination path"
    "Started process for creating Auto-Mapping of M365 group containers to Migration destination path" | Out-File $logFilePath -Append
    
    $rowHashSet = @{}
    if($MigrationDestinationCsvFilePath -eq $null -Or $MigrationDestinationCsvFilePath -eq $empty) {
        Write-Host "MigrationDestinationCsvFilePath is required. Currently its null or empty"
        "MigrationDestinationCsvFilePath is required. Currently its null or empty" | Out-File $logFilePath -Append
        exit
    } else {
        Write-Host "Reading csv file:" + $MigrationDestinationCsvFilePath
        "Reading csv file:" + $MigrationDestinationCsvFilePath| Out-File $logFilePath -Append
    }
    
    # Read migration destination csv file
    $csvFile = Import-Csv -Path $MigrationDestinationCsvFilePath

    if(@($csvFile).Length -gt 0)
    {
        $csvFile | ForEach-Object{
            #extract ContainerId and ContainerEmailId
            if("" -eq $_.'Destination Path' -And $_.'Source Path (for reference only)'.ToString().EndsWith('M365Group'))
            {
                $row = '';
                $row += $_.'Name (for reference only)' + ',';
                $row += $_.'Source Path (for reference only)';
                $row += ',#DestinationPathPlaceHolder#,'
                $row += $_.'Migration status (for reference only)' + ',"';
                $row += $_.'Created on (for reference only)' + '",';
                $row += $_.'Task ID';
                
                $sourcePathFragments = $_.'Source Path (for reference only)'.Split("|")
                if($sourcePathFragments.Count -eq 3) {
                    $groupType = $sourcePathFragments[2]
                    $groupEmail = $sourcePathFragments[0]
                
                    # If container type is M365Group, extract the same
                    if($groupType -eq 'M365Group' -And !$rowHashSet.ContainsKey($row)) {
                        $rowHashSet.Add($row, $groupEmail)
                    }
                } else {
                    "Could not parse source path (for reference only). Incorrect value:" + $_.'Source Path (for reference only)'.ToString() | Out-File $logFilePath -Append
                }
            }
        }
    } else {
        Write-Host "Could not read file at MigrationDestinationCsvFilePath. Please check path"
        "Could not read file at MigrationDestinationCsvFilePath. Please check path" | Out-File $logFilePath -Append
        exit
    }

    $mappingFilePath = $OutDir + '\MigrationDestinationMappingForM365GroupContainers_' + $randomGuidForOutFiles.Guid + '.csv'
    $headers = "Name (for reference only),Source Path (for reference only),Destination Path,Migration status (for reference only),Created on (for reference only),Task ID"
    $headers | Out-File $mappingFilePath -encoding utf8 -Append
    $atleastOneRowOutputIsDone = $false

    Foreach ($row in $rowHashSet.Keys){
        # Get Mapping of Default site lib path
        $emaildId = $rowHashSet[$row]
        $DocumentLibrary = GetM365GroupSharepointDocumentUrl $emaildId
        $destinationPath = $DocumentLibrary.SharePointDocumentsUrl

        if($destinationPath -eq $null -or $destinationPath.ToString().Trim(' ') -eq "") {
            $logVal = "Failed to retrive or No SharePointDocumentsUrl discovered for row:" + $row.ToString() 
            $logVal += " .Please check if sharepoint site exists for this group. Please check log, if there is an error. If no error is thrown, SharepointDocumentUrl doesn't exist for your site."
            $logVal += " You can manually assign Migration destination path for this case."
            $logVal | Out-File $logFilePath -Append 
        } else {
            $updatedRow = $row.ToString() -replace '#DestinationPathPlaceHolder#',$destinationPath
        
            # Push mapping data out to csv file
            $updatedRow | Out-File $mappingFilePath -encoding utf8 -Append
            $atleastOneRowOutputIsDone = $true
        }
    }

    if($atleastOneRowOutputIsDone -eq $true){
        Write-Host "Done. Exiting auto-mapping. A new csv with destination path mapping for M365Group containers, has been created in the OutDir. Please use this to upload destination path in migration tool. FilePath:" + $mappingFilePath
        "Done. Exiting auto-mapping. A new csv with destination path mapping for M365Group containers, has been created in the OutDir. Please use this to upload destination path in migration tool. FilePath:" + $mappingFilePath | Out-File $logFilePath -Append 
    } else {
        Write-Host "Done. Empty output file generated. Either no M365Group container with unassigned path was detected or there was error in fetching path. Please check logs. LogFilePath:" $logFilePath
    }
}

function Test-IsGuid {
    param([string]$guid, $logFilepath)
    $regex = [regex]::new('^[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}$')
    $isGuid = $regex.IsMatch($guid)
    if($isGuid -eq $false) {
        ((Get-Date).tostring() + " Invalid CWC Id $guid")| Out-File -FilePath $logFilepath -Append
    }
    return $isGuid
}

Function FetchCreatorDetails{
    Param(
        $baseUrl,
        $containerIds,
        $outputFilePath,
        $logFilepath,
        $stateFilepath
    )
    $cwcCreatorDetailsUri = $baseUrl+'migrationReports/cwcdetails?api-version=1.0-odsp-migration'

    $headers = @{
        Authorization = "Bearer $token"
    }
    $requestBody = @{
        ChannelIds = @()
    }     

    $expandedContainerIds = $containerIds | Select-object -ExpandProperty 'containerId'
    $requestBody.ChannelIds += $expandedContainerIds
    $apiRequestBody = $requestBody | ConvertTo-Json -Depth 10

    try
    {
        ((Get-Date).tostring() + " Calling the creator Uri $cwcCreatorDetailsUri with request body $apiRequestBody ")| Out-File -FilePath $logFilepath -Append
        $response = Invoke-RestMethod -Uri $cwcCreatorDetailsUri -Method Post -Body $apiRequestBody -Headers $headers -ContentType "application/json"

        foreach($value in $response.value){
            if ([string]::IsNullOrEmpty($value.creatorEmailId)){
                $status = 'failed'
            }
            else{
                $status = 'success'
            }
            $creatorobj= [PSCustomObject]@{
                'ContainerId' = $value.Id
                'Name' = $value.Name
                'Description' =  $value.Description
                'CreatorEmailId' =  $value.CreatorEmailId
            }
            $stateObj = [PSCustomObject]@{
                'ContainerId' = $value.Id
                'status' = $status
            }
            ((Get-Date).tostring() + " channel creator output CWC Id $value.Id")| Out-File -FilePath $logFilepath -Append
            $creatorobj | Export-Csv -Path $outputFilePath -Append -NoTypeInformation
            $stateObj | Export-Csv -Path $stateFilePath -Append -NoTypeInformation
        }
    }
    catch
    {
        $errorMessage = $_.Exception.Message
        $headers = $_.Exception.Response.Headers
        if($headers.Contains('x-ms-request-id')){
            $msRequestId = $headers.GetValues('x-ms-request-id')[0]
        }else{
            $msRequestId = ''
        }
        
        ((Get-Date).tostring() + " Failed to fetch Creator email ERROR Message $errorMessage having requestId $msRequestId")| Out-File -FilePath $logFilepath -Append

        if($_.Exception.Response.StatusCode.value__ -eq 401)
        {
            $statusCode = $_.Exception.Response.StatusCode.value__

            Write-Host "========Enter new token and start the script again======="
        }
            
        if($_.Exception.Response.StatusCode.value__ -eq 400)
        {
            $statusCode = $_.Exception.Response.StatusCode.value__
            Write-Host "$errorMessage"
        }

        Write-Host "`nSome error occurred. Check logs for more info.`n"
    }
}

Function GetCWCCreatorDetails{
    Param(
            $cwcOutputDirName,
            $logFilepath,
            $stateFilepath
        )
    
    # If execution is being started from scratch, create the create the new state, log and last run files .
    if($ResumeLastRun -eq 'false')
    {
        $csvHeaders = 'ContainerId','Status';
        CreateFile $stateFilePath
        Add-Content -Path $stateFilePath -Value (($csvHeaders| ForEach-Object { return $_}) -join ',');

        CreateFile $logFilePath
    }
    else
    {
        #Checking for state.csv for resumption.
        Write-Host "`nSearching for last saved state..."

        if (!(Test-Path $logFilePath -PathType leaf))
        {
            CreateFile $logFilePath
        }

        if (!(Test-Path $stateFilePath -PathType leaf))
        {
            Write-Host "No saved state found. Preparing to fetch from scratch...`n"
            $csvHeaders = 'ContainerId','Status';
            CreateFile $stateFilePath
            Add-Content -Path $stateFilePath -Value (($csvHeaders| ForEach-Object { return $_}) -join ',')
        }
    }

    ((Get-Date).tostring() + " ------START------`nScript version: $scriptVersion`n") | Out-File $logFilePath -Append

    $cwcOutputDirPath = $OutDir + "\" + $cwcOutputDirName
    $outFileName = "CWCCreatorValues.csv"
    $outputFilePath = "$cwcOutputDirPath\" + $outFileName

    $baseUrl = GetBaseUrl
    ((Get-Date).tostring() + " Base Url $baseUrl")| Out-File -FilePath  $logFilePath -Append
    Write-Host "------------------------------------------------------------`n"
    Write-Host "Started process for fetching creator email for input CWC Container ID's"
    "Started process for fetching creator email for input CWC Container ID's" | Out-File $logFilePath -Append
    $stateData = Import-Csv $stateFilePath

    # Read CWC containerId's csv file
    $cwcCsvInput= Import-Csv -Path $CWCCreatorCsvFilePath
    $csvInputData = @($cwcCsvInput)
    $batchSize = 100
    $j = 0
    while($j -lt $csvInputData.Count){
        if($csvInputData.Count -eq 1){
            $batch = $csvInputData[0]
        }else{
            $batch = $csvInputData[$j..($j + $batchSize-1)]
        }
        $data = $batch | Where-Object { ($_.'ContainerId' -notin $stateData.'ContainerId') -and (Test-IsGuid $_.'ContainerId' $logFilePath) }
        if([string]::IsNullOrEmpty($data)){
            $j = $j + $batchSize
            continue
        }
        if (!(Test-Path $cwcOutputDirName)) {
            $cwcOutputDirPath = CreateDirectory $cwcOutputDirName
        }
        if (!(Test-Path $outputFilePath -PathType leaf)) {
            CreateFile $outputFilePath 
        }

        FetchCreatorDetails -baseUrl $baseUrl -containerIds $data -outputFilePath $outputFilePath -logFilepath $logFilePath -stateFilepath $stateFilePath
        $j = $j + $batchSize
    }
}

$scriptVersion = '1.11'

Write-Host "Script version: $scriptVersion`n"

if($CreateDestinationPathMappingForM365GroupContainers -ne 'true'){
    Write-Host "`n////////////////////////////////////////////////////"
    Write-Host "Generating Stream Classic video report..."
    Write-Host "////////////////////////////////////////////////////`n"

    $token = Get-Content -Path $InputFile

    $logFilePath = $OutDir+'\logs.txt'

    $parallelThreads = 0
    $firstOrchestration = $true;

    [string]$filter = ''

    GenerateFilterString ([ref]$filter)

    StartReportGeneration
} else {

    $randomGuidForOutFiles = New-Guid
    $logFilePath = $OutDir+'\logs_' + $randomGuidForOutFiles.Guid + '.txt'

    if (!(Test-Path $logFilePath -PathType leaf))
    {
        CreateFile $logFilePath
    }

    Write-Host "Installing module ExchangeOnlineManagement, which requires admin priviledge on script"
    #Install Exchange Online Management Shell. Needs admin access to run
    Install-Module -Name ExchangeOnlineManagement
 
    #Connect to Exchange Online to allow M365Group details to be fetched. User will be prompted to login. This will work only for Exchange online admins. Documentation: https://learn.microsoft.com/en-us/powershell/module/exchange/get-unifiedgroup?view=exchange-ps
    Connect-ExchangeOnline

    try
    {
        #Invoke function to generate destination path mapping.
        CreateMigrationDestinationPathMapping
    }
    catch
    {
        #Log error.
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append
        
        Write-Host "Please check error for details. If required, kindly reach out to Customer Support and share the log file: $logFilePath"
    }    
} 


if($CWCCreatorDetailMapping -eq 'true') {

    $token = Get-Content -Path $InputFile
    $cwcreportFolderName =  'CWCCreatorReport'
    $CWCCreatorReport = CreateDirectory $cwcreportFolderName
    $timeStamp = Get-Date -Format FileDateTime
    # Creating the new output directory 
    $cwcOutputDirName = $cwcreportFolderName+ '\' + $timeStamp
    
    $stateDirectoryPath = CreateDirectory 'CWCCreatorReport\state'
    $stateFilePath = $stateDirectoryPath+'\state_cwc.csv'
    $logFilePath = $CWCCreatorReport +'\logs.txt'
    if (!(Test-Path $logFilePath -PathType leaf))
    {
        CreateFile $logFilePath
    }
    $cwcStartTime = Get-Date
    $isValidPath = $true
    try
    {
        $cwcDirectoryPath = $OutDir+'\CWC'
        $cwcDirectoryCsvPath = $cwcDirectoryPath+'\CWCContainerIds.csv'
        if ([string]::IsNullOrEmpty($CWCCreatorCsvFilePath)){
            $CWCCreatorCsvFilePath = $cwcDirectoryCsvPath
            if (Test-Path $CWCCreatorCsvFilePath) {
                $csvContent = Get-Content $CWCCreatorCsvFilePath
                if ($csvContent.Length -gt 0) {
                    Write-Host "Reading csv file on CWCCreatorCsvFilePath: $CWCCreatorCsvFilePath"
                    "Reading csv file on CWCCreatorCsvFilePath: $CWCCreatorCsvFilePath"| Out-File $logFilePath -Append
                    
                } else {
                    Write-Host "CSV file exists at CWCCreatorCsvFilePath $CWCCreatorCsvFilePath but is empty." -ForegroundColor Yellow
                    "CSV file exists at CWCCreatorCsvFilePath $CWCCreatorCsvFilePath but is empty"| Out-File $logFilePath -Append
                    $isValidPath = $false
                }
            } else {
                Write-Host "CWCCreatorCsvFilePath is required. CSV file does not exist at $cwcDirectoryCsvPath" -ForegroundColor Yellow
                "CWCCreatorCsvFilePath is required. CSV file does not exist at $cwcDirectoryCsvPath" | Out-File $logFilePath -Append
                $isValidPath = $false
            }
           
        }
        #Invoke function to generate CWC creator mapping.
        if( $isValidPath){
            GetCWCCreatorDetails $cwcOutputDirName $logFilepath $stateFilepath
        }
        if(Test-Path $cwcDirectoryPath){
            Remove-Item $cwcDirectoryPath -Force -Recurse
        }
        $cwcEndTime = Get-Date
        Write-Host "Time elapsed to fetch CWC creator details: "($cwcEndTime - $cwcStartTime)"`n"
        ((Get-Date).tostring() + " Time elapsed to fetch CWC creator details: "+($cwcEndTime - $cwcStartTime)+"`n") | Out-File $logFilePath -Append
    }
    catch
    {
        ((Get-Date).tostring() + ' ' + $Error[0].ErrorDetails + "`n") | Out-File $logFilePath -Append
        ((Get-Date).tostring() + ' ' + $Error[0].Exception + "`n") | Out-File $logFilePath -Append

        Write-Host "Please check error for details. If required, kindly reach out to Customer Support and share the log file: $logFilePath"  -ForegroundColor Red
    }
}
# SIG # Begin signature block
# MIInvwYJKoZIhvcNAQcCoIInsDCCJ6wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCVery3jeCRdO3S
# 82mjOkamoCnv9zU9ESnLDovFpVBwFqCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGZ8wghmbAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIBWbLMzIH20JyXNzouPDbhgB
# 6R3y+dnX+niYW2kF3FqbMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEARv2oK5S9MwKDMDS+epdI4z3hRgzjzCK7l+Ld8G/0Pgv+lhinnLzGIDa9
# OJp1FSRjrjWDI9wY8VeNjDTQq82ivDFBx624tLr8A7XgGY6E7aJMG+Rm3n6bHs6U
# W9S4XB2ZJakHaEvuE9q7d50zluX1Y0SFS6er9KRFKeRRb50dZjTo4bByGAC800jL
# ghmpTTYNskIyRY62cNPXHKiXc2IUj1aotjsAHOI/KdDqnBY+6HGEtG/gWpkS9yr3
# gdKbQMVaEXBxMs1wbjaYSpkO8zORJPyjl5yn8X/NKxZ1R4/s5R0PutvU74tenLef
# XFjbEgNr33eOASaehyPaGjsn5cm0E6GCFykwghclBgorBgEEAYI3AwMBMYIXFTCC
# FxEGCSqGSIb3DQEHAqCCFwIwghb+AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFZBgsq
# hkiG9w0BCRABBKCCAUgEggFEMIIBQAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCALQFPKpDu4yp/IBdPPa2kCY8W4XqtbFjPZYR4LAm4MtQIGZN/NV8ol
# GBMyMDIzMDkxMzE1NTExNy42NjNaMASAAgH0oIHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OkQwODItNEJGRC1FRUJBMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloIIReDCCBycwggUPoAMCAQICEzMAAAG6Hz8Z98F1vXwAAQAAAbowDQYJ
# KoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjIw
# OTIwMjAyMjE5WhcNMjMxMjE0MjAyMjE5WjCB0jELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3Bl
# cmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpEMDgyLTRC
# RkQtRUVCQTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIhOFYMzkjWAE9UVnXF9hRGv
# 0xBRxc+I5Hu3hxVFXyK3u38xusEb0pLkwjgGtDsaLLbrlMxqX3tFb/3BgEPEC3L0
# wX76gD8zHt+wiBV5mq5BWop29qRrgMJKKCPcpQnSjs9B/4XMFFvrpdPicZDv43FL
# gz9fHqMq0LJDw5JAHGDS30TCY9OF43P4d44Z9lE7CaVS2pJMF3L453MXB5yYK/KD
# bilhERP1jxn2yl+tGCRguIAsMG0oeOhXaw8uSGOhS6ACSHb+ebi0038MFHyoTNhK
# f+SYo4OpSY3xP4+swBBTKDoYP1wH+CfxG6h9fymBJQPQZaqfl0riiDLjmDunQtH1
# GD64Air5k9Jdwhq5wLmSWXjyFVL+IDfOpdixJ6f5o+MhE6H4t31w+prygHmd2UHQ
# 657UGx6FNuzwC+SpAHmV76MZYac4uAhTgaP47P2eeS1ockvyhl9ya+9JzPfMkug3
# xevzFADWiLRMr066EMV7q3JSRAsnCS9GQ08C4FKPbSh8OPM33Lng0ffxANnHAAX/
# DE7cHcx7l9jaV3Acmkj7oqir4Eh2u5YxwiaTE37XaMumX2ES3PJ5NBaXq7YdLJwy
# SD+U9pk/tl4dQ1t/Eeo7uDTliOyQkD8I74xpVB0T31/67KHfkBkFVvy6wye21V+9
# IC8uSD++RgD3RwtN2kE/AgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUimLm8QMeJa25
# j9MWeabI2HSvZOUwHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYD
# VR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwG
# CCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIw
# MjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcD
# CDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAF/I8U6hbZhvDcn9
# 6nZ6tkbSEjXPvKZ6wroaXcgstEhpgaeEwleLuPXHLzEWtuJuYz4eshmhXqFr49lb
# AcX5SN5/cEsP0xdFayb7U5P94JZd3HjFvpWRNoNBhF3SDM0A38sI2H+hjhB/VfX1
# XcZiei1ROPAyCHcBgHLyQrEu6mnb3HhbIdr8h0Ta7WFylGhLSFW6wmzKusP6aOlm
# nGSac5NMfla6lRvTYHd28rbbCgfSm1RhTgoZj+W8DTKtiEMwubHJ3mIPKmo8xtJI
# WXPnXq6XKgldrL5cynLMX/0WX65OuWbHV5GTELdfWvGV3DaZrHPUQ/UP31Keqb2x
# jVCb30LVwgbjIvYS77N1dARkN8F/9pJ1gO4IvZWMwyMlKKFGojO1f1wbjSWcA/57
# tsc+t2blrMWgSNHgzDr01jbPSupRjy3Ht9ZZs4xN02eiX3eG297NrtC6l4c/gzn2
# 0eqoqWx/uHWxmTgB0F5osBuTHOe77DyEA0uhArGlgKP91jghgt/OVHoH65g0QqCt
# gZ+36mnCEg6IOhFoFrCc0fJFGVmb1+17gEe+HRMM7jBk4O06J+IooFrI3e3PJjPr
# Qano/MyE3h+zAuBWGMDRcUlNKCDU7dGnWvH3XWwLrCCIcz+3GwRUMsLsDdPW2OVv
# 7v1eEJiMSIZ2P+M7L20Q8aznU4OAMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJ
# mQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
# dGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1
# WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjK
# NVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQGOhg
# fWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJp
# rx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVWTe/d
# vI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka9
# 7aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De+JKR
# Hh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM9itu
# qBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyO
# ArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2ZItb
# oKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6
# bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q4i6t
# AgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQW
# BBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacb
# UzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYz
# aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnku
# aHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIA
# QwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2
# VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwu
# bWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYt
# MjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8qW/q
# XBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6
# U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2CDPVt
# I1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis
# 9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTp
# kbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0
# sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz138e
# W0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJ
# sWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7
# Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0
# dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQ
# tB1VM1izoXBm8qGCAtQwggI9AgEBMIIBAKGB2KSB1TCB0jELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxh
# bmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpE
# MDgyLTRCRkQtRUVCQTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZaIjCgEBMAcGBSsOAwIaAxUAdqNHe113gCJ87aZIGa5QBUqIwvKggYMwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIF
# AOir5/AwIhgPMjAyMzA5MTMxNTQ5MDRaGA8yMDIzMDkxNDE1NDkwNFowdDA6Bgor
# BgEEAYRZCgQBMSwwKjAKAgUA6Kvn8AIBADAHAgEAAgISODAHAgEAAgIUVjAKAgUA
# 6K05cAIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAID
# B6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAKqYWvGairHmNJb8qHTT
# tBO8BriSycKhsQ4nudDmAps8mMXSQmvgg11Hgc9pyzCFu9jITCYAaWJP2wd+Z0G+
# iiW7sR6Nz10nUu5c5bwtCW+slrDN3mRE1E2MKu+DVe9fyTLAh8JM/8AZn+7Gdz4q
# 1MXzs5RveL67SqYuagVd1ZmQMYIEDTCCBAkCAQEwgZMwfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAG6Hz8Z98F1vXwAAQAAAbowDQYJYIZIAWUDBAIB
# BQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQgjBP/c1ATEC0hAVXm/f7MPDaIZG+Wxqru2txgpCoIyKowgfoGCyqGSIb3DQEJ
# EAIvMYHqMIHnMIHkMIG9BCApVb08M25w+tYGWsmlGtp1gy1nPcqWfqgMF3nlWYVz
# BTCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABuh8/
# GffBdb18AAEAAAG6MCIEIBEBxapDzQFu22A9J15GXHZNk+LbiOoR/ilw1370fMr3
# MA0GCSqGSIb3DQEBCwUABIICAHMFhHuEDSQTzhIsoqAss+IViglbAXWztJg7gUlE
# v9oBCuWjXdxVGiJQa9/UVYL+yvk/qPMv8kYOlH0j3uFxRXDkj22FFAO1UjZiQLYp
# YWAWNePdp/QJzUBbe56Ok/zrT/CUmN/1/raYfIGuPQYaLf12XydWpNXsaR6LoxrJ
# wPEeVUWNTN/fGhxciYn8Mt43t4Z1ejEaPQu3PUMYCCt9lnNptUc2N3M1jmCYGpeq
# H1OK/wWTuWzWuH8xZc+P9ovXKSZn+ZH3uw/Xuowugv2c0fwebGvwCUgfhutXYYu4
# 8M25nic3uGdXd4Xj6SHbh1o3y21DQ+c7l0KArscm+8bJcOw5ld1V5rMz7Dn8Hpz8
# ep+BDoTltXEJQzlXhwHaEZgyV5GUxIALqxkACoumG8S0yLvtwdTzE/Ht0YO2NK4W
# ZyT+fhGE+/f/lRJSSKxvab+tQna2vZNBuQWJYdzHDIdlh8Wc1QY7EaSandqNwvqx
# qT9ys2tQRpTSC/9bFN92dc0fSTLRxj3NJB9YaHbta/J2nK1Q+RWS0UzCcDuyQTH4
# dpUF9HOXBJuxFeH5MwnMs2hxUeW7CerMBcHkG279wcyBTwm1laN4I+exUPApvW6q
# E3IiNOpP6WG4RIZEMBTxjZU/zlWYD/NApyhH+iDn0Bv1Mt7gyWDFOm4aK3dG//O1
# k713
# SIG # End signature block
