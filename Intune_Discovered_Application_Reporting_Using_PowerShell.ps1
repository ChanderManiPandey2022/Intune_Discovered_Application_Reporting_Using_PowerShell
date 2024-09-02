# <Create Intune Discovered Appllication Report using  PowerShell>
#.DESCRIPTION
 # <Create Intune Discovered Appllication Report using  PowerShell>
#.Demo
#<YouTube video link-->https://www.youtube.com/@ChanderManiPandey
#.INPUTS
 # <Provide all required inforamtion in User Input Section-line No 24>
#.OUTPUTS
 # <This will Create Intune Discovered Application Report using PowerShell>

#.NOTES
 
 <#
  Version:         1.1
  Author:          Chander Mani Pandey
  Creation Date:   30 Aug 2024
  
  Find Author on 
  Youtube:-        https://www.youtube.com/@chandermanipandey8763
  Twitter:-        https://twitter.com/Mani_CMPandey
  LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
 #>
#=================================================================================================================================================
#------------------------------------------------------ User Input Section Start------------------------------------------------------------------
#=================================================================================================================================================

$WorkingFolder   = "C:\TEMP\Intune_Discovered_Apps"           # Application reporting folder Location
$AppName         = "Google Chrome"                            # Enter the application name e.g- "Google Chrome"
$Platform        = "Windows"                                  # Enter the $Platform e.g-Windows,AndroidFullyManagedDedicated,Other,AndroidWorkProfile,IOS,AndroidDeviceAdministrator
$FilterOperator  = "eq"                                       # Choose filter operator: 'like' or 'eq'
$tenantNameOrID  = "ABC.com"                                  # Tenant Name or ID
$clientAppId     = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"           # Client Application ID
$clientAppSecret = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"  # Client Application Secret

#=================================================================================================================================================
#------------------------------------------------------ User Input Section End--------------------------------------------------------------------
#=================================================================================================================================================

Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
$error.Clear() # Clear error history
cls
$ErrorActionPreference = 'SilentlyContinue'

#================================================================================================================================================
$overallStartTime = Get-Date
Write-Host "=================================== Creating $AppName Application Report for $Platform Platform =====================================" -ForegroundColor Magenta
# Validate the filter operator input
if ($FilterOperator -notin @('like', 'eq')) {
    Write-Host "Invalid filter operator. Please enter 'like' or 'eq'." -ForegroundColor Red
    exit
}

# Define folder path based on $AppName and $Platform
$AppFolderPath = "$WorkingFolder\$AppName`_$Platform"
$ReportName = "AppInvRawData"
$Path = "$AppFolderPath\$ReportName\"

# Start overall time tracking
# Check if the $AppFolderPath folder exists, and delete it if it does

if (Test-Path -Path $AppFolderPath) {
    Write-Host " "
        Write-Host " "
    Write-Host "$AppName Application Folder present. Removing existing folder: $AppFolderPath" -ForegroundColor Cyan
    Remove-Item -Path $AppFolderPath -Recurse -Force | Out-Null
    Write-Host "" -ForegroundColor Yellow
}

# Create the new working folder and subfolder
Write-Host " "
Write-Host "Creating working folder and subfolder at $AppFolderPath & $Path" -ForegroundColor Yellow
New-Item -ItemType Directory -Path $AppFolderPath -Force -InformationAction SilentlyContinue | Out-Null
New-Item -ItemType Directory -Path $Path -Force -InformationAction SilentlyContinue | Out-Null
Write-Host "Created working folder and subfolder at $AppFolderPath & $Path" -ForegroundColor Green

#=================================================================================================================================================
# Check and install the Microsoft.Graph.Intune module
Write-Host ""
Write-Host "Checking for Microsoft.Graph.Intune module" -ForegroundColor Cyan
if (-not (Get-Module -Name "Microsoft.Graph.Intune" -ListAvailable)) {
    Write-Host "Installing Microsoft.Graph.Intune module" -ForegroundColor Cyan
    Install-Module -Name Microsoft.Graph.Intune -Force
}
Write-Host "Importing Microsoft.Graph.Intune module" -ForegroundColor Cyan
Import-Module Microsoft.Graph.Intune -Force -InformationAction SilentlyContinue

# Authentication
Write-Host "Setting up authentication for MS Graph" -ForegroundColor Cyan

$tenant = $tenantNameOrID
$authority = "https://login.windows.net/$tenant"
$clientId = $clientAppId
$clientSecret = $clientAppSecret

Update-MSGraphEnvironment -AppId $clientId -AuthUrl $authority -SchemaVersion "Beta" -Quiet -InformationAction SilentlyContinue
Connect-MSGraph -ClientSecret $clientSecret -InformationAction SilentlyContinue -Quiet

#=================================================================================================================================================
# Create request body and initiate export job
Write-Host "" 
Write-Host "Initiating export job for '$AppName' application and for '$Platform' Platform" -ForegroundColor Yellow
$exportJobStartTime = Get-Date
$postBody = @{ 
    'reportName' = $ReportName 
    'search' = $AppName
}
$exportJob = Invoke-MSGraphRequest -HttpMethod POST -Url "DeviceManagement/reports/exportJobs" -Content $postBody
$exportJobEndTime = Get-Date
$exportJobDuration = $exportJobEndTime - $exportJobStartTime
Write-Host "Export Job initiated. Monitoring Downloading status..." -ForegroundColor Cyan
Write-Host ""

# Polling for export job status
$dotCount = 0
$pollingStartTime = Get-Date
do {
    Start-Sleep -Seconds 2
    $exportJob = Invoke-MSGraphRequest -HttpMethod Get -Url "DeviceManagement/reports/exportJobs('$($exportJob.id)')" -InformationAction SilentlyContinue
    Write-Host -NoNewline '.'
    $dotCount++
    if ($dotCount -eq 100) {
        Write-Host ""
        $dotCount = 0
    }
} while ($exportJob.status -eq 'inprogress')

$pollingEndTime = Get-Date
$pollingDuration = $pollingEndTime - $pollingStartTime
Write-Host ""

if ($exportJob.status -eq 'completed') {
    $fileName = (Split-Path -Path $exportJob.url -Leaf).Split('?')[0]
    Write-Host ""
    Write-Host "Export Job completed. Writing File $fileName to Disk..." -ForegroundColor Cyan
    $downloadStartTime = Get-Date
    Invoke-WebRequest -Uri $exportJob.url -Method Get -OutFile "$Path$fileName"
    $downloadEndTime = Get-Date
    $downloadDuration = $downloadEndTime - $pollingStartTime
    Write-Host "Time taken to Initiate and download $AppName file: $([math]::Round($downloadDuration.TotalMinutes, 2)) minutes" -ForegroundColor White

    Remove-Item -Path "$Path*" -Include *.csv -Force
    Expand-Archive -Path "$Path$fileName" -DestinationPath $Path
    #=============================================================================================================================================
    # Processing CSV file
    Write-Host ""
    Write-Host "Processing CSV file. Please wait..." -ForegroundColor Cyan

    # Get the path of the CSV file
    $csvPath = Get-ChildItem -Path $Path -Filter *.csv | Where-Object {! $_.PSIsContainer} | Select-Object -ExpandProperty FullName

    # Check if the CSV file exists
    if (-not (Test-Path $csvPath)) {
        Write-Host "CSV file not found at $csvPath" -ForegroundColor Red
        exit
    }

    $overallCSVImportStartTime = Get-Date
    # Read the CSV file
    $csvData = Import-Csv -Path $csvPath

    $overallCSVImportEndTime = Get-Date
    $overallImportDuration = $overallCSVImportEndTime - $overallCSVImportStartTime
    Write-Host "Time taken to Import CSV: $([math]::Round($overallImportDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    Write-Host ""

    # Filter data based on user choice and Platform
    if ($FilterOperator -eq 'like') {
        $filteredData = $csvData | Where-Object {
            $_.ApplicationName -like "*$AppName*" -and $_.Platform -eq $Platform
        }
    } elseif ($FilterOperator -eq 'eq') {
        $filteredData = $csvData | Where-Object {
            $_.ApplicationName -eq $AppName -and $_.Platform -eq $Platform
        }
    }

    $processingStartTime = Get-Date
    # Write filtered data to CSV
    $filteredOutputPath = "$AppFolderPath\Filtered_$($AppName)_$($Platform).csv"
    $filteredData | Export-Csv -Path $filteredOutputPath -NoTypeInformation -Encoding utf8

    # End time for processing CSV
    $processingEndTime = Get-Date
    $processingDuration = $processingEndTime - $processingStartTime
    Write-Host "Processing complete. Filtered data saved to $filteredOutputPath" -ForegroundColor Cyan
    Write-Host "Time taken to process CSV: $([math]::Round($processingDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    Write-Host ""

    # Import and summarize the final report
    Write-Host "Summarizing final report for $AppName..." -ForegroundColor Cyan
    Write-Host ""
    $finalizeStartTime = Get-Date
    $FinalReport = Import-Csv -Path $filteredOutputPath
    $TotalDevices = ($FinalReport | Measure-Object | Select-Object -ExpandProperty Count)
    $TotalApplicationVersions = ($FinalReport | Select-Object -ExpandProperty ApplicationVersion -Unique | Measure-Object | Select-Object -ExpandProperty Count)

    Write-Host "Total Number of $Platform Devices where $AppName is Installed:          $TotalDevices" -ForegroundColor Yellow
    Write-Host "Total Number of Detected Application Versions on $Platform Platform:         $TotalApplicationVersions" -ForegroundColor Yellow
    Write-Host ""

    # Format and export the final report with dynamic filename
    Write-Host "Formatting and exporting final report..." -ForegroundColor Cyan
    $formattedReportStartTime = Get-Date
    $formattedReportPath = "$AppFolderPath\$($AppName)_Report.csv"
    $FormattedReport = $FinalReport | Select-Object DeviceName, UserName, EmailAddress, OSDescription, OSVersion, Platform, ApplicationName, ApplicationVersion, ApplicationPublisher
    $FormattedReport | Export-Csv -Path $formattedReportPath -NoTypeInformation

    # End time for finalizing report
    $formattedReportEndTime = Get-Date
    $formattedReportDuration = $formattedReportEndTime - $formattedReportStartTime
    Write-Host "Time taken to finalize report: $([math]::Round($formattedReportDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
    Write-Host ""

    # End overall time tracking
    $overallEndTime = Get-Date
    $overallDuration = $overallEndTime - $overallStartTime
    Write-Host "Overall time to complete the script: $([math]::Round($overallDuration.TotalMinutes, 2)) minutes" -ForegroundColor Green
}

#=================================================================================================================================================
# Define the list of paths to clean up

$pathsToRemove = @($filteredOutputPath, $Path)
Write-Host "Performing Cleanup work" -ForegroundColor Cyan
foreach ($path in $pathsToRemove) {
    if (Test-Path -Path $path) {
        Remove-Item -Path $path -Recurse -Force
    }
}
Write-Host ""
Write-Host "Final report saved location: $formattedReportPath" -ForegroundColor Yellow
Write-Host "Overall time to complete the script: $([math]::Round($overallDuration.TotalMinutes, 2)) minutes" -ForegroundColor White
Write-Host ""
Write-Host "============================= Successfully created $AppName Application Report for $Platform Platform ============================" -ForegroundColor Magenta
# Optionally output the formatted report to the console
#$FormattedReport | Out-GridView
#=================================================================================================================================================
