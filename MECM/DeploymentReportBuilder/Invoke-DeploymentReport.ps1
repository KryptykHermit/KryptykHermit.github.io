<#
    .SYNOPSIS
    Create an SCCM report with Success, Error, In Progress, and Unknown collection results

    .DESCRIPTION
    Creates an SCCM report based off of a selected deployment and exports the result to CSV or XLSX format (if module is installed)
    Places the result file on the desktop for easy consumption

    .PARAMETER MECMServer
    Specifies the name of the MECM (Microsoft Endpoint Configuration Manager) Primary Site Server

    .PARAMETER Namespace
    Specifies the name of the MECM CIM site path

    .INPUTS

    .OUTPUTS

    .LINK
    Online version: https://kryptykhermit.github.io/MECM/DeploymentReportBuilder/Invoke-ReportBuilder.html
#>
[string]$MECMServer   = 'sccm.home.lab'
[string]$namespace    = 'root\sms\site_P00'

# Using the registry hive in case this is launched from VDI, where the paths do not translate correctly
[string]$desktopPath = 'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
######################################
# LOOKUP TABLES
######################################
$appStatusType = @{
    [uint32]1 = 'Success'
    [uint32]2 = 'In Progress'
    [uint32]3 = 'Requirements Not Met'
    [uint32]4 = 'Unknown'
    [uint32]5 = 'Error'
}
$collectionType = @{
    [uint32]1 = 'User'
    [uint32]2 = 'Computer'
}
$deploymentIntent = @{
    [uint32]1 = 'Required'
    [uint32]2 = 'Available'
    [uint32]3 = 'Simulate'
}
$desiredConfigType = @{
    [int32]1 = 'Install'
    [int32]2 = 'Uninstall'
    [int32]3 = 'Monitor'
    [int32]4 = 'Remediate'
}
$enforcementState = @{
    [int32]0  = 'Enforcement State Unknown'
    [int32]1  = 'Enforcement started'
    [int32]2  = 'Enforcement waiting for content'
    [int32]3  = 'Waiting for another installation to complete'
    [int32]4  = 'Waiting for maintenance window before installing'
    [int32]5  = 'Restart required before installing'
    [int32]6  = 'General failure'
    [int32]7  = 'Pending installation'
    [int32]8  = 'Installing update'
    [int32]9  = 'Pending system restart'
    [int32]10 = 'Successfully installed update'
    [int32]11 = 'Failed to install update'
    [int32]12 = 'Downloading update'
    [int32]13 = 'Downloaded update'
    [int32]14 = 'Failed to download update'
}
$enforcementState2 = @{
    [uint32]1000 = 'Success'
    [uint32]1001 = 'Already compliant'
    [uint32]1002 = 'Simulate success'
    [uint32]1003 = 'Fast status success'
    [uint32]2000 = 'In progress'
    [uint32]2001 = 'Waiting for content'
    [uint32]2002 = 'Installing'
    [uint32]2003 = 'Restart to continue'
    [uint32]2004 = 'Waiting for maintenance window'
    [uint32]2005 = 'Waiting for schedule'
    [uint32]2006 = 'Downloading dependent content'
    [uint32]2007 = 'Installing dependent content'
    [uint32]2008 = 'Restart to complete'
    [uint32]2009 = 'Content downloaded'
    [uint32]2010 = 'Waiting for update'
    [uint32]2011 = 'Waiting for user session reconnect'
    [uint32]2012 = 'Waiting for user logoff'
    [uint32]2013 = 'Waiting for user logon'
    [uint32]2014 = 'Waiting To Install'
    [uint32]2015 = 'Waiting Retry'
    [uint32]2016 = 'Waiting For Presentation Mode'
    [uint32]2017 = 'Waiting For Orchestration'
    [uint32]2018 = 'Waiting For Network'
    [uint32]2019 = 'Pending App-V Virtual Environment Update'
    [uint32]2020 = 'Updating App-V Virtual Environment'
    [uint32]3000 = 'Requirements not met'
    [uint32]3001 = 'Host Platform Not Applicable'
    [uint32]4000 = 'Unknown'
    [uint32]5000 = 'Deployment failed'
    [uint32]5001 = 'Evaluation failed'
    [uint32]5002 = 'Deployment failed'
    [uint32]5003 = 'Failed to locate content'
    [uint32]5004 = 'Dependency installation failed'
    [uint32]5005 = 'Failed to download dependent content'
    [uint32]5006 = 'Conflicts with another application deployment'
    [uint32]5007 = 'Waiting Retry'
    [uint32]5008 = 'Failed to uninstall superseded deployment type'
    [uint32]5009 = 'Failed to download superseded deployment type'
    [uint32]5010 = 'Error updating VE'
    [uint32]5011 = 'Error installinge license'
    [uint32]5012 = 'Error retrieving allow all trusted apps'
    [uint32]5013 = 'Error no licenses available'
    [uint32]5014 = 'Error OS not supported'
    [uint32]6000 = 'Launch succeeded'
    [uint32]6010 = 'Launch error'
    [uint32]6020 = 'Launch unknown'
}
$featureType = @{
    [uint32]1  = 'Application'
    [uint32]2  = 'Program'
    [uint32]3  = 'MobileProgram'
    [uint32]4  = 'Script'
    [uint32]5  = 'SoftwareUpdate'
    [uint32]6  = 'Baseline'
    [uint32]7  = 'TaskSequence'
    [uint32]8  = 'ContentDistribution'
    [uint32]9  = 'DistributionPointGroup'
    [uint32]10 = 'DistributionPointHealth'
    [uint32]11 = 'ConfigurationPolicy'
    [uint32]28 = 'AbstractConfigurationItem'
}
$installedState = @{
    [uint32]1 = 'Uninstall'
    [uint32]2 = 'Install'
    [uint32]3 = 'Unknown'
}
$resourceType = @{
    [UInt32]3 = 'User Group'
    [UInt32]4 = 'User'
    [UInt32]5 = 'System'
}
######################################
# FUNCTIONS
######################################
# Convert length in number to alphabet
function NumberToAlphabet {
    param(
        [int]$n
    )
    if ($n -gt 26) { $n = 26 }
    ([char](64+$n),(+$n-64))[$n -ge 65]
}
######################################
# Get application deployment details
Write-Host
Write-Host "Acquiring Application Deployments" -ForegroundColor 'Yellow'
[string]$cimClassName = 'SMS_DeploymentSummary'
$appDeployments = Get-CimInstance -ClassName $cimClassName -Namespace $namespace -ComputerName $MECMServer
Write-Host "--- Detected : $($appDeployments.Count) Application Deployments" -ForegroundColor 'Green'

# Prompt for application selection
[Array]$AssignmentID = $appDeployments |
    Where-Object FeatureType -match '[12]' |
    Select-Object -Property 'ApplicationName', 'CollectionName', 'CollectionID', `
                            @{name='CollectionType';e={$CollectionType[$_.CollectionType]}}, `
                            @{name='DeploymentIntent';e={$deploymentIntent[$_.DeploymentIntent]}}, `
                            @{name='DesiredConfigType';e={$DesiredConfigType[$_.DesiredConfigType]}}, `
                            @{name='FeatureType';e={$FeatureType[$_.FeatureType]}}, `
                            @{name='Targeted';e={$_.NumberTargeted}}, `
                            @{name='Success';e={$_.NumberSuccess}}, `
                            @{name='InProgress';e={$_.NumberInProgress}}, `
                            @{name='Error';e={$_.NumberErrors}}, `
                            @{name='Other';e={$_.NumberOther}}, `
                            @{name='EnforcementDeadline';e={Get-Date $_.EnforcementDeadline -format 'yyyy-MM-dd HH:mm:ss'}} |
    Sort-Object -Property 'ApplicationName' |
    Out-GridView -Title "Select a deployment to query" -PassThru
Write-Host "--- Selected : $($AssignmentID.ApplicationName)" -ForegroundColor 'Green'

# Set the log file name
[string]$LogFile = (Get-ItemProperty -Path $desktopPath -Name 'Desktop').Desktop + "\$($AssignmentID.ApplicationName).csv"
Write-Host "--- Creating : Log file '$LogFile'" -ForegroundColor 'Green'
Write-Host

# Status Details per object
Write-Host "Acquiring Deployment Statuses" -ForegroundColor 'Yellow'
[string]$cimClassName = 'SMS_AppDeploymentAssetDetails'
$DeploymentDevices = @(Get-CimInstance -ClassName $cimClassName -Namespace $namespace -ComputerName $MECMServer -Filter "AppName='$($AssignmentID.ApplicationName)' and CollectionID='$($AssignmentID.CollectionID)'" |
    Select-Object -Property @{n='AppStatusType';e={$appStatusType[$_.AppStatusType]}}, `
                            'CollectionID', 'CollectionName', `
                            @{n='DeploymentIntent';e={$deploymentIntent[$_.DeploymentIntent]}}, `
                            @{n='InstalledType';e={$installedState[$_.InstalledState]}}, `             # Install, Uninstall, Unknown
                            'MachineID', 'MachineName', `
                            @{n='StatusType';e={$appStatusType[$_.StatusType]}}, `                     # Uses same IDs as AppStatusType
                            @{n='EnforcementState';e={$enforcementState2[$_.EnforcementState]}})
Write-Host "--- Detected : $($DeploymentDevices.Count) Installation Statuses" -ForegroundColor 'Green'

# Get each object and append status details
[string]$cimClassName = 'SMS_CM_RES_COLL_' + $DeploymentDevices[0].CollectionID
Write-Host "--- Detected : Deployment Collection Name is '$cimClassName'" -ForegroundColor 'Green'
$deploymentCollection = @(Get-CimInstance -ClassName $cimClassName -Namespace $namespace -ComputerName $MECMServer |
    Select-Object -Property 'IsClient', 'Name', 'ResourceID', 'ResourceType', 'ADLastLogonTime', 'BoundaryGroups', 'ClientVersion', 'LastMPServerName', 'LastActiveTime', `
                            'LastClientCheckTime', 'LastHardwareScan', 'LastPolicyRequest', 'LastSoftwareScan', 'LastStatusMessage')
Write-Host "--- Detected : $($deploymentCollection.Count) $($AssignmentID[0].CollectionType) Objects" -ForegroundColor 'Green'
Write-Host

Write-Host "Compiling Report Results" -ForegroundColor 'Yellow'
$reportInfo = [System.Collections.ArrayList]::new()
Write-Host "--- Compiling Results" -ForegroundColor 'Green'
$deploymentCollection |
    ForEach-Object {
        # if only 1 object in the deployment, use the whole collection for results
        if ($DeploymentDevices.Count -eq 1) {
            $obj = $DeploymentDevices
        }
        else {
            $obj = $DeploymentDevices[$DeploymentDevices.MachineID.IndexOf($_.ResourceID)]
        }

        if ($obj.MachineName -eq $_.Name) {
            $null = $ReportInfo.Add(
                [pscustomobject]@{
                    IsClient            = $_.IsClient
                    Name                = $_.Name
                    ResourceID          = $_.ResourceID
                    ResourceType        = $resourceType[$_.ResourceType]
                    InstallStatus       = $obj.AppStatusType
                    EnforcementState    = $obj.EnforcementState
                    DeploymentType      = $obj.DeploymentIntent
                    StatusType          = $obj.StatusType
                    Application         = $AssignmentID.ApplicationName
                    CollectionName      = $AssignmentID.CollectionName
                    ADLastLogonTime     = $(if ($null -eq $_.ADLastLogonTime) {$null} else {Get-Date $_.ADLastLogonTime -Format 'yyyy-MM-dd HH:mm:ss'})
                    BoundaryGroups      = $_.BoundaryGroups
                    ClientVersion       = $_.ClientVersion
                    ManagementPoint     = $_.LastMPServerName
                    LastActiveTime      = $(if ($null -eq $_.LastActiveTime) {$null} else {Get-Date $_.LastActiveTime -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastHardwareScan    = $(if ($null -eq $_.LastHardwareScan) {$null} else {Get-Date $_.LastHardwareScan -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastSoftwareScan    = $(if ($null -eq $_.LastSoftwareScan) {$null} else {Get-Date $_.LastSoftwareScan -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastPolicyRequest   = $(if ($null -eq $_.LastPolicyRequest) {$null} else {Get-Date $_.LastPolicyRequest -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastStatusMessage   = $(if ($null -eq $_.LastStatusMessage) {$null} else {Get-Date $_.LastStatusMessage -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastClientCheckTime = $(if ($null -eq $_.LastClientCheckTime) {$null} else {Get-Date $_.LastClientCheckTime -Format 'yyyy-MM-dd HH:mm:ss'})
                }
            )
        }
        else {
            $null = $ReportInfo.Add(
                [pscustomobject]@{
                    IsClient            = $_.IsClient
                    Name                = $_.Name
                    ResourceID          = $_.ResourceID
                    ResourceType        = $resourceType[$_.ResourceType]
                    InstallStatus       = 'Unknown'
                    EnforcementState    = $null
                    DeploymentType      = $null
                    StatusType          = $null
                    Application         = $AssignmentID.ApplicationName
                    CollectionName      = $AssignmentID.CollectionName
                    ADLastLogonTime     = $(if ($null -eq $_.ADLastLogonTime) {$null} else {Get-Date $_.ADLastLogonTime -Format 'yyyy-MM-dd HH:mm:ss'})
                    BoundaryGroups      = $_.BoundaryGroups
                    ClientVersion       = $_.ClientVersion
                    ManagementPoint     = $_.LastMPServerName
                    LastActiveTime      = $(if ($null -eq $_.LastActiveTime) {$null} else {Get-Date $_.LastActiveTime -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastHardwareScan    = $(if ($null -eq $_.LastHardwareScan) {$null} else {Get-Date $_.LastHardwareScan -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastSoftwareScan    = $(if ($null -eq $_.LastSoftwareScan) {$null} else {Get-Date $_.LastSoftwareScan -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastPolicyRequest   = $(if ($null -eq $_.LastPolicyRequest) {$null} else {Get-Date $_.LastPolicyRequest -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastStatusMessage   = $(if ($null -eq $_.LastStatusMessage) {$null} else {Get-Date $_.LastStatusMessage -Format 'yyyy-MM-dd HH:mm:ss'})
                    LastClientCheckTime = $(if ($null -eq $_.LastClientCheckTime) {$null} else {Get-Date $_.LastClientCheckTime -Format 'yyyy-MM-dd HH:mm:ss'})
                }
            )
        }
    }

# REPORT PROCESSING
try {
    Write-Host "--- Adding ImportExcel Module" -ForegroundColor 'Green'
    Import-Module -Name 'ImportExcel' -ErrorAction 'Stop'

    # Create Excel based log file name
    Write-Host "--- Mapping Output File: $($LogFile.Replace('.csv','.xlsx'))" -ForegroundColor 'Green'
    [string]$ExcelLogFile = $LogFile.Replace('.csv','.xlsx')

    # Remove any old file if found
    Write-Host "--- Removing Old Excel File" -ForegroundColor 'Green'
    if (Test-Path -Path $ExcelLogFile) { $null = Remove-Item -Path $ExcelLogFile -Force }

    # Max Length of table headers
    Write-Host "--- Processing Column Count" -ForegroundColor 'Green'
    $titleLength = NumberToAlphabet -n ($reportInfo | Get-Member -MemberType:NoteProperty).Count

    # Create new Excel report file
    Write-Host "--- Creating Conditional Text Coloring Schema" -ForegroundColor 'Green'
    $text1 = New-ConditionalText -Text 'Success' -BackgroundColor 'Green' -ConditionalTextColor 'White' -Range E:E
    $text2 = New-ConditionalText -Text 'In Progress' -BackgroundColor 'Yellow' -ConditionalTextColor 'Black' -Range E:E
    $text3 = New-ConditionalText -Text 'Error' -BackgroundColor 'Red' -ConditionalTextColor 'White' -Range E:E
    $text4 = New-ConditionalText -Text 'Unknown' -BackgroundColor 'Black' -ConditionalTextColor 'White' -Range E:E
    $excelArgs = @{
        Path                 = $ExcelLogFile
        Title                = "Application: $($reportInfo[0].Application)"
        TableStyle           = 'Medium16'
        TitleBackgroundColor = 'Orange'
        AutoSize             = $true
        AutoFilter           = $true
        FreezeTopRow         = $false
        Show                 = $false
        ConditionalText      = @($text1, $text2, $text3, $text4)
    }
    Write-Host "--- Creating Excel Workbook" -ForegroundColor 'Green'
    $xl = $ReportInfo |
        Select-Object -Property * -ExcludeProperty 'Application' |
        Sort-Object -Property 'InstallStatus', 'Name' |
        Export-Excel @excelArgs -PassThru
    # Add the Title into the worksheet
    Write-Host "--- Adding Title" -ForegroundColor 'Green'
    $sheet1 = $xl.Workbook.Worksheets['Sheet1']
    $sheet1.Cells["A1:${titleLength}1"].merge = $true
    # Freeze the top 2 rows, not just the first...
    Write-Host "--- Freezing Columns" -ForegroundColor 'Green'
    $sheet1.View.FreezePanes(3,1)

    Close-ExcelPackage $xl -Show
}
catch {
    Write-Host "--- Creating CSV File" -ForegroundColor 'Green'
    $ReportInfo |
        Sort-Object -Property 'InstallStatus', 'Name' |
        Export-CSV -Path $LogFile -NoTypeInformation -Force
    Write-Host "Please install the ImportExcel module to report results to Excel format" -ForegroundColor 'Cyan'
    Write-Host "PS> " -ForegroundColor 'White' -NoNewline
    Write-Host 'Install-Module ' -ForegroundColor 'Yellow' -NoNewline
    Write-Host '-Name' -ForegroundColor 'Gray' -NoNewline
    Write-Host " 'ImportExcel' " -ForegroundColor 'Blue' -NoNewline
    Write-Host '-Scope:' -ForegroundColor 'Gray' -NoNewline
    Write-Host 'CurrentUser' -ForegroundColor 'White' -NoNewline
    Write-Host ' -Force' -ForegroundColor 'Gray' -NoNewline
}
Write-Host
Write-Host " -= Processing Complete! =-" -ForegroundColor 'Yellow'
Start-Sleep -Seconds 5
