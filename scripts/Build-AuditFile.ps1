# Build-AuditFile.ps1
# Converts a Purview-exported CSV into the same format produced by Audit-Get-Events.ps1

param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "Copilot_Events.csv"
)

if (-not (Test-Path $InputFile -PathType Leaf)) {
    Write-Host "Input file not found: $InputFile" -ForegroundColor Red
    exit
}

$Records = Import-Csv -Path $InputFile

if ($Records.Count -eq 0) {
    Write-Host "No records found in input file." -ForegroundColor Red
    exit
}

# Make sure that everything is sorted in date order
$Records = $Records | Sort-Object {$_.CreationDate -as [datetime]}

Write-Host ("{0} Copilot audit records found. Now analyzing the content" -f $Records.Count)

$Report = [System.Collections.Generic.List[Object]]::new()

ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    $CopilotApp = 'Copilot for M365'; $Context = $null; $CopilotLocation = $null

    Switch ($AuditData.copiloteventdata.contexts.type) {
        "xlsx" {
            $CopilotApp = "Excel"
        }
        "docx" {
            $CopilotApp = "Word"
        }
        "pptx" {
            $CopilotApp = "PowerPoint"
        }
        "TeamsMeeting" {
            $CopilotApp = "Teams"
            $CopilotLocation = "Teams meeting"
        }
        "whiteboard" {
            $CopilotApp = "Whiteboard"
        }
        "loop" {
            $CopilotApp = "Loop"
        }
        "StreamVideo" {
            $CopilotApp = "Stream"
            $CopilotLocation = "Stream video player"
        }
    }

    If ($AuditData.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotApp = "Teams"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "bizchat" -or $AuditData.AppIdentity -like "*Bizchat*") {
        $CopilotApp = "Copilot for M365 Chat"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "Outlook") {
        $CopilotApp = "Outlook"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "Copilot Studio") {
        $CopilotApp = "Copilot Studio Agent"
    }

    If ($AuditData.copiloteventdata.contexts.id) {
        $Context = $AuditData.copiloteventdata.contexts.id
    } ElseIf ($AuditData.copiloteventdata.threadid) {
        $Context = $AuditData.copiloteventdata.threadid
    }

    If ($AuditData.copiloteventdata.contexts.id -like "*/sites/*") {
        $CopilotLocation = "SharePoint Online"
    } ElseIf ($AuditData.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotLocation = "Teams"
        If ($AuditData.copiloteventdata.contexts.id -like "*ctx=channel*") {
            $CopilotLocation = "Teams Channel"
        } Else {
            $CopilotLocation = "Teams Chat"
        }
    } ElseIf ($AuditData.copiloteventdata.contexts.id -like "*/personal/*") {
        $CopilotLocation = "OneDrive for Business"
    }

    [array]$AccessedResources = $AuditData.copiloteventdata.accessedResources.name | Sort-Object -Unique
    [string]$AccessedResources = $AccessedResources -join ", "
    [array]$AccessedResourceLocations = $AuditData.copiloteventdata.accessedResources.id | Sort-Object -Unique
    [string]$AccessedResourceLocations = $AccessedResourceLocations -join ", "
    [array]$AccessedResourceActions = $AuditData.copiloteventdata.accessedResources.action | Sort-Object -Unique
    [string]$AccessedResourceActions = $AccessedResourceActions -join ", "

    # Use CreationDate from AuditData JSON, fall back to the CSV's CreationDate column
    $Timestamp = $null
    if ($AuditData.CreationTime) {
        $Timestamp = Get-Date $AuditData.CreationTime -Format "dd-MMM-yyyy HH:mm:ss"
    } elseif ($Rec.CreationDate) {
        $Timestamp = Get-Date $Rec.CreationDate -Format "dd-MMM-yyyy HH:mm:ss"
    }

    # Use UserId from AuditData JSON, fall back to the CSV's UserId column
    $User = if ($AuditData.UserId) { $AuditData.UserId } else { $Rec.UserId }

    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp                     = $Timestamp
        User                          = $User
        App                           = $CopilotApp
        Location                      = $CopilotLocation
        'App context'                 = $Context
        'Accessed Resources'          = $AccessedResources
        'Accessed Resource Locations' = $AccessedResourceLocations
        Action                        = $AccessedResourceActions
    }
    $Report.Add($ReportLine)
}

$OutputDir = Split-Path -Path $OutputFile -Parent
if ($OutputDir -and -not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$Report | Export-Csv -Path $OutputFile -NoTypeInformation
Write-Host "Successfully processed $($Report.Count) records. Output saved to $OutputFile" -ForegroundColor Green