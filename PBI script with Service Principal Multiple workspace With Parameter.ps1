# Load required assemblies
try {
    Add-Type -Path "C:\Program Files\Microsoft.NET\ADOMD.NET\160\Microsoft.AnalysisServices.AdomdClient.dll"
    Add-Type -Path "C:\Program Files (x86)\Tabular Editor\Microsoft.AnalysisServices.Tabular.dll"
} catch {
    Write-Error "Failed to load required assemblies. Please verify paths."
    exit
}

# Get the folder where the script resides. Prefer $PSScriptRoot (works in script scope),
# fallback to MyInvocation for older hosts. This prevents using the PowerShell exe folder (System32).
if ($PSCommandPath) {
    $scriptPath = Split-Path -Parent $PSCommandPath
} elseif ($PSScriptRoot) {
    $scriptPath = $PSScriptRoot
} else {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

$configPath = Join-Path $scriptPath "config.txt"

if (-Not (Test-Path $configPath)) {
    @"
tenantId=
clientId=
clientSecret=
"@ | Set-Content $configPath
    Write-Host "`n⚠️ 'config.txt' not found. A blank template has been created at:`n$configPath"
    Write-Host "Please update the file and run again."
    Start-Sleep -Seconds 30
    exit
}

# Load config.txt
$config = @{}
Get-Content $configPath | ForEach-Object {
    if ($_ -match '^\s*([^=]+)\s*=\s*(.*?)\s*$') {
        $key = $matches[1].Trim()
        $value = $matches[2].Trim()
        $config[$key] = $value
    }
}

# Validate required fields
$tenantId     = $config["tenantId"]
$clientId     = $config["clientId"]
$clientSecret = $config["clientSecret"]

if (-not $tenantId -or -not $clientId -or -not $clientSecret) {
    Write-Host "`n❌ One or more required parameters are missing in 'config.txt'. Please check and try again."
    Start-Sleep -Seconds 30
    exit
}

# Output configuration: place output CSV in the same folder as this script (project folder)
$outputPath = Join-Path -Path $scriptPath -ChildPath "PowerBI_Metadata_Extract_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Initialize data collection
$metadata = [System.Collections.Generic.List[object]]::new()

# Function to get auth token
function Get-PowerBIAccessToken {
    param (
        [string]$tenantId,
        [string]$clientId,
        [string]$clientSecret
    )

    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        resource      = "https://analysis.windows.net/powerbi/api"
    }

    try {
        $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $body
        return $response.access_token
    } catch {
        Write-Error "Failed to get access token: $_"
        exit
    }
}

# Function to get accessible workspaces
function Get-PowerBIWorkspaces {
    param (
        [string]$accessToken
    )

    $headers = @{
        "Authorization" = "Bearer $accessToken"
    }

    try {
        $workspaces = @()
        $url = "https://api.powerbi.com/v1.0/myorg/groups"
        
        do {
            $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
            $workspaces += $response.value
            $url = $response.'@odata.nextLink'
        } while ($null -ne $url)

        return $workspaces
    } catch {
        Write-Error "Failed to retrieve workspaces: $_"
        exit
    }
}

# Main execution
try {
    $accessToken = Get-PowerBIAccessToken -tenantId $tenantId -clientId $clientId -clientSecret $clientSecret

    Write-Host "Discovering accessible workspaces..." -ForegroundColor Yellow
    $workspaces = Get-PowerBIWorkspaces -accessToken $accessToken

    if ($workspaces.Count -eq 0) {
        Write-Host "No workspaces found for this service principal." -ForegroundColor Red
        Start-Sleep -Seconds 30
        exit
    }

    Write-Host "Found $($workspaces.Count) accessible workspaces:" -ForegroundColor Green
    $workspaces | ForEach-Object { Write-Host " - $($_.name) (ID: $($_.id))" -ForegroundColor Cyan }

    foreach ($workspace in $workspaces) {
        $workspaceName = $workspace.name
        $workspaceId = $workspace.id

        # Use workspace ID in the XMLA endpoint to avoid issues with duplicate or special-character workspace names.
        # Safer URI format: powerbi://api.powerbi.com/v1.0/myorg/<workspaceId>
        $workspaceConnection = "powerbi://api.powerbi.com/v1.0/myorg/$workspaceId"

        Write-Host "`nProcessing workspace: $workspaceName ($workspaceId)" -ForegroundColor Green

        try {
            # Use workspace name for XMLA connection (this was working before).
            $workspaceConnection = "powerbi://api.powerbi.com/v1.0/myorg/$workspaceName"
            $connectionString = "DataSource=$workspaceConnection;User ID=app:$clientId@$tenantId;Password=$clientSecret;"

            $server = New-Object Microsoft.AnalysisServices.Tabular.Server
            Write-Host "Attempting XMLA connect using workspace Name URI: $workspaceConnection" -ForegroundColor Yellow
            $server.Connect($connectionString)
            Write-Host "Connected to workspace. Processing datasets..." -ForegroundColor Green

            foreach ($db in $server.Databases) {
                Write-Host "Processing dataset: $($db.Name)" -ForegroundColor Cyan
                $model = $db.Model
                $dataSources = @{}
                foreach ($ds in $model.DataSources) {
                    $dataSources[$ds.Name] = @{
                        Database         = $ds.Database
                        Schema           = $ds.Schema
                        ConnectionString = $ds.ConnectionString
                        Type             = $ds.Type.ToString()
                    }
                }

                foreach ($table in $model.Tables) {
                    foreach ($column in $table.Columns) {
                        $metadata.Add([PSCustomObject]@{
                            WorkspaceName = $workspaceName
                            WorkspaceId   = $workspaceId
                            DatasetName   = $db.Name
                            TableName     = $table.Name
                            ObjectType    = "Column"
                            ObjectName    = $column.Name
                            DataType      = $column.DataType.ToString()
                            IsHidden      = $column.IsHidden
                            Description   = $column.Description
                            DataSource    = $null
                            Database      = $null
                            Schema        = $null
                            MCode         = $null
                        })
                    }

                    foreach ($measure in $table.Measures) {
                        $metadata.Add([PSCustomObject]@{
                            WorkspaceName = $workspaceName
                            WorkspaceId   = $workspaceId
                            DatasetName   = $db.Name
                            TableName     = $table.Name
                            ObjectType    = "Measure"
                            ObjectName    = $measure.Name
                            DataType      = $measure.DataType.ToString()
                            IsHidden      = $measure.IsHidden
                            Description   = $measure.Description
                            DataSource    = $null
                            Database      = $null
                            Schema        = $null
                            MCode         = $measure.Expression
                        })
                    }

                    foreach ($partition in $table.Partitions) {
                        if ($partition.SourceType -eq "M") {
                            $mExpr = $partition.Source.Expression
                            $srcName = "Power Query"
                            $dbName = $null

                            if ($mExpr -match 'Sql\.Database\("([^"]+)"\s*,\s*"([^"]+)"') {
                                $srcName = $matches[1]
                                $dbName  = $matches[2]
                            } elseif ($mExpr -match 'Odbc\.Datasource\("([^"]+)"[^;]*;Database=([^;"]+)') {
                                $srcName = $matches[1]
                                $dbName  = $matches[2]
                            }

                            $sourceInfo = @{
                                MCode      = $mExpr
                                DataSource = $srcName
                                Database   = $dbName
                                Schema     = $null
                            }
                        } else {
                            $sourceInfo = @{
                                MCode      = $null
                                DataSource = $partition.Source.DataSource
                                Database   = $partition.Source.Database
                                Schema     = $partition.Source.Schema
                            }
                        }

                        $metadata.Add([PSCustomObject]@{
                            WorkspaceName = $workspaceName
                            WorkspaceId   = $workspaceId
                            DatasetName   = $db.Name
                            TableName     = $table.Name
                            ObjectType    = "Partition"
                            ObjectName    = $partition.Name
                            DataType      = $null
                            IsHidden      = $null
                            Description   = $partition.Description
                            DataSource    = $sourceInfo.DataSource
                            Database      = $sourceInfo.Database
                            Schema        = $sourceInfo.Schema
                            MCode         = $sourceInfo.MCode
                        })
                    }
                }

                foreach ($dsName in $dataSources.Keys) {
                    $ds = $dataSources[$dsName]
                    $metadata.Add([PSCustomObject]@{
                        WorkspaceName = $workspaceName
                        WorkspaceId   = $workspaceId
                        DatasetName   = $db.Name
                        TableName     = $null
                        ObjectType    = "DataSource"
                        ObjectName    = $dsName
                        DataType      = $ds.Type
                        IsHidden      = $null
                        Description   = $null
                        DataSource    = $dsName
                        Database      = $ds.Database
                        Schema        = $ds.Schema
                        MCode         = $ds.ConnectionString
                    })
                }
            }
        } catch {
            Write-Host "Error processing workspace $workspaceName ($workspaceId): $_" -ForegroundColor Red
            continue
        } finally {
            if ($server -and $server.Connected) {
                $server.Disconnect()
                Write-Host "Disconnected from workspace $workspaceName" -ForegroundColor Green
            }
        }
    }

    $metadata | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

    Write-Host "`n✅ Metadata extraction completed!" -ForegroundColor Green
    Write-Host "Results saved to: $outputPath" -ForegroundColor Yellow
    Write-Host "Total objects extracted: $($metadata.Count)" -ForegroundColor Cyan
    Write-Host "Total workspaces processed: $($workspaces.Count)" -ForegroundColor Cyan
}
catch {
    Write-Error "Script failed: $_"
}
finally {
    Write-Host "`n⏳ Waiting 30 seconds before closing..."
    Start-Sleep -Seconds 30
}
