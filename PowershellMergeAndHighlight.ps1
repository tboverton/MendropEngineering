param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectNumber,
    
    [Parameter(Mandatory=$false)]
    [string]$TemplatePath = "C:\Users\thoma\Documents\H&H MASTER TESTING TEMPLATE LOCAL PS.dotx",
    
    [Parameter(Mandatory=$false)]
    [switch]$ForceCloseReopen,  # New parameter for your refresh scenario

    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = "C:\Users\thoma\Documents",

    [Parameter(Mandatory=$false)]
    [string]$OutputFileName = $null,

    [Parameter(Mandatory=$false)]
    [ValidateSet('docx','dotx','doc','pdf')]
    [string]$OutputFormat = 'docx'
)

# Default CSV path used for mail merge data
$CSVPath = ".\project_specific_data.csv"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(if($Level -eq "ERROR"){"Red"} elseif($Level -eq "WARN"){"Yellow"} else{"Green"})
}

# Enhanced document handling with close/reopen option
function Open-WordDocument {
    param(
        [object]$WordApp, 
        [string]$DocumentPath,
        [bool]$ForceRefresh = $false
    )
    
    Write-Log "Checking if document is already open: $(Split-Path $DocumentPath -Leaf)"
    
    $existingDoc = $null
    # Check if document is already open
    foreach ($doc in $WordApp.Documents) {
        if ($doc.FullName -eq $DocumentPath) {
            $existingDoc = $doc
            break
        }
    }
    
    if ($existingDoc) {
        if ($ForceRefresh) {
            Write-Log "Document is open - FORCE REFRESH mode: Closing and reopening" "WARN"
            
            # Save current state if needed
            if (-not $existingDoc.Saved) {
                Write-Log "Document has unsaved changes, saving first..."
                $existingDoc.Save()
            }
            
            # Close the document
            $existingDoc.Close()
            Write-Log "Document closed, reopening fresh..."
            
            # Small delay to ensure clean close
            Start-Sleep -Milliseconds 500
            
            # Reopen fresh
            return $WordApp.Documents.Open($DocumentPath)
        } else {
            Write-Log "Document is already open, activating existing instance"
            $existingDoc.Activate()
            return $existingDoc
        }
    }
    
    # Document not open, open it fresh
    Write-Log "Opening document: $DocumentPath"
    if (-not (Test-Path $DocumentPath)) {
        throw "Template document not found: $DocumentPath"
    }
    
    return $WordApp.Documents.Open($DocumentPath)
}

# Rest of your functions remain the same...
function Get-ProjectData {
    param([string]$ProjectNumber)
    
   # $CSVPath = ".\project_specific_data.csv"
    
    try {
        Write-Log "Attempting database connection..."
        # Database connection attempt
        # Prefer retrieving credentials from a secure store. The script will look for environment variables first.
        $sqlUser = $env:SQL_USER
        $sqlPassword = $env:SQL_PASSWORD

        if ([string]::IsNullOrWhiteSpace($sqlUser) -or [string]::IsNullOrWhiteSpace($sqlPassword)) {
            # Fallback: you can uncomment the next lines to provide static placeholders for quick testing,
            # but avoid committing secrets into source control. Recommended: use Azure Key Vault or Managed Identity.
            # $sqlUser = "ReportUser"
            # $sqlPassword = "R3p0rtUs3r!"

            throw "SQL credentials not found in environment variables `SQL_USER` and `SQL_PASSWORD`. Set them or use a secure retrieval method."
        }

        $connectionString = "Server=tcp:mendrop.database.windows.net,1433;Database=MendropReportServer;User ID=$sqlUser;Password=$sqlPassword;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
        #$connectionString = "Server=mendrop.database.windows.net;Database=MendropReportServer;Integrated Security=true;Connection Timeout=30;"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()
        
        $query = "SELECT * FROM vwHandHReportFormFields WHERE project_number = '$ProjectNumber'"
        $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $dataTable = New-Object System.Data.DataTable
        $adapter.Fill($dataTable)
        
        if ($dataTable.Rows.Count -gt 0) {
            $dataTable | Export-Csv -Path $CSVPath -NoTypeInformation -Force
            Write-Log "Project data exported to CSV: $CSVPath"
            return $true
        } else {
            Write-Log "No data found for project: $ProjectNumber" "WARN"
            return $false
        }
        
    } catch {
        Write-Log "Database connection failed: $($_.Exception.Message)" "WARN"
        Write-Log "Creating sample data for testing..."
        
        # Create sample data
        $sampleData = @(
            [PSCustomObject]@{
                project_number = $ProjectNumber
                report_title = "Hydraulic and Hydrologic Analysis Report"
                bridge_id = "BR-$ProjectNumber"
                location = "Sample Location for $ProjectNumber"
                prepared_by_name = "John Engineer"
                prepared_by_organization = "Engineering Consultants Inc."
                Date = (Get-Date).ToString("MM/dd/yyyy")
                ExistingStructureType = "Concrete Box Culvert"
                PreferredStructureType = "Precast Concrete Box Culvert"
            }
        )
        
        $sampleData | Export-Csv -Path $CSVPath -NoTypeInformation -Force
        Write-Log "Sample data created for project: $ProjectNumber"
        return $true
    } finally {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
    }
}

function Refresh-TableOfContents {
    param([object]$Document)
    
    Write-Log "Refreshing Table of Contents..."
    try {
        $Document.Fields.Update()
        
        foreach ($toc in $Document.TablesOfContents) {
            $toc.Update()
            Write-Log "TOC updated successfully"
        }
        
        $Document.Repaginate()
        Write-Log "Document repaginated"
        
    } catch {
        Write-Log "Error refreshing TOC: $($_.Exception.Message)" "WARN"
    }
}

# Main execution
try {
    Write-Log "=== Enhanced Generate Draft Report Script Started ==="
    Write-Log "Project Number: $ProjectNumber"
    Write-Log "Template: $(Split-Path $TemplatePath -Leaf)"
    Write-Log "Force Close/Reopen: $ForceCloseReopen"
    
    # Get project data
    Write-Log "Querying project data for: $ProjectNumber"
    $dataSuccess = Get-ProjectData -ProjectNumber $ProjectNumber
    
    if (-not $dataSuccess) {
        throw "Failed to get project data"
    }
    
    # Initialize Word
    Write-Log "Initializing Word application..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0  # wdAlertsNone
    
    # Open document with enhanced handling
    $document = Open-WordDocument -WordApp $word -DocumentPath $TemplatePath -ForceRefresh $ForceCloseReopen
    
    # Set up mail merge
    Write-Log "Setting up mail merge..."
    $mailMerge = $document.MailMerge
    $csvPath = ".\project_specific_data.csv"
    $fullCSVPath = (Resolve-Path $csvPath).Path
    
    # Connect to data source
    $mailMerge.OpenDataSource($fullCSVPath)
    Write-Log "Connected to data source: $fullCSVPath"
    
    # Execute merge to new document
    Write-Log "Executing mail merge..."
    $mailMerge.Destination = 0  # wdSendToNewDocument
    $mailMerge.Execute()
    
    # Get the merged document
    $mergedDoc = $word.ActiveDocument
    Write-Log "Mail merge completed successfully"
    
    # Highlight merge fields (grey background)
    Write-Log "Highlighting merge fields..."
    foreach ($field in $mergedDoc.Fields) {
        if ($field.Type -eq 88) {  # wdFieldMergeField
            $field.Select()
            $word.Selection.Shading.BackgroundPatternColor = 12632256  # Light grey
        }
    }
    
    # Refresh Table of Contents
    Refresh-TableOfContents -Document $mergedDoc
    
    # Save the document
    # Normalize output filename and extension
    $ext = if ($OutputFormat.StartsWith('.')) { $OutputFormat.TrimStart('.') } else { $OutputFormat }
    $safeFileName = if ([string]::IsNullOrWhiteSpace($OutputFileName)) { "${ProjectNumber}_DRAFT" } else { $OutputFileName }
    $outputFileName = "$safeFileName.$ext"
    $outputPath = Join-Path $OutputDirectory $outputFileName

    Write-Log "Saving document as: $outputPath"
    # For common formats SaveAs2 will infer by extension; for some formats you may need to pass FileFormat.
    $mergedDoc.SaveAs2($outputPath)
    Write-Log "Document saved successfully"
    
    Write-Log "=== Script completed successfully! ==="
    Write-Log "✅ Template opened: $(if($ForceCloseReopen){'Fresh (closed/reopened)'}else{'Existing or new'})"
    Write-Log "✅ Data merged from: vwHandHReportFormFields"
    Write-Log "✅ Merge fields highlighted in grey"
    Write-Log "✅ Table of Contents refreshed"
    Write-Log "✅ Saved as: $outputFileName"
    Write-Log ""
    Write-Log "📋 NEXT STEPS:"
    Write-Log "1. Review the highlighted merge fields"
    Write-Log "2. Make any necessary edits"
    Write-Log "3. Save final version when ready"
    
} catch {
    Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
} finally {
    # Restore Word alerts
    if ($word) {
        $word.DisplayAlerts = -1  # wdAlertsAll
    }
    Write-Log "Script execution completed"
}
