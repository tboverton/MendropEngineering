# Word Document Highlight Processor with Database Integration
# This script opens a Word template, processes highlights, integrates with database for project selection,
# and handles mail merge functionality with project and bridge ID selection

param(
    [Parameter(Mandatory=$false)]
    [string]$InputPath = "C:\Users\user\OneDrive\Documents\Work_Files\Word Doc to database\HandHPS.dotm",
   
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "",
   
    [Parameter(Mandatory=$false)]
    [string]$TemplatePath = "C:\Users\user\OneDrive\Documents\Work_Files\Word Doc to database\HandHPS.dotm",
   
    [Parameter(Mandatory=$false)]
    [string]$ProjectNumber = "H-025-999-25",
   
    [Parameter(Mandatory=$false)]
    [string]$BridgeId = "K160.0",
   
    [Parameter(Mandatory=$false)]
    [switch]$ShowProjectPicker = $false
)

# Set default paths based on current script location
$script:ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DefaultTemplatePath = Join-Path $script:ScriptDirectory "HandHPS.dotm"
$script:DefaultOutputDirectory = $script:ScriptDirectory

# Use current directory template if not specified
if ($TemplatePath -eq "") {
    $TemplatePath = $script:DefaultTemplatePath
}

# Database connection string
$script:ConnectionString = "Driver={ODBC Driver 17 for SQL Server};" +
                          "Server=tcp:mendrop.database.windows.net;" +
                          "Database=MendropReportServer;" +
                          "UID=ReportUser;" +
                          "PWD=R3p0rtUs3r!;"

# Logging configuration
$script:LogFilePath = Join-Path $script:ScriptDirectory "word_processor_log.txt"
$script:LoggingEnabled = $true

# Function to write to log file
function Write-LogFile {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    if ($script:LoggingEnabled) {
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Add-Content -Path $script:LogFilePath -Value $logEntry -Encoding UTF8
        } catch {
            # Silently continue if logging fails to avoid breaking the main script
        }
    }
}

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    
    # Write to console with color
    Write-Host $Message -ForegroundColor $Color
    
    # Also write to log file
    $logLevel = switch ($Color) {
        "Red" { "ERROR" }
        "Yellow" { "WARN" }
        "Green" { "SUCCESS" }
        "Cyan" { "INFO" }
        default { "INFO" }
    }
    Write-LogFile -Message $Message -Level $logLevel
}

# Function to get database connection
function Get-DatabaseConnection {
    try {
        $connection = New-Object -ComObject ADODB.Connection
        $connection.Open($script:ConnectionString)
        Write-ColorOutput "Database connection established successfully" "Green"
        return $connection
    } catch {
        Write-ColorOutput "Failed to connect to database: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to get distinct project numbers from database
function Get-ProjectNumbers {
    try {
        Write-ColorOutput "Retrieving project numbers from database..." "Yellow"
       
        $connection = Get-DatabaseConnection
        $recordset = $connection.Execute("SELECT * FROM [dbo].[vwHandHReportFormFields] WHERE project_number = '$ProjectNumber'AND bridge_id = '$BridgeId' ")
       
        $projects = @()
        while (-not $recordset.EOF) {
            $projects += $recordset.Fields("project_number").Value
            $recordset.MoveNext()
        }
       
        $recordset.Close()
        $connection.Close()
       
        Write-ColorOutput "Retrieved $($projects.Count) project numbers" "Green"
        return $projects
       
    } catch {
        Write-ColorOutput "Error retrieving project numbers: $($_.Exception.Message)" "Red"
        throw
    }
}

# # Function to get bridge IDs for a specific project
# function Get-BridgeIds {
#     param([string]$ProjectNumber)
   
#     try {
#         Write-ColorOutput "Retrieving bridge IDs for project: $ProjectNumber" "Yellow"
       
#         $connection = Get-DatabaseConnection
#         $query = "SELECT bridge_id FROM [dbo].[vwHandHReportFormFields] WHERE project_number = '$ProjectNumber' AND bridge_id = '$BridgeId'"
#         $recordset = $connection.Execute($query)
       
#         $bridgeIds = @()
#         while (-not $recordset.EOF) {
#             $bridgeIds += $recordset.Fields("bridge_id").Value
#             $recordset.MoveNext()
#         }
       
#         $recordset.Close()
#         $connection.Close()
       
#         Write-ColorOutput "Retrieved $($bridgeIds.Count) bridge IDs" "Green"
#         return $bridgeIds
       
#     } catch {
#         Write-ColorOutput "Error retrieving bridge IDs: $($_.Exception.Message)" "Red"
#         throw
#     }
# }

# Function to show project picker dialog
function Show-ProjectPicker {
    try {
        Write-ColorOutput "Loading project picker..." "Yellow"
       
        # Get project numbers
        $projects = Get-ProjectNumbers
       
        if ($projects.Count -eq 0) {
            Write-ColorOutput "No projects found in database" "Red"
            return $null
        }
       
        # Create selection dialog
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
       
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select Project and Bridge ID"
        $form.Size = New-Object System.Drawing.Size(400, 300)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
       
        # Project label and combobox
        $projectLabel = New-Object System.Windows.Forms.Label
        $projectLabel.Location = New-Object System.Drawing.Point(10, 20)
        $projectLabel.Size = New-Object System.Drawing.Size(100, 20)
        $projectLabel.Text = "Project Number:"
        $form.Controls.Add($projectLabel)
       
        $projectCombo = New-Object System.Windows.Forms.ComboBox
        $projectCombo.Location = New-Object System.Drawing.Point(120, 18)
        $projectCombo.Size = New-Object System.Drawing.Size(250, 20)
        $projectCombo.DropDownStyle = "DropDownList"
        $projects | ForEach-Object { $projectCombo.Items.Add($_) }
        $form.Controls.Add($projectCombo)
       
        # Bridge ID label and combobox
        $bridgeLabel = New-Object System.Windows.Forms.Label
        $bridgeLabel.Location = New-Object System.Drawing.Point(10, 60)
        $bridgeLabel.Size = New-Object System.Drawing.Size(100, 20)
        $bridgeLabel.Text = "Bridge ID:"
        $form.Controls.Add($bridgeLabel)
       
        $bridgeCombo = New-Object System.Windows.Forms.ComboBox
        $bridgeCombo.Location = New-Object System.Drawing.Point(120, 58)
        $bridgeCombo.Size = New-Object System.Drawing.Size(250, 20)
        $bridgeCombo.DropDownStyle = "DropDownList"
        $bridgeCombo.Enabled = $false
        $form.Controls.Add($bridgeCombo)
       
        # Event handler for project selection
        $projectCombo.Add_SelectedIndexChanged({
            $bridgeCombo.Items.Clear()
            $bridgeCombo.Enabled = $false
           
            if ($projectCombo.SelectedItem) {
                $bridgeIds = Get-BridgeIds -ProjectNumber $projectCombo.SelectedItem
                $bridgeIds | ForEach-Object { $bridgeCombo.Items.Add($_) }
                $bridgeCombo.Enabled = $true
            }
        })
       
        # OK button
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(200, 120)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Text = "OK"
        $okButton.DialogResult = "OK"
        $okButton.Enabled = $false
        $form.Controls.Add($okButton)
       
        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(295, 120)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = "Cancel"
        $form.Controls.Add($cancelButton)
       
        # Enable OK button when both selections are made
        $bridgeCombo.Add_SelectedIndexChanged({
            $okButton.Enabled = ($projectCombo.SelectedItem -and $bridgeCombo.SelectedItem)
        })
       
        $form.AcceptButton = $okButton
        $form.CancelButton = $cancelButton
       
        $result = $form.ShowDialog()
       
        if ($result -eq "OK") {
            return @{
                ProjectNumber = $projectCombo.SelectedItem
                BridgeId = $bridgeCombo.SelectedItem
            }
        }
       
        return $null
       
    } catch {
        Write-ColorOutput "Error showing project picker: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to perform mail merge with database data
function Invoke-MailMerge {
    param(
        [object]$Document,
        [string]$ProjectNumber,
        [string]$BridgeId
    )
   
    try {
        Write-ColorOutput "Performing mail merge for Project: $ProjectNumber, Bridge: $BridgeId" "Yellow"
       
        # Get data from database
        $connection = Get-DatabaseConnection
        $query = "SELECT * FROM [dbo].[vwHandHReportFormFields] WHERE project_number = '$ProjectNumber' AND bridge_id = '$BridgeId'"
        $recordset = $connection.Execute($query)
       
        if ($recordset.EOF) {
            Write-ColorOutput "No data found for the specified project and bridge ID" "Red"
            $recordset.Close()
            $connection.Close()
            return $false
        }
       
        # Get field names and values
        $fieldData = @{}
        for ($i = 0; $i -lt $recordset.Fields.Count; $i++) {
            $fieldName = $recordset.Fields($i).Name
            
            # Handle NULL/DBNull values properly
            try {
                $fieldValue = $recordset.Fields($i).Value
                if ($fieldValue -eq $null -or [System.DBNull]::Value.Equals($fieldValue)) {
                    $fieldValue = ""
                } else {
                    # Convert to string to ensure compatibility
                    $fieldValue = $fieldValue.ToString()
                }
            } catch {
                # If there's any error accessing the value, use empty string
                $fieldValue = ""
            }
            
            $fieldData[$fieldName] = $fieldValue
        }
       
        $recordset.Close()
        $connection.Close()
       
        Write-ColorOutput "Retrieved $($fieldData.Count) fields from database" "Green"
        
        # Show first few field names for debugging
        $sampleFields = ($fieldData.Keys | Select-Object -First 10) -join ", "
        Write-ColorOutput "Sample field names: $sampleFields" "Cyan"
        
        # Show template fields we're looking for (from the document you showed)
        $templateFields = @("hydraulic_modeling_tool", "tool_version", "tool_source", "model_range_upstream", "upstream_bridge_id", "model_range_downstream", "geospatial_tool", "geospatial_software", "elevation_model_detail", "elevation_model_source", "elevation_model_website", "data_source_type", "data_source_location", "data_collection_year", "manning_n_min", "manning_n_max", "flood_event_1", "flood_event_2", "flood_event_3", "boundary_condition_type", "slope_value", "number_of_alternatives", "alt1_structure_type", "alt1_span_length", "alt1_total_length", "alt1_girder_depth", "alt1_low_chord_elevation", "alt1_length_comparison", "alt2_structure_type", "alt2_span_length", "alt2_total_length", "alt2_girder_depth", "alt2_low_chord_elevation", "alt2_length_comparison")
        
        $matchingFields = @()
        foreach ($templateField in $templateFields) {
            if ($fieldData.ContainsKey($templateField)) {
                $matchingFields += $templateField
            }
        }
        
        if ($matchingFields.Count -gt 0) {
            Write-ColorOutput "Found matching template fields: $($matchingFields -join ', ')" "Green"
        } else {
            Write-ColorOutput "No matching template fields found in database" "Yellow"
            Write-ColorOutput "Template expects fields like: $($templateFields[0..4] -join ', ')..." "Yellow"
        }
       
        # Replace merge fields in document
        $Range = $Document.Range()
        $replacedFields = 0
        foreach ($fieldName in $fieldData.Keys) {
            try {
                $fieldValue = $fieldData[$fieldName]
                $fieldReplaced = $false
               
                # Try both merge field formats: «field» and <<field>>
                $mergeFormats = @("«$fieldName»", "<<$fieldName>>")
                
                foreach ($mergeField in $mergeFormats) {
                    # Reset range for each search
                    $Range = $Document.Range()
                    
                    # Find and replace merge fields
                    $Range.Find.ClearFormatting()
                    $Range.Find.Replacement.ClearFormatting()
                    $Range.Find.Text = $mergeField
                    $Range.Find.Replacement.Text = $fieldValue
                    
                    # Set orange highlighting for replaced text
                    $Range.Find.Replacement.Highlight = $true
                    $Range.Find.Replacement.Font.Color = 255  # Orange color (RGB: 255, 165, 0)
                    
                    # Execute the find and replace with ReplaceAll (2)
                    $replaceCount = $Range.Find.Execute("", $false, $false, $false, $false, $false, $true, 1, $false, "", 2)
                    
                    if ($replaceCount -gt 0) {
                        $fieldReplaced = $true
                        Write-ColorOutput "Replaced field: $mergeField -> $fieldValue (Count: $replaceCount)" "Green"
                    } else {
                        Write-ColorOutput "Field not found in document: $mergeField" "Yellow"
                    }
                }
                
                if ($fieldReplaced) {
                    $replacedFields++
                }
               
            } catch {
                Write-ColorOutput "Warning: Could not replace field '$fieldName': $($_.Exception.Message)" "Yellow"
                # Continue with next field
            }
        }
        
        Write-ColorOutput "Successfully replaced $replacedFields merge fields" "Green"
       
        Write-ColorOutput "Mail merge completed successfully" "Green"
        return $true
       
    } catch {
        Write-ColorOutput "Error during mail merge: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to process Word document with project integration
function Process-WordDocument {
    param(
        [string]$DocumentPath,
        [string]$SavePath = "",
        [string]$ProjectNumber = "",
        [string]$BridgeId = ""
    )
   
    try {
        Write-ColorOutput "Starting Word document processing..." "Green"
       
        # Create Word Application COM object
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Word.DisplayAlerts = 0  # Disable alerts
       
        Write-ColorOutput "Opening document: $DocumentPath" "Yellow"
       
        # Check if file exists
        if (-not (Test-Path $DocumentPath)) {
            throw "File not found: $DocumentPath"
        }
       
        # Open the document
        $Document = $Word.Documents.Open($DocumentPath)
       
        Write-ColorOutput "Document opened successfully" "Green"
       
        # Step 1: Remove all yellow highlights
        Write-ColorOutput "Removing yellow highlights..." "Yellow"
        $Range = $Document.Range()
       
        # Find and remove yellow highlights (wdYellow = 7)
        $Range.Find.ClearFormatting()
        $Range.Find.Replacement.ClearFormatting()
        $Range.Find.Highlight = $true
        $Range.Find.Replacement.Highlight = $false
        $Range.Find.Execute("", $false, $false, $false, $false, $false, $true, 1, $false, "", 2)
       
        Write-ColorOutput "Yellow highlights removed" "Green"
       
        # Step 2: Perform mail merge if project data is provided
        if ($ProjectNumber -and $BridgeId) {
            try {
                Write-ColorOutput "Attempting database connection for mail merge..." "Yellow"
                $mergeSuccess = Invoke-MailMerge -Document $Document -ProjectNumber $ProjectNumber -BridgeId $BridgeId
                if (-not $mergeSuccess) {
                    Write-ColorOutput "Mail merge failed - no data found for project" "Yellow"
                    # Insert project info manually when no data is found
                    $Document.Range().InsertBefore("Project Number: $ProjectNumber`r`nBridge ID: $BridgeId`r`nGenerated: $(Get-Date -Format 'MM/dd/yyyy hh:mm tt')`r`n`r`n")
                }
            } catch {
                Write-ColorOutput "Database connection failed - preserving template content" "Yellow"
                Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
                Write-ColorOutput "All merge fields from template will be preserved and highlighted" "Green"
                # Insert project info manually since database is unavailable
                $Document.Range().InsertBefore("Project Number: $ProjectNumber`r`nBridge ID: $BridgeId`r`nGenerated: $(Get-Date -Format 'MM/dd/yyyy hh:mm tt')`r`n`r`n")
                Write-ColorOutput "Template content preserved - merge fields will be highlighted for manual completion" "Cyan"
            }
        } else {
            Write-ColorOutput "No project data provided - preserving all template content" "Yellow"
        }
       
        # Step 3: Find and highlight remaining merge field patterns
        Write-ColorOutput "Adding highlights for remaining merge field patterns..." "Yellow"
       
        $totalHighlighted = 0
        
        # Highlight remaining «*» patterns (guillemets)
        $Range = $Document.Range()
        $Range.Find.ClearFormatting()
        $Range.Find.Replacement.ClearFormatting()
        $Range.Find.Text = "«*»"
        $Range.Find.MatchWildcards = $true
        $Range.Find.Replacement.Highlight = $true
        $Range.Find.Replacement.Text = "^&"  # Replace with same text but highlighted
        $guillemetsCount = $Range.Find.Execute("", $false, $false, $false, $false, $false, $true, 1, $false, "", 2)
        if ($guillemetsCount) { $totalHighlighted++ }
        
        # Highlight remaining <<*>> patterns (angle brackets)
        $Range = $Document.Range()
        $Range.Find.ClearFormatting()
        $Range.Find.Replacement.ClearFormatting()
        $Range.Find.Text = "<<*>>"
        $Range.Find.MatchWildcards = $true
        $Range.Find.Replacement.Highlight = $true
        $Range.Find.Replacement.Text = "^&"  # Replace with same text but highlighted
        $angleBracketsCount = $Range.Find.Execute("", $false, $false, $false, $false, $false, $true, 1, $false, "", 2)
        if ($angleBracketsCount) { $totalHighlighted++ }
       
        Write-ColorOutput "Merge field patterns highlighted (both «field» and <<field>> formats)" "Green"
       
        # Step 4: Generate save path with project naming convention
        if ($SavePath -eq "") {
            $fileInfo = Get-Item $DocumentPath
            if ($ProjectNumber -and $BridgeId) {
                # Use project_number_bridge_id_DRAFT.docx naming convention
                $fileName = "${ProjectNumber}_${BridgeId}_DRAFT.docx"
                $SavePath = Join-Path $fileInfo.DirectoryName $fileName
            } else {
                # Default naming - preserve original extension or use .docx
                $extension = if ($fileInfo.Extension -eq ".dotm" -or $fileInfo.Extension -eq ".dotx") { ".docx" } else { $fileInfo.Extension }
                $SavePath = Join-Path $fileInfo.DirectoryName ($fileInfo.BaseName + "_processed" + $extension)
            }
        }
       
        Write-ColorOutput "Saving document to: $SavePath" "Yellow"
       
        # Ensure we save as Word document format (not template)
        # Use wdFormatDocumentDefault (16) for .docx format
        $Document.SaveAs2($SavePath, 16)
       
        Write-ColorOutput "Document saved successfully" "Green"
       
        # Close document and quit Word
        $Document.Close()
        $Word.Quit()
       
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
       
        Write-ColorOutput "Processing completed successfully!" "Green"
        Write-ColorOutput "Output file: $SavePath" "Cyan"
       
        return $SavePath
       
    } catch {
        Write-ColorOutput "Error occurred: $($_.Exception.Message)" "Red"
       
        # Cleanup in case of error
        if ($Document) {
            try { $Document.Close($false) } catch { }
        }
        if ($Word) {
            try { $Word.Quit() } catch { }
        }
       
        throw
    }
}

# Function to create Word template with VBA integration
function New-WordTemplateWithVBA {
    param([string]$TemplatePath = "handhps.dotm")
   
    try {
        Write-ColorOutput "Creating Word template with VBA integration..." "Yellow"
       
        # This function would create the .dotm template with embedded VBA
        # For now, we'll provide the VBA code that needs to be manually added
       
        $vbaCode = @"
' VBA Code for handhps.dotm template
' This code should be added to the template manually

' UserForm: frmProjectPicker
' Controls: ComboBox1 (project numbers), ComboBox2 (bridge IDs), btnSelect (command button)

Private Sub btnSelect_Click()
    Dim selectedProject As String
    Dim selectedBridge As String
   
    selectedProject = ComboBox1.Value
    selectedBridge = ComboBox2.Value
   
    If selectedProject = "" Or selectedBridge = "" Then
        MsgBox "Please select both project number and bridge ID.", vbExclamation
        Exit Sub
    End If
   
    ' Call PowerShell script with selected values
    Dim shell As Object
    Dim psCommand As String
   
    Set shell = CreateObject("WScript.Shell")
    psCommand = "powershell.exe -ExecutionPolicy Bypass -File ""$($PSScriptRoot)\word_highlight_processor.ps1"" " & _
                "-TemplatePath ""$($Document.FullName)"" " & _
                "-ProjectNumber """ & selectedProject & """ " & _
                "-BridgeId """ & selectedBridge & """"
   
    shell.Run psCommand, 0, True
   
    Me.Hide
End Sub

' Main macro in ThisDocument
Sub PickProjectReport()
    Dim conn As Object
    Dim rs As Object
    Dim connectionString As String
   
    ' Define connection string
    connectionString = "Driver={ODBC Driver 17 for SQL Server};" & _
                      "Server=tcp:mendrop.database.windows.net;" & _
                      "Database=MendropReportServer;" & _
                      "UID=ReportUser;" & _
                      "PWD=R3p0rtUs3r!;"
   
    ' Create and open connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connectionString
   
    ' Execute query for projects
    Set rs = conn.Execute("SELECT DISTINCT project_number FROM [dbo].[vwHandHReportFormFields] ORDER BY project_number")
   
    ' Populate ComboBox1
    frmProjectPicker.ComboBox1.Clear
    Do While Not rs.EOF
        frmProjectPicker.ComboBox1.AddItem rs.Fields("project_number").Value
        rs.MoveNext
    Loop
   
    ' Show the form
    frmProjectPicker.Show
   
    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

' Auto-launch on document open
Private Sub Document_Open()
    PickProjectReport
End Sub
"@
       
        Write-ColorOutput "VBA Code generated. This needs to be manually added to the Word template." "Green"
        Write-ColorOutput "Save the following VBA code to a file for reference:" "Yellow"
       
        $vbaCodePath = Join-Path (Split-Path $PSScriptRoot) "VBA_Code_for_Template.txt"
        $vbaCode | Out-File -FilePath $vbaCodePath -Encoding UTF8
       
        Write-ColorOutput "VBA code saved to: $vbaCodePath" "Cyan"
       
        return $vbaCodePath
       
    } catch {
        Write-ColorOutput "Error creating VBA template: $($_.Exception.Message)" "Red"
        throw
    }
}

# Main execution
try {
    # Initialize logging
    Write-LogFile -Message "=== NEW SESSION STARTED ===" -Level "INFO"
    Write-LogFile -Message "Script: word_marco_dotm.ps1" -Level "INFO"
    Write-LogFile -Message "User: $env:USERNAME" -Level "INFO"
    Write-LogFile -Message "Computer: $env:COMPUTERNAME" -Level "INFO"
    Write-LogFile -Message "Working Directory: $(Get-Location)" -Level "INFO"
    Write-LogFile -Message "Log File: $script:LogFilePath" -Level "INFO"
    
    Write-ColorOutput "=== Word Document Processor with Database Integration ===" "Cyan"
   
    # Interactive input for missing parameters
    if ($TemplatePath -eq "" -or $ProjectNumber -eq "" -or $BridgeId -eq "") {
        Write-ColorOutput "Interactive mode: Please provide the following information..." "Yellow"
       
        # Prompt for template path if not provided
        if ($TemplatePath -eq "") {
            do {
                $TemplatePath = Read-Host "Enter the path to the .dotm template file (or press Enter for default HandHPS.dotm)"
                if ($TemplatePath -eq "") {
                    $TemplatePath = $script:DefaultTemplatePath
                }
                if (-not (Test-Path $TemplatePath)) {
                    Write-ColorOutput "File not found: $TemplatePath" "Red"
                    $TemplatePath = ""
                }
            } while ($TemplatePath -eq "")
        }
       
        # Prompt for project number if not provided
        if ($ProjectNumber -eq "") {
            do {
                $ProjectNumber = Read-Host "Enter the Project Number (e.g., H-025-999-25)"
                if ($ProjectNumber -eq "") {
                    Write-ColorOutput "Project Number is required!" "Red"
                }
            } while ($ProjectNumber -eq "")
        }
       
        # Prompt for bridge ID if not provided
        if ($BridgeId -eq "") {
            do {
                $BridgeId = Read-Host "Enter the Bridge ID (e.g., K160.0)"
                if ($BridgeId -eq "") {
                    Write-ColorOutput "Bridge ID is required!" "Red"
                }
            } while ($BridgeId -eq "")
        }
       
        Write-ColorOutput "Using Template: $TemplatePath" "Green"
        Write-ColorOutput "Project Number: $ProjectNumber" "Green"
        Write-ColorOutput "Bridge ID: $BridgeId" "Green"
    }
   
    # Handle different execution modes
    if ($ShowProjectPicker) {
        # Show project picker dialog
        Write-ColorOutput "Launching project picker..." "Yellow"
        $selection = Show-ProjectPicker
       
        if ($selection) {
            $ProjectNumber = $selection.ProjectNumber
            $BridgeId = $selection.BridgeId
            Write-ColorOutput "Selected Project: $ProjectNumber, Bridge: $BridgeId" "Green"
        } else {
            Write-ColorOutput "No selection made. Exiting." "Yellow"
            exit 0
        }
    }
   
    # Determine input path
    if ($InputPath -eq "") {
        if (Test-Path $TemplatePath) {
            $InputPath = $TemplatePath
            Write-ColorOutput "Using template: $TemplatePath" "Green"
        } else {
            Write-ColorOutput "Template not found: $TemplatePath" "Red"
            Write-ColorOutput "Please ensure the Word file (.docx, .dotx, or .dotm) exists in the specified path." "Yellow"
            Write-ColorOutput "Usage: .\word_marco_dotm.ps1 -ProjectNumber 'H-025-999-25' -BridgeId 'K160.0'" "White"
            exit 1
        }
    }
   
    # Set output directory to current script directory if not specified
    if ($OutputPath -eq "") {
        $fileName = "${ProjectNumber}_${BridgeId}_DRAFT.docx"
        $OutputPath = Join-Path $script:DefaultOutputDirectory $fileName
    }
   
    Write-ColorOutput "Input Path: $InputPath" "White"
   
    if ($ProjectNumber -and $BridgeId) {
        Write-ColorOutput "Project Number: $ProjectNumber" "White"
        Write-ColorOutput "Bridge ID: $BridgeId" "White"
    }
   
    Write-ColorOutput "Output Path: $OutputPath" "White"
    $result = Process-WordDocument -DocumentPath $InputPath -SavePath $OutputPath -ProjectNumber $ProjectNumber -BridgeId $BridgeId
   
    Write-ColorOutput "=== Processing Complete ===" "Cyan"
    Write-ColorOutput "Generated file: $result" "Green"
   
    # Generate VBA template code if requested
    if ($TemplatePath -like "*.dotm") {
        $vbaPath = New-WordTemplateWithVBA -TemplatePath $TemplatePath
        Write-ColorOutput "VBA template code generated at: $vbaPath" "Cyan"
    }
    
    # Log successful completion
    Write-LogFile -Message "=== SESSION COMPLETED SUCCESSFULLY ===" -Level "SUCCESS"
    Write-LogFile -Message "Output file: $result" -Level "SUCCESS"
   
} catch {
    Write-ColorOutput "Script failed: $($_.Exception.Message)" "Red"
    Write-ColorOutput "Stack trace: $($_.ScriptStackTrace)" "Red"
    
    # Log error details
    Write-LogFile -Message "=== SESSION FAILED ===" -Level "ERROR"
    Write-LogFile -Message "Error: $($_.Exception.Message)" -Level "ERROR"
    Write-LogFile -Message "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    
    exit 1
}