# Word Document Highlight Processor with Database Integration
# This script opens a Word template, processes highlights, integrates with database for project selection,
# and handles mail merge functionality with project and bridge ID selection

param(
    [Parameter(Mandatory=$false)]
    [string]$InputPath = "C:\Users\user\OneDrive\Documents\Work_Files\Word Doc to database\HandHPS.docx",
   
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "",
   
    [Parameter(Mandatory=$false)]
    [string]$TemplatePath = "C:\Users\user\OneDrive\Documents\Work_Files\Word Doc to database\HandHPS.docx",
   
    [Parameter(Mandatory=$false)] 
    [string]$ProjectNumber = "H-025-999-25",
   
    [Parameter(Mandatory=$false)]
    [string]$BridgeId = "K160.0",
   
    [Parameter(Mandatory=$false)] 
    [switch]$ShowProjectPicker = $false
)

# Set default paths based on current script location
$script:ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DefaultTemplatePath = Join-Path $script:ScriptDirectory "HandHPS.docx"
$script:DefaultOutputDirectory = $script:ScriptDirectory

# Use current directory template if not specified
if ($TemplatePath -eq "") {
    $TemplatePath = $script:DefaultTemplatePath
}

# Database connection string
$script:ConnectionString = "Driver={ODBC Driver 18 for SQL Server};" +
                          "Server=tcp:mendrop.database.windows.net;" +
                          "Database=MendropReportServer;" +
                          "UID=ReportUser;" +
                          "PWD=R3p0rtUs3r!;"

# Logging configuration
$script:LogFilePath = Join-Path $script:ScriptDirectory "word_processor_log.txt"
$script:LoggingEnabled = $true

# Highlight color configuration
$script:HighlightColorIndex = 4              # Green for merged values
$script:OrangeHighlightColorIndex = 14       # Dark Yellow (proxy for Orange) for placeholders

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
            Add-Content -Path $script:LogFilePath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { }
    }
}

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
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
        $recordset = $connection.Execute("SELECT DISTINCT project_number FROM [dbo].[vwHandHReportFormFields]")
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

# Function to get bridge IDs for a specific project
function Get-BridgeIds {
    param([string]$ProjectNumber)
    try {
        Write-ColorOutput "Retrieving bridge IDs for project: $ProjectNumber" "Yellow"
        $connection = Get-DatabaseConnection
        $query = "SELECT DISTINCT bridge_id FROM [dbo].[vwHandHReportFormFields] WHERE project_number = '$ProjectNumber'"
        $recordset = $connection.Execute($query)
        $bridgeIds = @()
        while (-not $recordset.EOF) {
            $bridgeIds += $recordset.Fields("bridge_id").Value
            $recordset.MoveNext()
        }
        $recordset.Close()
        $connection.Close()
        Write-ColorOutput "Retrieved $($bridgeIds.Count) bridge IDs" "Green"
        return $bridgeIds
    } catch {
        Write-ColorOutput "Error retrieving bridge IDs: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to show project picker dialog
function Show-ProjectPicker {
    try {
        Write-ColorOutput "Loading project picker..." "Yellow"
        $projects = Get-ProjectNumbers
        if ($projects.Count -eq 0) {
            Write-ColorOutput "No projects found in database" "Red"
            return $null
        }
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Select Project and Bridge ID"
        $form.Size = New-Object System.Drawing.Size(400, 300)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
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
        $projectCombo.Add_SelectedIndexChanged({
            $bridgeCombo.Items.Clear()
            $bridgeCombo.Enabled = $false
            if ($projectCombo.SelectedItem) {
                $bridgeIds = Get-BridgeIds -ProjectNumber $projectCombo.SelectedItem
                $bridgeIds | ForEach-Object { if ($_ -and $_ -ne "") { $bridgeCombo.Items.Add($_) } }
                $bridgeCombo.Enabled = ($bridgeCombo.Items.Count -gt 0)
            }
        })
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(200, 120)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Text = "OK"
        $okButton.DialogResult = "OK"
        $okButton.Enabled = $false
        $form.Controls.Add($okButton)
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(295, 120)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = "Cancel"
        $form.Controls.Add($cancelButton)
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
        $connection = Get-DatabaseConnection
        $query = "SELECT * FROM [dbo].[vwHandHReportFormFields] WHERE project_number = '$ProjectNumber' AND bridge_id = '$BridgeId'"
        Write-ColorOutput "Executing query: $query" "Yellow"
        $recordset = $connection.Execute($query)
        if ($recordset.EOF) {
            Write-ColorOutput "No data found for the specified project and bridge" "Red"
            $recordset.Close()
            $connection.Close()
            return $false
        }
        $fieldData = @{}
        for ($i = 0; $i -lt $recordset.Fields.Count; $i++) {
            $fieldName = $recordset.Fields($i).Name
            try {
                $fieldValue = $recordset.Fields($i).Value
                if ($fieldValue -eq $null -or [System.DBNull]::Value.Equals($fieldValue)) {
                    $fieldValue = ""
                } else {
                    $fieldValue = $fieldValue.ToString()
                }
            } catch { $fieldValue = "" }
            $fieldData[$fieldName] = $fieldValue
        }
        $recordset.Close()
        $connection.Close()
        Write-ColorOutput "Retrieved $($fieldData.Count) fields from database" "Green"
        
        $replacedFields = 0
        $replacedFieldNames = @()
        $missingFields = @()
        $emptyFields = @()
        
        try {
            $fields = $Document.Fields
            Write-ColorOutput "Found $($fields.Count) fields in document" "Cyan"
            $mergeFields = @()
            foreach ($field in $fields) {
                try {
                    if ($field.Type -eq 59) { # wdFieldMergeField
                        $mergeFields += $field
                    }
                } catch { Write-ColorOutput "Error analyzing field: $($_.Exception.Message)" "Red" }
            }
            Write-ColorOutput "MERGEFIELD count: $($mergeFields.Count)" "Cyan"
            
            foreach ($field in $mergeFields) {
                try {
                    $fieldCode = $field.Code.Text
                    $fieldName = $null
                    if ($fieldCode -match 'MERGEFIELD\s+"([^"]+)"') {
                        $fieldName = $matches[1]
                    } elseif ($fieldCode -match 'MERGEFIELD\s+([^\s\\]+)') {
                        $fieldName = $matches[1]
                    }
                    if (-not $fieldName) { Write-ColorOutput "Could not extract field name from: $fieldCode" "Yellow"; continue }
                    if (-not $fieldData.ContainsKey($fieldName)) {
                        try { $field.Result.HighlightColorIndex = $script:OrangeHighlightColorIndex } catch {}
                        $missingFields += $fieldName
                        Write-ColorOutput "MERGEFIELD '$fieldName' not found in database - highlighted orange" "Yellow"
                        continue
                    }
                    $value = $fieldData[$fieldName]
                    if ($value -match '^\s*NULL\s*$') {
                        $resultRange = $field.Result.Duplicate
                        $resultRange.Text = $value
                        try { $field.Unlink() } catch {
                            try { $field.Select(); $Document.Application.Selection.TypeText($value) } catch { }
                        }
                        try { $resultRange.HighlightColorIndex = $script:OrangeHighlightColorIndex } catch {
                            try { $resultRange.Font.HighlightColorIndex = $script:OrangeHighlightColorIndex } catch { }
                        }
                        $replacedFields++; $replacedFieldNames += $fieldName
                        Write-ColorOutput "MERGEFIELD '$fieldName' value is NULL - replaced and highlighted orange" "Yellow"
                        continue
                    }
                    if ([string]::IsNullOrWhiteSpace($value)) {
                        try { $field.Result.HighlightColorIndex = $script:OrangeHighlightColorIndex } catch {}
                        $emptyFields += $fieldName
                        Write-ColorOutput "MERGEFIELD '$fieldName' is empty in database - highlighted orange" "Yellow"
                        continue
                    }
                    $resultRange = $field.Result.Duplicate
                    $resultRange.Text = $value
                    try { $field.Unlink() } catch {
                        try { $field.Select(); $Document.Application.Selection.TypeText($value) } catch {
                            Write-ColorOutput "Fallback selection replacement failed for '$fieldName': $($_.Exception.Message)" "Red"
                            continue
                        }
                    }
                    try { $resultRange.HighlightColorIndex = $script:HighlightColorIndex } catch {
                        try { $resultRange.Font.HighlightColorIndex = $script:HighlightColorIndex } catch { }
                    }
                    $replacedFields++; $replacedFieldNames += $fieldName
                    Write-ColorOutput "Replaced MERGEFIELD '$fieldName' with '$value' and unlinked field" "Green"
                } catch {
                    Write-ColorOutput "Error processing MERGEFIELD: $($_.Exception.Message)" "Red"
                    continue
                }
            }
            if ($replacedFields -gt 0) {
                Write-ColorOutput "Successfully replaced $replacedFields MERGEFIELD codes" "Green"
                Write-ColorOutput "Fields replaced: $($replacedFieldNames -join ', ')" "Cyan"
            } else {
                Write-ColorOutput "No MERGEFIELD codes were found or replaced" "Yellow"
            }
            if ($missingFields.Count -gt 0) {
                Write-ColorOutput "Unmatched fields (not in database, highlighted orange): $($missingFields -join ', ')" "Yellow"
            }
            if ($emptyFields.Count -gt 0) {
                Write-ColorOutput "Empty fields (empty/NULL in database, highlighted orange): $($emptyFields -join ', ')" "Yellow"
            }

            # Fallback visible markers replacement for merged values only (keeps green highlighting for actual replacements)
            Write-ColorOutput "Attempting fallback text replacement..." "Yellow"
            foreach ($fieldName in $fieldData.Keys) {
                $fieldValue = $fieldData[$fieldName]
                if ([string]::IsNullOrWhiteSpace($fieldValue)) { continue }
                $patterns = @("{ MERGEFIELD $fieldName }","«$fieldName»","<<$fieldName>>")
                foreach ($pattern in $patterns) {
                    $range = $Document.Content
                    $range.Find.ClearFormatting()
                    $range.Find.Replacement.ClearFormatting()
                    $range.Find.Text = $pattern
                    $range.Find.Replacement.Text = $fieldValue
                    $range.Find.Forward = $true
                    $range.Find.Wrap = 1
                    $range.Find.Format = $false
                    $range.Find.MatchCase = $false
                    $range.Find.MatchWholeWord = $false
                    $found = $range.Find.Execute($null,$false,$false,$false,$false,$false,$true,1,$true,$fieldValue,2)
                    if ($found) {
                        $hlRange = $Document.Content
                        $hlRange.Find.Text = $fieldValue
                        $hlRange.Find.Execute()
                        try { $hlRange.HighlightColorIndex = $script:HighlightColorIndex } catch { }
                        $replacedFields++; $replacedFieldNames += $fieldName
                        Write-ColorOutput "Replaced pattern '$pattern' with '$fieldValue' (fallback)" "Green"
                        break
                    }
                }
            }
            Write-ColorOutput "Fallback replacement completed: $replacedFields fields" "Green"
        } catch {
            Write-ColorOutput "Error during field processing: $($_.Exception.Message)" "Red"
            return $false
        }
        return $true
    } catch {
        Write-ColorOutput "Error during mail merge: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to highlight unmerged placeholders using multiple reliable methods
function Highlight-UnmergedPlaceholders {
    param([object]$Document)
    try {
        Write-ColorOutput "Highlighting unmerged placeholders using multiple methods..." "Yellow"
        $highlightedCount = 0
        
        # Method 1: Iterate through all story ranges and use text comparison
        $storyRanges = @($Document.StoryRanges)
        foreach ($storyRange in $storyRanges) {
            try {
                if ($storyRange -eq $null) { continue }
                
                $text = $storyRange.Text
                if ([string]::IsNullOrEmpty($text)) { continue }
                
                # Look for guillemet patterns «...»
                $startIndex = 0
                while (($startIndex -lt $text.Length) -and ($startIndex -ne -1)) {
                    $openPos = $text.IndexOf('«', $startIndex)
                    if ($openPos -eq -1) { break }
                    
                    $closePos = $text.IndexOf('»', $openPos + 1)
                    if ($closePos -eq -1) { break }
                    
                    # Found a potential placeholder
                    $placeholderLength = $closePos - $openPos + 1
                    if ($placeholderLength -gt 2) { # At least «x»
                        try {
                            $placeholderRange = $storyRange.Duplicate
                            $placeholderRange.Start = $storyRange.Start + $openPos
                            $placeholderRange.End = $storyRange.Start + $closePos + 1
                            
                            # Verify it's actually a placeholder (not part of regular text)
                            $placeholderText = $placeholderRange.Text
                            if ($placeholderText -match '^«[^»]+»$') {
                                $placeholderRange.HighlightColorIndex = $script:OrangeHighlightColorIndex
                                $highlightedCount++
                                Write-ColorOutput "Highlighted guillemet placeholder: '$placeholderText'" "Yellow"
                            }
                        } catch {
                            # Continue with next placeholder
                        }
                    }
                    $startIndex = $closePos + 1
                }
                
                # Look for double-angle patterns <<...>>
                $startIndex = 0
                while (($startIndex -lt $text.Length) -and ($startIndex -ne -1)) {
                    $openPos = $text.IndexOf('<<', $startIndex)
                    if ($openPos -eq -1) { break }
                    
                    $closePos = $text.IndexOf('>>', $openPos + 2)
                    if ($closePos -eq -1) { break }
                    
                    # Found a potential placeholder
                    $placeholderLength = $closePos - $openPos + 2
                    if ($placeholderLength -gt 4) { # At least <<x>>
                        try {
                            $placeholderRange = $storyRange.Duplicate
                            $placeholderRange.Start = $storyRange.Start + $openPos
                            $placeholderRange.End = $storyRange.Start + $closePos + 2
                            
                            # Verify it's actually a placeholder
                            $placeholderText = $placeholderRange.Text
                            if ($placeholderText -match '^<<[^>]+>>$') {
                                $placeholderRange.HighlightColorIndex = $script:OrangeHighlightColorIndex
                                $highlightedCount++
                                Write-ColorOutput "Highlighted double-angle placeholder: '$placeholderText'" "Yellow"
                            }
                        } catch {
                            # Continue with next placeholder
                        }
                    }
                    $startIndex = $closePos + 2
                }
                
                # Clean up
                if ($storyRange -ne $Document.StoryRanges.First) {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($storyRange) | Out-Null
                }
            } catch {
                Write-ColorOutput "Warning: Error processing story range: $($_.Exception.Message)" "DarkYellow"
            }
        }
        
        # Method 2: Try simple Find without wildcards as backup
        try {
            Write-ColorOutput "Trying backup Find method without wildcards..." "Yellow"
            $range = $Document.Range()
            $find = $range.Find
            
            # Search for individual guillemet characters and expand
            $find.Text = "«"
            $find.MatchWildcards = $false
            $find.Forward = $true
            $find.Wrap = 1
            
            while ($find.Execute()) {
                if ($find.Found) {
                    try {
                        $foundRange = $find.Parent.Duplicate
                        # Try to expand to find the closing »
                        $originalEnd = $foundRange.End
                        $foundRange.End = $foundRange.End + 100 # Look ahead a bit
                        $closingPos = $foundRange.Text.IndexOf('»')
                        if ($closingPos -ne -1) {
                            $foundRange.End = $foundRange.Start + $closingPos + 1
                            $placeholderText = $foundRange.Text
                            if ($placeholderText -match '^«[^»]+»$') {
                                $foundRange.HighlightColorIndex = $script:OrangeHighlightColorIndex
                                $highlightedCount++
                                Write-ColorOutput "Backup method highlighted: '$placeholderText'" "Yellow"
                            }
                        }
                    } catch {
                        # Continue with next find
                    }
                }
            }
        } catch {
            Write-ColorOutput "Backup Find method failed: $($_.Exception.Message)" "DarkYellow"
        }
        
        if ($highlightedCount -gt 0) {
            Write-ColorOutput "Successfully highlighted $highlightedCount unmerged placeholders in orange" "Green"
        } else {
            Write-ColorOutput "No unmerged placeholders found" "Yellow"
        }
        return $highlightedCount
    } catch {
        Write-ColorOutput "Error during placeholder highlighting: $($_.Exception.Message)" "Red"
        return 0
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
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        $Word.DisplayAlerts = 0
        Write-ColorOutput "Opening document: $DocumentPath" "Yellow"
        if (-not (Test-Path $DocumentPath)) { throw "File not found: $DocumentPath" }
        $Document = $Word.Documents.Open($DocumentPath)
        Write-ColorOutput "Document opened successfully" "Green"

        # Clear existing highlights
        Write-ColorOutput "Removing yellow highlights..." "Yellow"
        try {
            $storyRangeTypes = 1..12
            foreach ($type in $storyRangeTypes) {
                try {
                    $rng = $Document.StoryRanges.Item($type)
                    while ($rng -ne $null) {
                        try { $rng.HighlightColorIndex = 0 } catch { }
                        $rng = $rng.NextStoryRange
                    }
                } catch { }
            }
        } catch {
            $Range = $Document.Range()
            try { $Range.HighlightColorIndex = 0 } catch { }
        }
        Write-ColorOutput "Yellow highlights removed" "Green"
       
        # Mail merge
        if ($ProjectNumber -and $BridgeId) {
            try {
                Write-ColorOutput "Attempting database connection for mail merge..." "Yellow"
                $mergeSuccess = Invoke-MailMerge -Document $Document -ProjectNumber $ProjectNumber -BridgeId $BridgeId
                if (-not $mergeSuccess) {
                    Write-ColorOutput "Mail merge failed - no data found for project/bridge" "Yellow"
                    Write-ColorOutput "Template layout preserved - merge fields will be highlighted for manual completion" "Cyan"
                }
            } catch {
                Write-ColorOutput "Database connection failed - preserving template content" "Yellow"
                Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
                Write-ColorOutput "Template content preserved - merge fields will be highlighted for manual completion" "Cyan"
            }
        } else {
            Write-ColorOutput "No project data provided - preserving all template content" "Yellow"
            Write-ColorOutput "You can run with -ShowProjectPicker to select from DB" "Cyan"
        }
       
        # Highlight remaining placeholders in orange using reliable method
        try {
            Write-ColorOutput "Highlighting unmerged placeholders..." "Yellow"
            $highlightedCount = Highlight-UnmergedPlaceholders -Document $Document
            if ($highlightedCount -gt 0) {
                Write-ColorOutput "Highlighted $highlightedCount unmerged placeholders in orange" "Green"
            }
        } catch {
            Write-ColorOutput "Error highlighting unmerged placeholders: $($_.Exception.Message)" "Red"
        }
       
        Write-ColorOutput "Skipping legacy text-placeholder highlighting; only replaced text is highlighted" "Cyan"
       
        # Save output
        if ($SavePath -eq "") {
            $fileInfo = Get-Item $DocumentPath
            if ($ProjectNumber -and $BridgeId) {
                $fileName = "${ProjectNumber}_${BridgeId}_DRAFT.docx"
                $SavePath = Join-Path $fileInfo.DirectoryName $fileName
            } else {
                $extension = if ($fileInfo.Extension -eq ".dotm" -or $fileInfo.Extension -eq ".dotx") { ".docx" } else { $fileInfo.Extension }
                $SavePath = Join-Path $fileInfo.DirectoryName ($fileInfo.BaseName + "_processed" + $extension)
            }
        }
        Write-ColorOutput "Saving document to: $SavePath" "Yellow"
        $Document.SaveAs2($SavePath, 16) # wdFormatDocumentDefault
        Write-ColorOutput "Document saved successfully" "Green"
        $Document.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-ColorOutput "Processing completed successfully!" "Green"
        Write-ColorOutput "Output file: $SavePath" "Cyan"
        return $SavePath
    } catch {
        Write-ColorOutput "Error occurred: $($_.Exception.Message)" "Red"
        if ($Document) { try { $Document.Close($false) } catch { } }
        if ($Word) { try { $Word.Quit() } catch { } }
        throw
    }
}

# Function to create Word template with VBA integration (optional helper)
function New-WordTemplateWithVBA {
    param([string]$TemplatePath = "handhps.dotm")
    try {
        Write-ColorOutput "Creating Word template with VBA integration..." "Yellow"
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
   
    connectionString = "Driver={ODBC Driver 18 for SQL Server};" & _
                      "Server=tcp:mendrop.database.windows.net;" & _
                      "Database=MendropReportServer;" & _
                      "UID=ReportUser;" & _
                      "PWD=R3p0rtUs3r!;"
   
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connectionString
   
    Set rs = conn.Execute("SELECT DISTINCT project_number FROM [dbo].[vwHandHReportFormFields] ORDER BY project_number")
   
    frmProjectPicker.ComboBox1.Clear
    Do While Not rs.EOF
        frmProjectPicker.ComboBox1.AddItem rs.Fields("project_number").Value
        rs.MoveNext
    Loop
   
    frmProjectPicker.Show
   
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

Private Sub Document_Open()
    PickProjectReport
End Sub
"@
        Write-ColorOutput "VBA Code generated. This needs to be manually added to the Word template." "Green"
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
    Write-LogFile -Message "=== NEW SESSION STARTED ===" -Level "INFO"
    Write-LogFile -Message "Script: word_marco_dotm.ps1" -Level "INFO"
    Write-LogFile -Message "User: $env:USERNAME" -Level "INFO"
    Write-LogFile -Message "Computer: $env:COMPUTERNAME" -Level "INFO"
    Write-LogFile -Message "Working Directory: $(Get-Location)" -Level "INFO"
    Write-LogFile -Message "Log File: $script:LogFilePath" -Level "INFO"
    
    Write-ColorOutput "=== Word Document Processor with Database Integration ===" "Cyan"
   
    if ($TemplatePath -eq "" -or $ProjectNumber -eq "" -or $BridgeId -eq "") {
        Write-ColorOutput "Interactive mode: Please provide the following information..." "Yellow"
        if ($TemplatePath -eq "") {
            do {
                $TemplatePath = Read-Host "Enter the path to the .dotm/.docx template file (or press Enter for default)"
                if ($TemplatePath -eq "") { $TemplatePath = $script:DefaultTemplatePath }
                if (-not (Test-Path $TemplatePath)) { Write-ColorOutput "File not found: $TemplatePath" "Red"; $TemplatePath = "" }
            } while ($TemplatePath -eq "")
        }
        if ($ProjectNumber -eq "") {
            do {
                $ProjectNumber = Read-Host "Enter the Project Number (e.g., H-025-999-25)"
                if ($ProjectNumber -eq "") { Write-ColorOutput "Project Number is required!" "Red" }
            } while ($ProjectNumber -eq "")
        }
        if ($BridgeId -eq "") {
            do {
                $BridgeId = Read-Host "Enter the Bridge ID (e.g., K160.0)"
                if ($BridgeId -eq "") { Write-ColorOutput "Bridge ID is required!" "Red" }
            } while ($BridgeId -eq "")
        }
        Write-ColorOutput "Using Template: $TemplatePath" "Green"
        Write-ColorOutput "Project Number: $ProjectNumber" "Green"
        Write-ColorOutput "Bridge ID: $BridgeId" "Green"
    }
   
    if ($ShowProjectPicker) {
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
   
    if ($TemplatePath -like "*.dotm") {
        $vbaPath = New-WordTemplateWithVBA -TemplatePath $TemplatePath
        Write-ColorOutput "VBA template code generated at: $vbaPath" "Cyan"
    }
    
    Write-LogFile -Message "=== SESSION COMPLETED SUCCESSFULLY ===" -Level "SUCCESS"
    Write-LogFile -Message "Output file: $result" -Level "SUCCESS"
   
} catch {
    Write-ColorOutput "Script failed: $($_.Exception.Message)" "Red"
    Write-ColorOutput "Stack trace: $($_.ScriptStackTrace)" "Red"
    Write-LogFile -Message "=== SESSION FAILED ===" -Level "ERROR"
    Write-LogFile -Message "Error: $($_.Exception.Message)" -Level "ERROR"
    Write-LogFile -Message "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
    exit 1
}