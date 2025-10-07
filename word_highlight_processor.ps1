# Word Document Highlight Processor 
# This script opens a Word template, removes yellow highlights, adds highlights for <<*>> patterns, and saves the document
   
param(
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ""
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to process Word document highlights
function Process-WordHighlights {
    param(
        [string]$DocumentPath,
        [string]$SavePath = ""
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
        
        # Step 2: Find and highlight text matching <<*>> pattern
        Write-ColorOutput "Adding highlights for merge field patterns..." "Yellow"
        
        # Reset the range for new search
        $Range = $Document.Range()
        $Range.Find.ClearFormatting()
        $Range.Find.Replacement.ClearFormatting()
        
        # Set up the search for <<*>> pattern
        $Range.Find.Text = "<<*>>"
        $Range.Find.MatchWildcards = $true
        $Range.Find.Replacement.Highlight = $true
        $Range.Find.Replacement.Text = "^&"  # Replace with same text but highlighted
        
        # Execute the replacement
        $replaceCount = $Range.Find.Execute("", $false, $false, $false, $false, $false, $true, 1, $false, "", 2)
        
        Write-ColorOutput "Merge field patterns highlighted" "Green"
        
        # Step 3: Save the document
        if ($SavePath -eq "") {
            # Generate output path if not provided
            $fileInfo = Get-Item $DocumentPath
            $SavePath = Join-Path $fileInfo.DirectoryName ($fileInfo.BaseName + "_processed" + $fileInfo.Extension)
        }
        
        Write-ColorOutput "Saving document to: $SavePath" "Yellow"
        
        # Save as new document
        $Document.SaveAs2($SavePath)
        
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

# Main execution
try {
    Write-ColorOutput "=== Word Document Highlight Processor ===" "Cyan"
    Write-ColorOutput "Input Path: $InputPath" "White"
    
    if ($OutputPath -ne "") {
        Write-ColorOutput "Output Path: $OutputPath" "White"
        $result = Process-WordHighlights -DocumentPath $InputPath -SavePath $OutputPath
    } else {
        $result = Process-WordHighlights -DocumentPath $InputPath
    }
    
    Write-ColorOutput "=== Processing Complete ===" "Cyan"
    
} catch {
    Write-ColorOutput "Script failed: $($_.Exception.Message)" "Red"
    exit 1
}