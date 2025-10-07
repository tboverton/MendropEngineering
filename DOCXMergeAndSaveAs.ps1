# Word Template Mail Merge Automation with Azure SQL
# This script automates the process of:
# 1. Opening a Word template (.dotx) from a specified location
# 2. Converting it to .docx to avoid template behavior issues
# 3. Connecting to Azure SQL database and performing mail merge
# 4. Highlighting populated fields in grey and missing fields in yellow
# 5. Saving the final document with a timestamp
 
# ===== Configuration =====
$ErrorActionPreference = 'Stop'
# Use Join-Path and $PSScriptRoot to build clean paths (avoid accidental string concatenation)
$logFile = Join-Path -Path $PSScriptRoot -ChildPath 'word_merge_log.txt'

# Template and output settings
$templatePath = Join-Path -Path $PSScriptRoot -ChildPath 'HandHPS.docx'
$outputFolder = $PSScriptRoot
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$tempDocxPath = Join-Path -Path $env:TEMP -ChildPath "temp_merged_report_$timestamp.docx"
$finalOutputPath = Join-Path -Path $outputFolder -ChildPath "merged_report_$timestamp.docx"
 
# Azure SQL Connection Settings
$sqlServer = "mendrop.database.windows.net"
$sqlDatabase = "MendropReportServer"
$sqlUsername = "ReportUser"
$sqlPassword = "R3p0rtUs3r!" # For production, use secure credential management
 
# ===== Functions =====
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
   
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$level] $message"
   
    # Write to console
    if ($level -eq "ERROR") {
        Write-Host $logMessage -ForegroundColor Red
    } elseif ($level -eq "WARNING") {
        Write-Host $logMessage -ForegroundColor Yellow
    } else {
        Write-Host $logMessage
    }
   
    # Append to log file
    Add-Content -Path $logFile -Value $logMessage
}
 
function Initialize-Environment {
    # Create output folder if it doesn't exist
    if (-not (Test-Path -Path $outputFolder)) {
        New-Item -Path $outputFolder -ItemType Directory | Out-Null
        Write-Log "Created output folder: $outputFolder"
    }
   
    # Check if template exists
    if (-not (Test-Path -Path $templatePath)) {
        Write-Log "Template file not found at: $templatePath" -level "ERROR"
        throw "Template file not found"
    }
}
 
function Get-SecureCredentials {
    # In production, replace this with Azure Key Vault or other secure credential management
    # For demo purposes, we'll use a simple prompt if password is empty
    if ([string]::IsNullOrEmpty($sqlPassword)) {
        $securePassword = Read-Host "Enter SQL password for $sqlUsername" -AsSecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $sqlPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    }
   
    return $sqlPassword
}
 
function Open-WordDocument {
    try {
        Write-Log "Starting Microsoft Word..."
        $word = New-Object -ComObject Word.Application
        $word.Visible = $true
       
        Write-Log "Opening template: $templatePath"
        $doc = $word.Documents.Open($templatePath)
       
        Write-Log "Converting template to docx format..."
        # Use a small watchdog job to prevent SaveAs2 from hanging the script indefinitely.
        $marker = Join-Path -Path $env:TEMP -ChildPath "word_save_inprogress_$timestamp.flag"
        New-Item -Path $marker -ItemType File -Force | Out-Null
        # If Save does not complete within $saveTimeout seconds, kill WINWORD to recover.
        $saveTimeout = 30
        $watchJob = Start-Job -ArgumentList $marker, $saveTimeout -ScriptBlock {
            param($m, $t)
            Start-Sleep -Seconds $t
            if (Test-Path $m) {
                try { Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
                Remove-Item -Path $m -ErrorAction SilentlyContinue
            }
        }
        try {
            # Try to disable alerts; some Word hosts reject numeric values for DisplayAlerts
            try { $word.DisplayAlerts = 0 } catch { Write-Log "Warning: couldn't set DisplayAlerts - continuing. $_" -level "WARNING" }

            try {
                $doc.SaveAs($tempDocxPath, 17) # 17 = wdFormatDocumentDefault (.docx)
            }
            finally {
                if (Test-Path $marker) { Remove-Item -Path $marker -ErrorAction SilentlyContinue }
                if ($watchJob) {
                    # Don't block waiting for the job to finish its Sleep; stop and remove it if still running
                    try {
                        if ($watchJob.State -eq 'Running') { Stop-Job -Id $watchJob.Id -Force -ErrorAction SilentlyContinue }
                    } catch {}
                    try { Remove-Job -Id $watchJob.Id -ErrorAction SilentlyContinue } catch {}
                }
            }
        }
        catch {
            Write-Log "Error during template SaveAs: $_" -level "ERROR"
            throw
        }

        $doc.Close()
       
        Write-Log "Opening converted docx file..."
        $doc = $word.Documents.Open($tempDocxPath)
       
        return $word, $doc
    }
    catch {
        Write-Log "Error opening Word document: $_" -level "ERROR"
        if ($word) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        throw
    }
}
 
function Connect-ToAzureSQL {
    param (
        [string]$server,
        [string]$database,
        [string]$username,
        [System.Security.SecureString]$securePassword
    )
   
    try {
        Write-Log "Connecting to Azure SQL database..."
       
        # Convert secure password to plain text for connection string
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
       
        # Build connection string
        $connectionString = "Server=tcp:$server,1433;Initial Catalog=$database;Persist Security Info=False;User ID=$username;Password=$password;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
       
        # For production, consider using System.Data.SqlClient for more robust connection handling
        return $connectionString
    }
    catch {
        Write-Log "Error connecting to Azure SQL: $_" -level "ERROR"
        throw
    }
}
 
function Start-MailMerge {
    param (
        $wordDoc,
        [string]$connectionString
    )
   
    try {
        Write-Log "Setting up mail merge..."
       
        # Set up mail merge
        $wordDoc.MailMerge.MainDocumentType = 0 # wdFormLetters
       
        # Set up SQL connection for mail merge
        $wordDoc.MailMerge.OpenDataSource("", "OLEDB;$connectionString", "SELECT * FROM [dbo].[vwHandHReportFormFields]") # Replace YourTableName with your actual table or view name
       
        Write-Log "Executing mail merge..."
        $wordDoc.MailMerge.Execute()
       
        Write-Log "Mail merge completed successfully"
        return $true
    }
    catch {
        Write-Log "Error performing mail merge: $_" -level "ERROR"
        return $false
    }
}
 
function Set-MergeFieldHighlighting {
    param (
        $wordDoc
    )
   
    try {
        Write-Log "Highlighting merge fields..."
       
        # Constants for Word colors
        $wdColorYellow = 7
        $wdColorGray25 = 16
       
        # Process each field in the document
        $fieldCount = 0
        $emptyFieldCount = 0
       
        foreach ($field in $wordDoc.Fields) {
            # Check if it's a MERGEFIELD
            if ($field.Type -eq 15) { # 15 = wdFieldMergeField
                $fieldCount++
                $text = $field.Result.Text.Trim()
               
                if ([string]::IsNullOrWhiteSpace($text)) {
                    # Highlight empty fields in yellow
                    $field.Result.Shading.BackgroundPatternColor = $wdColorYellow
                    $emptyFieldCount++
                }
                else {
                    # Highlight populated fields in grey
                    $field.Result.Shading.BackgroundPatternColor = $wdColorGray25
                }
            }
        }
       
        # Also look for any <<Field>> patterns that might be missing
        $findObject = $wordDoc.Content.Find
        $findObject.ClearFormatting()
        $findObject.Text = "<<*>>"
        $findObject.Forward = $true
        $findObject.MatchWildcards = $true
       
        $missingFieldCount = 0
       
        while ($findObject.Execute()) {
            $missingFieldCount++
            $wordDoc.Range($findObject.Found.Start, $findObject.Found.End).Shading.BackgroundPatternColor = $wdColorYellow
        }
       
        Write-Log "Field highlighting complete. Total fields: $fieldCount, Empty fields: $emptyFieldCount, Missing fields: $missingFieldCount"
       
        if ($emptyFieldCount -gt 0 -or $missingFieldCount -gt 0) {
            Write-Log "WARNING: Document contains empty or missing fields" -level "WARNING"
        }
    }
    catch {
        Write-Log "Error highlighting merge fields: $_" -level "ERROR"
    }
}
 
function Save-FinalDocument {
    param (
        $wordDoc,
        [string]$outputPath
    )
   
    try {
        Write-Log "Saving final document to: $outputPath"
        # Use watchdog job for final save as well, and disable alerts to avoid modal prompts
        $marker = Join-Path -Path $env:TEMP -ChildPath "word_save_inprogress_final_$timestamp.flag"
        New-Item -Path $marker -ItemType File -Force | Out-Null
        $saveTimeout = 30
        $watchJob = Start-Job -ArgumentList $marker, $saveTimeout -ScriptBlock {
            param($m, $t)
            Start-Sleep -Seconds $t
            if (Test-Path $m) {
                try { Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
                Remove-Item -Path $m -ErrorAction SilentlyContinue
            }
        }
        try {
            try { $wordDoc.Application.DisplayAlerts = 0 } catch { Write-Log "Warning: couldn't set DisplayAlerts on final save - continuing. $_" -level "WARNING" }

            try {
                $wordDoc.SaveAs($outputPath)
            }
            finally {
                if (Test-Path $marker) { Remove-Item -Path $marker -ErrorAction SilentlyContinue }
                if ($watchJob) {
                    try {
                        if ($watchJob.State -eq 'Running') { Stop-Job -Id $watchJob.Id -Force -ErrorAction SilentlyContinue }
                    } catch {}
                    try { Remove-Job -Id $watchJob.Id -ErrorAction SilentlyContinue } catch {}
                }
            }

            Write-Log "Document saved successfully"
            return $true
        }
        catch {
            Write-Log "Error during final SaveAs2: $_" -level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error saving document: $_" -level "ERROR"
        return $false
    }
}
 
function Remove-Resources {
    param (
        $wordApp,
        $wordDoc
    )
   
    try {
        Write-Log "Cleaning up resources..."
       
        if ($wordDoc) {
            $wordDoc.Close([ref]$false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc) | Out-Null
        }
       
        if ($wordApp) {
            $wordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        }
       
        # Remove temp file
        if (Test-Path -Path $tempDocxPath) {
            Remove-Item -Path $tempDocxPath -Force
        }
       
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
       
        Write-Log "Cleanup completed"
    }
    catch {
        Write-Log "Error during cleanup: $_" -level "WARNING"
    }
}
 
# ===== Main Script Execution =====
try {
    Write-Log "=== Word Template Mail Merge Automation Started ==="
   
    # Initialize environment
    Initialize-Environment
   
    # Get secure credentials
    $sqlPassword = Get-SecureCredentials
    $securePassword = ConvertTo-SecureString -String $sqlPassword -AsPlainText -Force
   
    # Open Word and document
    $word, $doc = Open-WordDocument
   
    # Connect to Azure SQL
    $connectionString = Connect-ToAzureSQL -server $sqlServer -database $sqlDatabase -username $sqlUsername -securePassword $securePassword
   
    # Perform mail merge
    $mergeSuccess = Start-MailMerge -wordDoc $doc -connectionString $connectionString
   
    if ($mergeSuccess) {
        # Highlight merge fields
        Set-MergeFieldHighlighting -wordDoc $doc
       
        # Save final document
        Save-FinalDocument -wordDoc $doc -outputPath $finalOutputPath
    }
   
    Write-Log "=== Word Template Mail Merge Automation Completed ==="
}
catch {
    Write-Log "Critical error in mail merge automation: $_" -level "ERROR"
}
finally {
    # Always clean up resources
    if (Get-Variable -Name word -ErrorAction SilentlyContinue) {
        Remove-Resources -wordApp $word -wordDoc $doc
    }
}
 