param([string]$Target)
$errors = $null; $tokens = $null
[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $Target).Path, [ref]$tokens, [ref]$errors)
if ($errors) {
    $errors | Select-Object -Property Line,Column,Message | Format-Table -AutoSize
    exit 1
} else {
    Write-Host 'PARSE_OK'
    exit 0
}