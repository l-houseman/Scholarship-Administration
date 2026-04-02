
<# 
USRA 2026 — Offline organizer (CLM-compatible)
- No non-core .NET types, no [pscustomobject], no generic lists.
- Uses core cmdlets, strings, arrays, hashtables.
#>

# =========================== CONFIG ===========================
# Source folders (local)
$StudentAppDir = '\USRA 2026 Applications\Inputs\Student Application Forms'
$FacultyAppDir = '\USRA 2026 Applications\Inputs\Faculty Application Forms'
$TranscriptsDir = '\USRA 2026 Applications\Inputs\Student Transcripts'

# Destination organized root (local)
$OrganizedRoot = '\USRA 2026 Applications\USRA 2026_Organized'

# Year label used in standardized names
$YearLabel = '2026'

# Safety: start as DryRun. Set to $false to execute changes.
$DryRun = $false

# If a target file already exists, add " (1)", " (2)", ... instead of overwriting.
$AllowOverwrite = $false

# Optional: process only a subset (exact folder display names)
# $ProcessOnly = @('Student Transcripts')
$ProcessOnly = $null
# =============================================================

# Map: source folder display name -> standardized title prefix
$TitleMap = @{
    'Student Application Forms' = "$YearLabel USRA Student Application Form"
    'Faculty Application Forms' = "$YearLabel USRA Faculty Application Form"
    'Student Transcripts'       = "$YearLabel USRA Non-USask Transcript"
}

# Build source list (as items so we can get .Name)
$Sources = @()
if (Test-Path -LiteralPath $StudentAppDir) { $Sources += (Get-Item -LiteralPath $StudentAppDir) }
if (Test-Path -LiteralPath $FacultyAppDir) { $Sources += (Get-Item -LiteralPath $FacultyAppDir) }
if (Test-Path -LiteralPath $TranscriptsDir) { $Sources += (Get-Item -LiteralPath $TranscriptsDir) }

if ($ProcessOnly -ne $null) {
    $Sources = $Sources | Where-Object { $ProcessOnly -contains $_.Name }
}

# Prepare destination
if (-not $DryRun) {
    if (-not (Test-Path -LiteralPath $OrganizedRoot)) {
        New-Item -ItemType Directory -Path $OrganizedRoot -Force | Out-Null
    }
}

# Logs (plain CSV via strings; CLM-safe)
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath   = Join-Path (Split-Path $OrganizedRoot -Parent) ("USRA_2026_organize_log_{0}.csv" -f $timestamp)
$MissingReportPath = Join-Path (Split-Path $OrganizedRoot -Parent) ("USRA_2026_missing_report_{0}.csv" -f $timestamp)

# Write headers
'Timestamp,SourceFolder,OriginalName,StudentName,DocType,NewFileName,TargetPath,Action,Notes' | Out-File -FilePath $LogPath -Encoding UTF8 -Force

# Helper: CSV escape (double quotes + surround with quotes)
function To-CsvField([string]$s) {
    if ($null -eq $s) { return '""' }
    $escaped = $s -replace '"','""'
    return '"' + $escaped + '"'
}

# Helper: safe folder names (Windows)
function Safe-FolderName([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $t = $s.Trim()
    $t = $t -replace '[<>:"/\\|?*]','-'
    $t = ($t -replace '\s{2,}',' ') -replace '-{2,}','-'
    $t = $t -replace '[\.\s]+$',''
    if ($t -eq '') { return $null }
    return $t
}

# Helper: extract student from base name (string methods are allowed in CLM)
function Get-StudentNameFromBase([string]$base) {
    if ($null -eq $base) { return $null }
    $idx = $base.LastIndexOf('_')
    if ($idx -lt 0) { return $null }
    $student = $base.Substring($idx + 1).Trim(' ','_','-')
    if ([string]::IsNullOrWhiteSpace($student)) { return $null }
    return $student
}

# Helper: ensure unique destination path
function Get-UniquePath([string]$path, [bool]$AllowOverwriteLocal) {
    if ($AllowOverwriteLocal -or -not (Test-Path -LiteralPath $path)) { return $path }
    $dir = Split-Path -LiteralPath $path
    $name = (Get-Item -LiteralPath $path).BaseName 2>$null
    if (-not $name) {
        # fallback if file doesn't exist yet: compute base name manually
        $leaf = Split-Path -Leaf -Path $path
        $lastDot = $leaf.LastIndexOf('.')
        if ($lastDot -gt 0) { $name = $leaf.Substring(0,$lastDot) } else { $name = $leaf }
    }
    $ext  = [System.IO.Path]::GetExtension($path) 2>$null
    if (-not $ext) {
        $leaf = Split-Path -Leaf -Path $path
        $lastDot = $leaf.LastIndexOf('.')
        if ($lastDot -gt 0) { $ext = $leaf.Substring($lastDot) } else { $ext = '' }
    }
    $i = 1
    $candidate = $path
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $dir ("{0} ({1}){2}" -f $name, $i, $ext)
        $i++
    }
    return $candidate
}

# Track document presence (hashtable of hashtables)
$StudentDocs = @{}  # $StudentDocs['Leah Houseman'] = @{ 'Student Application Forms'=$true; ... }

# Main loop
foreach ($srcFolder in $Sources) {
    $srcName  = $srcFolder.Name
    $docTitle = $TitleMap[$srcName]

    # Enumerate files
    Get-ChildItem -LiteralPath $srcFolder.FullName -Recurse -File | ForEach-Object {
        $f = $_
        # Use built-in properties to avoid .NET static calls
        $base = $f.BaseName
        $ext  = $f.Extension  # includes the leading dot (or empty)

        $student = Get-StudentNameFromBase $base

        if (-not $docTitle) {
            $line = (To-CsvField (Get-Date)) + ',' + (To-CsvField $srcFolder.FullName) + ',' +
                    (To-CsvField $f.Name) + ',' + (To-CsvField '') + ',' +
                    (To-CsvField '') + ',' + (To-CsvField '') + ',' +
                    (To-CsvField '') + ',' + (To-CsvField 'Skip_SourceNotMapped') + ',' +
                    (To-CsvField ("Folder '{0}' not mapped to a document title." -f $srcName))
            Add-Content -Path $LogPath -Value $line
            return
        }

        if (-not $student) {
            $line = (To-CsvField (Get-Date)) + ',' + (To-CsvField $srcFolder.FullName) + ',' +
                    (To-CsvField $f.Name) + ',' + (To-CsvField '') + ',' +
                    (To-CsvField $docTitle) + ',' + (To-CsvField '') + ',' +
                    (To-CsvField '') + ',' + (To-CsvField 'Skip_NoStudentSuffix') + ',' +
                    (To-CsvField "Filename lacks a final '_' student suffix")
            Add-Content -Path $LogPath -Value $line
            return
        }

        $safeStudentFolder = Safe-FolderName $student
        if (-not $safeStudentFolder) {
            $line = (To-CsvField (Get-Date)) + ',' + (To-CsvField $srcFolder.FullName) + ',' +
                    (To-CsvField $f.Name) + ',' + (To-CsvField $student) + ',' +
                    (To-CsvField $docTitle) + ',' + (To-CsvField '') + ',' +
                    (To-CsvField '') + ',' + (To-CsvField 'Skip_InvalidFolderName') + ',' +
                    (To-CsvField "Sanitized student folder name became empty")
            Add-Content -Path $LogPath -Value $line
            return
        }

        $newName = ("{0}_{1}{2}" -f $docTitle, $student, $ext)
        $destFolder = Join-Path $OrganizedRoot $safeStudentFolder
        $destPath   = Join-Path $destFolder $newName
        $destPath   = Get-UniquePath $destPath $AllowOverwrite

        if ($DryRun) {
            $line = (To-CsvField (Get-Date)) + ',' + (To-CsvField $srcFolder.FullName) + ',' +
                    (To-CsvField $f.Name) + ',' + (To-CsvField $student) + ',' +
                    (To-CsvField $docTitle) + ',' + (To-CsvField $newName) + ',' +
                    (To-CsvField $destPath) + ',' + (To-CsvField 'Plan') + ',' +
                    (To-CsvField '')
            Add-Content -Path $LogPath -Value $line
        }
        else {
            if (-not (Test-Path -LiteralPath $destFolder)) {
                New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
            }
            Move-Item -LiteralPath $f.FullName -Destination $destPath -Force
            $line = (To-CsvField (Get-Date)) + ',' + (To-CsvField $srcFolder.FullName) + ',' +
                    (To-CsvField $f.Name) + ',' + (To-CsvField $student) + ',' +
                    (To-CsvField $docTitle) + ',' + (To-CsvField (Split-Path -Leaf $destPath)) + ',' +
                    (To-CsvField $destPath) + ',' + (To-CsvField 'Moved') + ',' +
                    (To-CsvField '')
            Add-Content -Path $LogPath -Value $line
        }

        # Track presence
        if (-not $StudentDocs.ContainsKey($student)) { $StudentDocs[$student] = @{} }
        $StudentDocs[$student][$srcName] = $true
    }
}

Write-Host "Log written to: $LogPath"
Write-Host ("DryRun = {0}; OrganizedRoot = {1}" -f $DryRun, $OrganizedRoot)

# --------- Build "Missing documents" report (CSV text) ----------
'StudentName,Has_StudentApp,Has_FacultyApp,Has_Transcript,Missing,StudentFolderPath' | Out-File -FilePath $MissingReportPath -Encoding UTF8 -Force

$allStudents = @()
$allStudents = $StudentDocs.Keys | Sort-Object

foreach ($s in $allStudents) {
    $got = $StudentDocs[$s]
    $hasStudentApp = $false
    $hasFacultyApp = $false
    $hasTranscript = $false
    if ($got.ContainsKey('Student Application Forms')) { $hasStudentApp = $true }
    if ($got.ContainsKey('Faculty Application Forms')) { $hasFacultyApp = $true }
    if ($got.ContainsKey('Student Transcripts'))       { $hasTranscript = $true }

    $missing = @()
    if (-not $hasStudentApp) { $missing += 'Student Application' }
    if (-not $hasFacultyApp) { $missing += 'Faculty Application' }
    if (-not $hasTranscript) { $missing += 'Transcript' }
    $missingText = ($missing -join '; ')

    $studentFolder = Safe-FolderName $s
    $studentPath = if ($studentFolder) { Join-Path $OrganizedRoot $studentFolder } else { '' }

    $line = (To-CsvField $s) + ',' +
            (To-CsvField ($hasStudentApp.ToString())) + ',' +
            (To-CsvField ($hasFacultyApp.ToString())) + ',' +
            (To-CsvField ($hasTranscript.ToString())) + ',' +
            (To-CsvField $missingText) + ',' +
            (To-CsvField $studentPath)
    Add-Content -Path $MissingReportPath -Value $line
}

Write-Host "Missing-documents report: $MissingReportPath"



