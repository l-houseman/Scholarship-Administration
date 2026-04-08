# ========================= CONFIG ============================
$SourceDir = '\USRA 2026 Applications\Inputs\Information for Student Hiring Forms' #update with direct file path
$DestDir   = '\USRA 2026 Applications\USRA 2026_Hiring Forms'                      #update with direct file path

# Year label
$YearLabel = '2026'

# Safety: DryRun = $true prints actions only
$DryRun = $false

# Create destination folder if missing
if (-not (Test-Path -LiteralPath $DestDir)) {
    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
}

# Helper: extract the student name after the last "_"
function Get-StudentNameFromBase([string]$base) {
    if ([string]::IsNullOrWhiteSpace($base)) { return $null }
    $i = $base.LastIndexOf('_')
    if ($i -lt 0) { return $null }
    $s = $base.Substring($i + 1).Trim(' ','_','-')
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s
}

# ======================== PROCESS =============================
Get-ChildItem -LiteralPath $SourceDir -File |
ForEach-Object {
    $file = $_
    $base = $file.BaseName
    $ext  = $file.Extension

    $student = Get-StudentNameFromBase $base

    if ($null -eq $student) {
        Write-Warning "Skipped '$($file.Name)' because no final '_' student suffix was found."
        return
    }

    # New standardized name
    $newName = "{0} USRA Hiring Form_{1}{2}" -f $YearLabel, $student, $ext
    $target  = Join-Path $DestDir $newName

    # Ensure uniqueness
    $i = 1
    $candidate = $target
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $DestDir ("{0} ({1}){2}" -f ($newName -replace '\.xlsx$',''), $i, $ext)
        $i++
    }

    if ($DryRun) {
        Write-Host "[DRY RUN] Would rename+move:"
        Write-Host "    From: $($file.FullName)"
        Write-Host "    To:   $candidate"
    }
    else {
        Move-Item -LiteralPath $file.FullName -Destination $candidate -Force
        Write-Host "Moved:  $($file.Name) --> $(Split-Path -Leaf $candidate)"
    }
}