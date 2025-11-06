<#
.SYNOPSIS
PowerShell script to copy CSV files from specified folders while skipping folders with 'nhp' in their names.

.DESCRIPTION
This script searches for CSV files in drishti_backup folders within the specified main folders.
It skips any folders containing 'nhp' in their names and copies files modified within the last 30 days.

.PARAMETER MainFolders
Array of folder paths to search for CSV files. Default: @("DWLR_118x2", "DWLR", "DWLR_118x3")

.PARAMETER OutputFolder
Destination folder for copied CSV files. Default: "gwfiles"

.EXAMPLE
.\copy_csv_files_skip_nhp.ps1 -MainFolders @("C:\MyFolder1", "C:\MyFolder2") -OutputFolder "C:\Output"

.EXAMPLE
.\copy_csv_files_skip_nhp.ps1
#>
param(
    [Parameter(Mandatory=$false)]
    [string[]]$MainFolders = @("DWLR_118x2", "DWLR", "DWLR_118x3"),
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFolder = "gwfiles"
)

if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

$duplicatesFolder = Join-Path $outputFolder "duplicates"
if (-not (Test-Path $duplicatesFolder)) {
    New-Item -ItemType Directory -Path $duplicatesFolder -Force | Out-Null
    Write-Host "Duplicates folder created: $duplicatesFolder" -ForegroundColor Cyan
}

$logDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$logFileName = "copy_csv_skip_nhp_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
$logFile = Join-Path $logDir $logFileName

function Write-Log {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    Write-Host $Message -ForegroundColor $ForegroundColor
    
    Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
}

$currentDate = Get-Date
$tenDaysAgo = $currentDate.AddDays(-30).Date

Write-Log ("=" * 70) "Gray"
Write-Log "CSV Copy Script Started" "Cyan"
Write-Log "Log File: $logFile" "Cyan"
Write-Log ("=" * 70) "Gray"

Write-Log "Current Date: $($currentDate.ToString('yyyy-MM-dd HH:mm:ss'))" "Cyan"
Write-Log "Looking for files from: $($tenDaysAgo.ToString('yyyy-MM-dd HH:mm:ss'))" "Cyan"
Write-Log ("=" * 70) "Gray"

$totalFilesCopied = 0
$totalFilesFound = 0
$totalSkippedFolders = 0
$processedFiles = 0
$folderStats = @{}

Write-Log "Loading existing files in memory..." "Cyan"
$existingFiles = [System.Collections.Generic.HashSet[string]]::new()
if (Test-Path $outputFolder) {
    $existingFilesList = Get-ChildItem -Path $outputFolder -Filter "*.csv" -File -ErrorAction SilentlyContinue
    foreach ($file in $existingFilesList) {
        [void]$existingFiles.Add($file.Name.ToLower())
    }
    Write-Log "  Found $($existingFiles.Count) existing CSV file(s) in gwfiles" "Gray"
}

foreach ($mainFolder in $mainFolders) {
    $folderStats[$mainFolder] = @{
        FilesFound = 0
        FilesCopied = 0
        UniqueFilesToGwfiles = 0
        FoldersSkipped = 0
    }
    if (-not (Test-Path $mainFolder)) {
        Write-Log "Folder not found: $mainFolder" "Yellow"
        continue
    }
    
    Write-Log "`nSearching in: $mainFolder" "Green"
    
    $drishtiBackupFolders = Get-ChildItem -Path $mainFolder -Recurse -Directory -Filter "drishti_backup*" -ErrorAction SilentlyContinue |                                        
        Where-Object { 
            if ($_.Name -eq "drishti_backup_nhp" -or $_.Name -like "*_nhp*") {
                return $false
            }
            if ($_.FullName -like "*drishti_backup_nhp*" -or $_.FullName -like "*\_nhp\*" -or $_.FullName -like "*/_nhp/*") {
                return $false
            }
            if ($_.Name -eq "drishti_backup") {
                return $true
            }
            return $false
        }
    
    $allDrishtiFolders = Get-ChildItem -Path $mainFolder -Recurse -Directory -Filter "drishti_backup*" -ErrorAction SilentlyContinue
    $skippedCount = ($allDrishtiFolders | Where-Object { $_.Name -eq "drishti_backup_nhp" -or $_.Name -like "*_nhp*" }).Count
    if ($skippedCount -gt 0) {
        $folderStats[$mainFolder].FoldersSkipped = $skippedCount
        $totalSkippedFolders += $skippedCount
    }
    
    foreach ($folder in $drishtiBackupFolders) {
        if ($folder.Name -like "*_nhp*" -or 
            $folder.Name -eq "drishti_backup_nhp" -or 
            $folder.FullName -like "*drishti_backup_nhp*" -or
            $folder.FullName -like "*\_nhp\*" -or
            $folder.FullName -like "*/_nhp/*") {    
            continue
        }
        
        $csvFiles = Get-ChildItem -Path $folder.FullName -Filter "*.csv" -File -ErrorAction SilentlyContinue
        
        if ($csvFiles.Count -eq 0) {
            continue
        }
        
        $filesInFolder = 0
        $filesCopiedInFolder = 0
        
        foreach ($csvFile in $csvFiles) {
            $fileMTime = $csvFile.LastWriteTime
            
            if ($fileMTime -ge $tenDaysAgo) {
                $folderStats[$mainFolder].FilesFound++
                $totalFilesFound++
                $filesInFolder++
                $processedFiles++
                
                if ($processedFiles % 1000 -eq 0) {
                    Write-Host "  Progress: Processed $processedFiles files, Copied: $totalFilesCopied" -ForegroundColor Cyan
                }
                
                $fileNameLower = $csvFile.Name.ToLower()
                $isDuplicate = $existingFiles.Contains($fileNameLower)
                
                if ($isDuplicate) {
                    $duplicateDestination = Join-Path $duplicatesFolder $csvFile.Name
                    
                    try {
                        Copy-Item -Path $csvFile.FullName -Destination $duplicateDestination -Force -ErrorAction Stop
                        $folderStats[$mainFolder].FilesCopied++
                        $totalFilesCopied++
                        $filesCopiedInFolder++
                    }
                    catch {
                        Write-Log "    ERROR copying duplicate: $($csvFile.Name) - $($_.Exception.Message)" "Red"
                    }
                }
                else {
                    $destination = Join-Path $outputFolder $csvFile.Name
                    
                    try {
                        Copy-Item -Path $csvFile.FullName -Destination $destination -ErrorAction Stop
                        [void]$existingFiles.Add($fileNameLower)
                        $folderStats[$mainFolder].FilesCopied++
                        $folderStats[$mainFolder].UniqueFilesToGwfiles++
                        $totalFilesCopied++
                        $filesCopiedInFolder++
                    }
                    catch {
                        Write-Log "    ERROR copying: $($csvFile.Name) - $($_.Exception.Message)" "Red"
                    }
                }
            }
        }
        
        if ($filesInFolder -gt 0) {
            Write-Log "  drishti_backup: $($folder.FullName) - Found: $filesInFolder, Copied: $filesCopiedInFolder" "Gray"
        }
    }
}

Write-Log "`n$("=" * 70)" "Gray"
Write-Log "SUMMARY" "Cyan"
Write-Log ("=" * 70) "Gray"

Write-Log "`nFolder-wise Statistics (Last 30 Days):" "Yellow"
foreach ($mainFolder in $mainFolders) {
    $stats = $folderStats[$mainFolder]
    Write-Log "  $mainFolder :" "White"
    Write-Log "    Files Found: $($stats.FilesFound)" "Green"
    Write-Log "    Files Copied: $($stats.FilesCopied)" "Cyan"
    Write-Log "    Unique Files to gwfiles: $($stats.UniqueFilesToGwfiles)" "Magenta"
    if ($stats.FoldersSkipped -gt 0) {
        Write-Log "    Folders Skipped (drishti_backup_nhp): $($stats.FoldersSkipped)" "DarkYellow"
    }
}

Write-Log "`n$("=" * 70)" "Gray"
Write-Log "Total Summary:" "Yellow"

$totalUniqueToGwfiles = 0
foreach ($mainFolder in $mainFolders) {
    $totalUniqueToGwfiles += $folderStats[$mainFolder].UniqueFilesToGwfiles
}

Write-Log "  Total Files Found (Last 30 Days): $totalFilesFound" "Green"
Write-Log "  Total Files Copied: $totalFilesCopied" "Cyan"
Write-Log "  Total Unique Files to gwfiles: $totalUniqueToGwfiles" "Magenta"
Write-Log "  Total Folders Skipped (drishti_backup_nhp): $totalSkippedFolders" "DarkYellow"
Write-Log "  Output folder: $(Resolve-Path $outputFolder)" "Gray"
Write-Log "  Duplicates folder: $(Resolve-Path $duplicatesFolder)" "Gray"
Write-Log ("=" * 70) "Gray"

Write-Log "`n$("=" * 70)" "Green"
Write-Log "Process completed successfully!" "Green"
Write-Log "Log file saved to: $logFile" "Cyan"
Write-Log ("=" * 70) "Green"
