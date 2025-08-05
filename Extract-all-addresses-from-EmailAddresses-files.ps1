# ChatGPT for bismarck-otto 2025-08-05 to Extract-all-addresses-from-EmailAddresses-files.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract Email Addresses from all files EmailAddresses-xxxx-NN.txt
# =====================================================================================
# Set the directory (optional: "." means current directory)
$folderPath = "."

# Get all files beginning with 'EmailAddresses-'
$emailFiles = Get-ChildItem -Path $folderPath -Filter "EmailAddresses-*" -File

# Read all lines, sort them, remove duplicates
$allLines = foreach ($file in $emailFiles) {
    Get-Content $file.FullName
}

# Sort and remove duplicates
$uniqueSortedLines = $allLines | Sort-Object -Unique

# Output to a file (optional)
$outputFile = Join-Path $folderPath "EmailAddresses-Combined-$($uniqueSortedLines.Count).txt"
$uniqueSortedLines | Set-Content -Encoding UTF8 $outputFile

# Optional: also print to console
# $uniqueSortedLines

Write-Host "✅ Extracted $($uniqueSortedLines.Count) unique email addresses to $outputFile"
