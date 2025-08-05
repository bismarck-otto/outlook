# ChatGPT for bismarck-otto 2025-08-05 to Extract-all-domains-and-create-info-list.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract all domains in EmailAddresses-Combined* and generate info@alldomains.com
# =====================================================================================

# Set folder path
$folderPath = "."

# Read all EmailAddresses-Combined*.txt files
$combinedFiles = Get-ChildItem -Path $folderPath -Filter "EmailAddresses-Combined*.txt" -File

# Collect all combined email addresses
$allCombinedEmails = @()
foreach ($file in $combinedFiles) {
    $content = Get-Content $file.FullName | ForEach-Object { $_.Trim() }
    $allCombinedEmails += $content
}
$allCombinedEmails = $allCombinedEmails | Where-Object { $_ -ne "" }

# Extract unique domains from combined email addresses
$combinedDomains = @()
foreach ($email in $allCombinedEmails) {
    if ($email -match '@(.+)$') {
        $combinedDomains += $matches[1].ToLower()
    }
}
$combinedDomains = $combinedDomains | Sort-Object -Unique

# Generate 'info@domain.com' addresses
$infoEmails = $combinedDomains | ForEach-Object { "info@$_" }

# Save to new file
$outputFile = Join-Path $folderPath "EmailAddresses-Combined-Info.txt"
$infoEmails | Set-Content -Encoding UTF8 $outputFile