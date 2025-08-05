# ChatGPT for bismarck-otto 2025-08-05 to Extract-and-split-Combined-vs-MSG.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract and split Combined Email Addresses in Combined-remaining vs MSG
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

# Read all EmailAddresses-MSG*.txt files
$msgFiles = Get-ChildItem -Path $folderPath -Filter "EmailAddresses-MSG*.txt" -File

# Collect all MSG email addresses
$allMsgEmails = @()
foreach ($file in $msgFiles) {
    $content = Get-Content $file.FullName | ForEach-Object { $_.Trim() }
    $allMsgEmails += $content
}
$allMsgEmails = $allMsgEmails | Where-Object { $_ -ne "" }

# Convert to hashset for faster lookup
$hashMsgEmails = @{}
foreach ($email in $allMsgEmails) {
    $hashMsgEmails[$email.ToLower()] = $true
}

# Split into two sets: those in MSG and not in MSG
$noMSGaddresses = @()
$isMSGaddresses = @()

foreach ($email in $allCombinedEmails) {
    $lowerEmail = $email.ToLower()
    if ($hashMsgEmails.ContainsKey($lowerEmail)) {
        $isMSGaddresses += $email
    } else {
        $noMSGaddresses += $email
    }
}

# Remove duplicates and sort
$uniqueNoMSGs = $noMSGaddresses | Sort-Object -Unique
$uniqueIsMSGs = $isMSGaddresses | Sort-Object -Unique

# Write output files
$noMSGaddressesOut = Join-Path $folderPath "EmailAddresses-Combined-remaining-$($uniqueNoMSGs.Count).txt"
$uniqueNoMSGs | Set-Content -Encoding UTF8 $noMSGaddressesOut

$isMSGaddressesOut = Join-Path $folderPath "EmailAddresses-Combined-MSG-$($uniqueIsMSGs.Count).txt"
$uniqueIsMSGs | Set-Content -Encoding UTF8 $isMSGaddressesOut

# Optional: print summary
Write-Output "✅ noMSGs written: $($uniqueNoMSGs.Count)"
Write-Output "✅ MSGs written: $($uniqueIsMSGs.Count)"
