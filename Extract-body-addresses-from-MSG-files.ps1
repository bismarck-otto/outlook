# ChatGPT for bismarck-otto 2025-08-05 to Extract-body-addresses-from-MSG-files.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract Email Addresses from mail bodies in MSG files on harddrive
# =====================================================================================

# Define folder with .msg files
$msgFolder = "."  # <-- Change this to your folder path

# Create Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$emailAddresses = @()

# Regex pattern for email extraction
$emailRegex = '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b'

# Loop through all .msg files in the folder
Get-ChildItem -Path $msgFolder -Filter *.msg | ForEach-Object {
    $msgFile = $_.FullName
    try {
        # Load the message
        $item = $outlook.CreateItemFromTemplate($msgFile)

        # Only process if item has a Body
        if ($item.Body) {
            $emailMatches = Select-String -InputObject $item.Body -Pattern $emailRegex -AllMatches
            if ($emailMatches) {
                $emailAddresses += $emailMatches.Matches.Value
            }
        }
    } catch {
        Write-Warning "⚠️ Failed to process: $msgFile"
    }
}

# Clean and output
$emailAddresses = $emailAddresses | ForEach-Object { $_.ToLower() } | Sort-Object -Unique
$outputFile = "$msgFolder\EmailAddresses-MSG-$($emailAddresses.Count).txt"
$emailAddresses | Out-File -FilePath $outputFile -Encoding UTF8

Write-Host "`n✅ Extracted $($emailAddresses.Count) unique email addresses."
Write-Host "📄 Saved to: $outputFile"
