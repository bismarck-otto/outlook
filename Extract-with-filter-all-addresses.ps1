# ChatGPT for bismarck-otto 2025-08-05 to Extract-with-filter-all-addresses.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Filters each EmailAddresses-*.txt file in the specified folder
# =====================================================================================

# Set folder path
$folderPath = "."

# Read all EmailAddresses-*.txt files
$combinedFiles = Get-ChildItem -Path $folderPath -Filter "EmailAddresses-*.txt" -File

# Define exclusion filter as a newline-separated string
$filterText = @"
news
newsletter
no-reply
noreply
notification
Office365Reports
booking
mailing
mailer-daemon
linkedin
facebook
"@

# Convert to array and trim empty entries
$filter = $filterText -split "`r?`n" | Where-Object { $_.Trim() -ne "" }

# Process each file
foreach ($file in $combinedFiles) {
    $lines = Get-Content $file.FullName

    $filtered = $lines | Where-Object {
        $line = $_.ToLower()
        -not ($filter | Where-Object { $line -like "*$_*" })
    }

    # Overwrite the original file with filtered content
    Set-Content -Path $file.FullName -Value $filtered
}
