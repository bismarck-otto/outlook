# ChatGPT for bismarck-otto 2025-08-05 to Extract-and-add-info@.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract Email Addresses from all files EmailAddresses-xxxx-NN.txt
# =====================================================================================
# Set the directory (optional: "." means current directory)
$folderPath = "."

# Get all files matching the pattern
$files = Get-ChildItem -Path $folderPath -Filter "EmailAddresses*.txt"

foreach ($file in $files) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $extension = $file.Extension
   
    $processedEmails = Get-Content $file.FullName | ForEach-Object {
        if ($_ -match "@") {
            $_ -replace '^[^@]+', 'info'
        } else {
            $_
        }
    }

    # Sort and remove duplicates
    $uniqueEmails = $processedEmails | Sort-Object -Unique

    # Output to a file
    $outputFile = "$baseName-info-$($uniqueEmails.Count)$extension"
    $uniqueEmails | Set-Content $outputFile

    # Write-Host "Processed $($file.Name) -> $outputFile"
    Write-Host "✅ Processed $($uniqueEmails.Count) unique email addresses to $outputFile"
}


    