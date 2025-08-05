# ChatGPT for bismarck-otto 2025-08-05 to Extract-body-addresses-from-Outlook-folder.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract Email Addresses from mail bodies in a Specific Outlook Folder
# =====================================================================================

# Set mailbox name and folder path (use exact names from Outlook)
$mailBoxName = ""          # Or your mailbox / group mailbox name
$folderPath  = "Inbox"     # Relative to the specified mailbox
$subfolders  = $false      # Set to $true to include all subfolders

# Define exclusion filter as a newline-separated string
$filterText = @"
news
newsletter
no-reply
noreply
notification
Office365Reports
microsoftexchange
prod.outlook.com
booking
mailing
mailer-daemon
linkedin
facebook
"@ # keep this end quote

# Convert to array and trim empty entries
$filter = $filterText -split "`r?`n" | Where-Object { $_.Trim() -ne "" }

# Sanitize Folder Path for File Name
function Get-SafeFileNameFromFolderPath {
    param ([string]$path)

    # Replace invalid characters with underscore
    $safeName = $path -replace '[\\/:*?"<>|]', '-'

    # Optionally limit length (e.g. max 100 chars)
    if ($safeName.Length -gt 100) {
        $safeName = $safeName.Substring(0, 100)
    }

    return $safeName
}

# Launch Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $null
$maxAttempts = 10
$attempt = 0

while (-not $namespace -and $attempt -lt $maxAttempts) {
    try {
        $namespace = $outlook.GetNamespace("MAPI")
    } catch {
        Write-Host "Outlook not ready, retrying in 1 second..."
        Start-Sleep -Seconds 1
        $attempt++
    }
}

if (-not $namespace) {
    throw "Outlook is not responding after multiple attempts."
}

# If empty, use default mailbox
if ([string]::IsNullOrWhiteSpace($mailBoxName)) {
    $rootMailbox = $namespace.Folders.Item(1)
    Write-Host "Using default mailbox: $($rootMailbox.Name)"
} else {
    $rootMailbox = $null
    for ($i = 1; $i -le $namespace.Folders.Count; $i++) {
        $mailbox = $namespace.Folders.Item($i)
        if ($mailbox.Name -eq $mailBoxName) {
            $rootMailbox = $mailbox
            break
        }
    }

    if (-not $rootMailbox) {
        throw "Mailbox not found. Check spelling and account access."
    } else {
        Write-Host "Using specified mailbox: $($rootMailbox.Name)"
    }
}

# Navigate to folder path within selected mailbox
function Get-OutlookFolder {
    param (
        $root,
        [string]$path
    )

    $folders = $path -split '\\'
    $currentFolder = $root
    foreach ($f in $folders) {
        $currentFolder = $currentFolder.Folders.Item($f)
        if (-not $currentFolder) {
            throw "Folder not found: $f"
        } else {
            Write-Host "Using specified folder: $f"
        }
    }
    return $currentFolder
}

# Recursively extract email addresses from body text
function Get-EmailsFromFolder {
    param (
        $folder,
        [ref]$emailList
    )

    foreach ($item in $folder.Items) {
        if ($item -and ($item.MessageClass -eq "IPM.Note" -or $item.MessageClass -like "IPM.Report*")) {
            $body = $item.Body
            if ($body) {
                $emailMatches = Select-String -InputObject $body -Pattern '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b' -AllMatches
                if ($emailMatches) {
                    $emailList.Value += $emailMatches.Matches.Value
                }
            }
        }
    }

    foreach ($subFolder in $folder.Folders) {
        Get-EmailsFromFolder -folder $subFolder -emailList $emailList
    }
}

$targetFolder = Get-OutlookFolder -root $rootMailbox -path $folderPath
$emailAddresses = @()

if ($subfolders) {
    Get-EmailsFromFolder -folder $targetFolder -emailList ([ref]$emailAddresses)
} else {
    foreach ($item in $targetFolder.Items) {
        if ($item -and ($item.MessageClass -eq "IPM.Note" -or $item.MessageClass -like "IPM.Report*")) {
            $body = $item.Body
            if ($body) {
                $emailMatches = Select-String -InputObject $body -Pattern '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b' -AllMatches
                if ($emailMatches) {
                    $emailAddresses += $emailMatches.Matches.Value
                }
            }
        }
    }
}

# Normalize: lowercase all addresses
$emailAddresses = $emailAddresses | ForEach-Object { $_.ToLower() }

# Exclude Exchange internal addresses
$emailAddresses = $emailAddresses | Where-Object { $_ -notlike "*ou=exchange*" } 

# Apply substring filters (case-insensitive)
$emailAddresses = $emailAddresses | Where-Object {
    $addr = $_.ToLower()
    -not ($filter | Where-Object { $addr -like "*$_*" })
}

# Remove duplicates and sort
$emailAddresses = $emailAddresses | Sort-Object -Unique

# Output to file
$safeFileName = Get-SafeFileNameFromFolderPath $folderPath
$outputFile = "EmailAddresses-$safeFileName-body-$($emailAddresses.Count).txt"
$emailAddresses | Out-File -FilePath $outputFile -Encoding UTF8
Write-Host "✅ Extracted $($emailAddresses.Count) unique email addresses to $outputFile"
