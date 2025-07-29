# ChatGPT for bismarck-otto 2025-07-29 to Extract-all-addresses-from-Outlook-folder.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract Email Addresses from a Specific Outlook Folder
# =====================================================================================

# Set mailbox name and folder path (use exact names from Outlook)
$mailBoxName = ""          # Or your mailbox / group mailbox name
$folderPath  = "Sent Items"     # Relative to the specified mailbox
$subfolders  = $false            # Set to $true to include all subfolders

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
        throw "Mailbox '$mailBoxName' not found. Check spelling and account access."
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
        }
    }
    return $currentFolder
}
function Get-EmailsFromFolder {
    param (
        $folder,
        [ref]$emailList
    )

    # Extract emails from this folder
    foreach ($item in $folder.Items) {
        if ($item -and $item.MessageClass -eq "IPM.Note") {
            $from = $item.SenderEmailAddress
            if ($from) { $emailList.Value += $from }

            foreach ($recip in $item.Recipients) {
                $address = $recip.Address
                if ($address) { $emailList.Value += $address }
            }
        }
    }

    # Recurse into subfolders
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
        if ($item -and $item.MessageClass -eq "IPM.Note") {
            $from = $item.SenderEmailAddress
            if ($from) { $emailAddresses += $from }

            foreach ($recip in $item.Recipients) {
                $address = $recip.Address
                if ($address) { $emailAddresses += $address }
            }
        }
    }
}

# Remove duplicates and sort
$emailAddresses = $emailAddresses | Sort-Object -Unique

# Output
$emailAddresses | Out-File -FilePath "EmailAddresses.txt" -Encoding UTF8
Write-Output "✅ Extracted $($emailAddresses.Count) unique email addresses to EmailAddresses.txt"