# ChatGPT for bismarck-otto 2025-07-29 to List-all-Outlook-folders.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# List All Outlook Folders (including shared mailboxes)
# =====================================================================

# Start Outlook COM object
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

# Recursive function to walk folders
function Get-FolderPaths {
    param (
        $ParentFolder,
        [string]$CurrentPath = ""
    )

    $folderPath = if ($CurrentPath) { "$CurrentPath\$($ParentFolder.Name)" } else { $ParentFolder.Name }
    Write-Output $folderPath

    foreach ($subFolder in $ParentFolder.Folders) {
        Get-FolderPaths -ParentFolder $subFolder -CurrentPath $folderPath
    }
}

# Loop through all top-level mailboxes
for ($i = 1; $i -le $namespace.Folders.Count; $i++) {
    $mailbox = $namespace.Folders.Item($i)
    Write-Host "`n📂 Mailbox: $($mailbox.Name)"
    Get-FolderPaths -ParentFolder $mailbox
}

$output = @()

for ($i = 1; $i -le $namespace.Folders.Count; $i++) {
    $mailbox = $namespace.Folders.Item($i)
    $output += "`n📂 Mailbox: $($mailbox.Name)"
    $output += Get-FolderPaths -ParentFolder $mailbox
}

$output | Out-File -FilePath "OutlookFolders.txt" -Encoding UTF8
Write-Host "✅ Folder list saved to OutlookFolders.txt"