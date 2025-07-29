# ChatGPT for bismarck-otto 2025-07-29 to Extract-all-addresses-from-Outlook-folder.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract Email Addresses from a Specific Outlook Folder
# =====================================================================================

# Set mailbox name and folder path (use exact names from Outlook)
$mailBoxName = ""               # Or your mailbox / group mailbox name
$folderPath  = "Sent Items"     # Relative to the specified mailbox
$subfolders  = $false           # Set to $true to include all subfolders

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
