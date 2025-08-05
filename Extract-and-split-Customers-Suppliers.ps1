# ChatGPT for bismarck-otto 2025-08-05 to Extract-and-split-Customers-Suppliers.ps1

# Copyright (c) 2025 Otto von Bismarck
# This project includes portions generated using OpenAI’s ChatGPT.
# All code is released under the MIT License.

# Extract and split Email Addresses for customers and suppliers
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

# Read all EmailAddresses-Suppliers*.txt files
$supplierFiles = Get-ChildItem -Path $folderPath -Filter "EmailAddresses-Suppliers*.txt" -File

# Collect all supplier email addresses
$allSupplierEmails = @()
foreach ($file in $supplierFiles) {
    $content = Get-Content $file.FullName | ForEach-Object { $_.Trim() }
    $allSupplierEmails += $content
}
$allSupplierEmails = $allSupplierEmails | Where-Object { $_ -ne "" }

# Extract unique domains from supplier emails
$supplierDomains = @()
foreach ($email in $allSupplierEmails) {
    if ($email -match '@(.+)$') {
        $supplierDomains += $matches[1].ToLower()
    }
}
$supplierDomains = $supplierDomains | Sort-Object -Unique

# Convert domain list to hash set for fast lookup
$domainSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($domain in $supplierDomains) {
    [void]$domainSet.Add($domain)
}

# Separate combined emails into customers and suppliers
$customers = @()
$suppliers = @()

foreach ($email in $allCombinedEmails) {
    if ($email -match '@(.+)$') {
        $domain = $matches[1].ToLower()
        if ($domainSet.Contains($domain)) {
            $suppliers += $email
        } else {
            $customers += $email
        }
    }
}

# Remove duplicates and sort
$uniqueCustomers = $customers | Sort-Object -Unique
$uniqueSuppliers = $suppliers | Sort-Object -Unique

# Write output files
$customersOut = Join-Path $folderPath "EmailAddresses-Combined-Customers-$($uniqueCustomers.Count).txt"
$uniqueCustomers | Set-Content -Encoding UTF8 $customersOut

$suppliersOut = Join-Path $folderPath "EmailAddresses-Combined-Suppliers-$($uniqueSuppliers.Count).txt"
$uniqueSuppliers | Set-Content -Encoding UTF8 $suppliersOut

# Optional: print summary
Write-Output "✅ Customers written: $($uniqueCustomers.Count)"
Write-Output "✅ Suppliers written: $($uniqueSuppliers.Count)"
