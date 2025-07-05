#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Converts Skype CSV export to VCF (vCard) format
.DESCRIPTION
    This script reads a CSV file exported from Skype and converts each contact to VCF format.
    Works on both Windows and Linux with PowerShell Core.
.EXAMPLE
    .\Convert-SkypeCSVToVCF.ps1
#>

# Function to show file picker dialog (cross-platform)
function Get-FilePath {
    param(
        [string]$Title,
        [string]$Filter = "All files (*.*)|*.*",
        [string]$InitialDirectory = $PWD,
        [switch]$Save
    )

    # Try to use GUI file picker if available
    $useGui = $false

    # Check if we're on Windows and can use Windows Forms
    if ($PSVersionTable.PSVersion.Major -ge 5 -or $IsWindows) {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $useGui = $true
        } catch {
            Write-Host "GUI file picker not available, using console input." -ForegroundColor Yellow
        }
    }

    if ($useGui) {
        try {
            if ($Save) {
                $dialog = New-Object System.Windows.Forms.SaveFileDialog
            } else {
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
            }
            $dialog.Title = $Title
            $dialog.Filter = $Filter
            $dialog.InitialDirectory = $InitialDirectory

            # Set properties that are available
            try {
                $dialog.RestoreDirectory = $true
            } catch {
                # Property not available in this version
            }

            # Show the dialog
            $result = $dialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                return $dialog.FileName
            } else {
                Write-Host "File selection cancelled." -ForegroundColor Yellow
                return $null
            }
        } catch {
            Write-Host "Error with file dialog: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Falling back to console input." -ForegroundColor Yellow
        }
    }
    
    # Console fallback for Linux or if GUI fails
    do {
        $path = Read-Host $Title
        if ([string]::IsNullOrWhiteSpace($path)) {
            Write-Host "Please enter a valid path." -ForegroundColor Yellow
            continue
        }
        
        # Expand relative paths
        $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
        
        if (-not $Save) {
            if (Test-Path $path -PathType Leaf) {
                return $path
            } else {
                Write-Host "File not found: $path" -ForegroundColor Red
            }
        } else {
            # For save dialog, just check if directory exists
            $directory = Split-Path $path -Parent
            if (Test-Path $directory -PathType Container) {
                return $path
            } else {
                Write-Host "Directory not found: $directory" -ForegroundColor Red
            }
        }
    } while ($true)
}

# Function to create VCF content from contact data
function New-VCardContent {
    param(
        [PSCustomObject]$Contact
    )
    
    $vcf = @()
    $vcf += "BEGIN:VCARD"
    $vcf += "VERSION:3.0"
    
    # Full name (display_name or constructed from first/last name)
    $fullName = ""
    if (-not [string]::IsNullOrWhiteSpace($Contact.display_name)) {
        $fullName = $Contact.display_name.Trim()
    } elseif (-not [string]::IsNullOrWhiteSpace($Contact.'profile.name.first') -or 
              -not [string]::IsNullOrWhiteSpace($Contact.'profile.name.surname')) {
        $firstName = if ($Contact.'profile.name.first') { $Contact.'profile.name.first'.Trim() } else { "" }
        $lastName = if ($Contact.'profile.name.surname') { $Contact.'profile.name.surname'.Trim() } else { "" }
        $fullName = "$firstName $lastName".Trim()
    }
    
    if ($fullName) {
        $vcf += "FN:$fullName"
    }
    
    # Structured name (N property)
    $firstName = if ($Contact.'profile.name.first') { $Contact.'profile.name.first'.Trim() } else { "" }
    $lastName = if ($Contact.'profile.name.surname') { $Contact.'profile.name.surname'.Trim() } else { "" }
    if ($firstName -or $lastName) {
        $vcf += "N:$lastName;$firstName;;;"
    }
    
    # Skype handle as instant messaging
    if (-not [string]::IsNullOrWhiteSpace($Contact.'profile.skype_handle')) {
        $skypeHandle = $Contact.'profile.skype_handle'.Trim()
        $vcf += "X-SKYPE:$skypeHandle"
        $vcf += "IMPP:skype:$skypeHandle"
    }
    
    # Website/URL
    if (-not [string]::IsNullOrWhiteSpace($Contact.'profile.website')) {
        $website = $Contact.'profile.website'.Trim()
        $vcf += "URL:$website"
    }
    
    # About/Note
    if (-not [string]::IsNullOrWhiteSpace($Contact.'profile.about')) {
        $about = $Contact.'profile.about'.Trim()
        # Escape special characters in notes
        $about = $about -replace '\\', '\\\\' -replace ',', '\,' -replace ';', '\;' -replace '\n', '\n' -replace '\r', ''
        $vcf += "NOTE:$about"
    }
    
    # Avatar URL as photo
    if (-not [string]::IsNullOrWhiteSpace($Contact.'profile.avatar_url')) {
        $avatarUrl = $Contact.'profile.avatar_url'.Trim()
        $vcf += "PHOTO;VALUE=URI:$avatarUrl"
    }
    
    # Phone numbers - handle Skype's specific phone field patterns and common formats
    $phoneNumbers = @()

    # Check for Skype's indexed phone fields (phones[0].number, phones[1].number, etc.)
    for ($i = 0; $i -lt 10; $i++) {
        $numberField = "phones[$i].number"
        $typeField = "phones[$i].type"

        if ($Contact.PSObject.Properties.Name -contains $numberField) {
            $phoneValue = $Contact.$numberField
            $phoneTypeValue = $Contact.$typeField

            if (-not [string]::IsNullOrWhiteSpace($phoneValue)) {
                $phoneNumbers += @{
                    Number = $phoneValue.Trim()
                    Type = if ($phoneTypeValue) { $phoneTypeValue.Trim() } else { "voice" }
                }
            }
        }
    }

    # Check for Skype's profile phone fields (profile.phones[0].number, etc.)
    for ($i = 0; $i -lt 10; $i++) {
        $numberField = "profile.phones[$i].number"
        $typeField = "profile.phones[$i].type"

        if ($Contact.PSObject.Properties.Name -contains $numberField) {
            $phoneValue = $Contact.$numberField
            $phoneTypeValue = $Contact.$typeField

            if (-not [string]::IsNullOrWhiteSpace($phoneValue)) {
                $phoneNumbers += @{
                    Number = $phoneValue.Trim()
                    Type = if ($phoneTypeValue) { $phoneTypeValue.Trim() } else { "voice" }
                }
            }
        }
    }

    # Check for other common phone number field names
    $commonPhoneFields = @(
        'phone', 'phone_number', 'mobile', 'mobile_phone', 'cell', 'cell_phone',
        'home_phone', 'work_phone', 'business_phone', 'telephone', 'tel',
        'profile.phone', 'profile.mobile', 'profile.phone_number',
        'contact.phone', 'contact.mobile', 'contact.phone_number'
    )

    foreach ($field in $commonPhoneFields) {
        if ($Contact.PSObject.Properties.Name -contains $field) {
            $phoneValue = $Contact.$field
            if (-not [string]::IsNullOrWhiteSpace($phoneValue)) {
                $phoneType = "voice"
                if ($field -match "mobile|cell") {
                    $phoneType = "mobile"
                } elseif ($field -match "home") {
                    $phoneType = "home"
                } elseif ($field -match "work|business") {
                    $phoneType = "work"
                }

                $phoneNumbers += @{
                    Number = $phoneValue.Trim()
                    Type = $phoneType
                }
            }
        }
    }

    # Remove duplicates and add to VCF
    $uniquePhones = $phoneNumbers | Sort-Object Number -Unique
    foreach ($phone in $uniquePhones) {
        $vcfPhoneType = switch ($phone.Type.ToLower()) {
            "mobile" { "CELL" }
            "cell" { "CELL" }
            "home" { "HOME" }
            "work" { "WORK" }
            "business" { "WORK" }
            default { "VOICE" }
        }
        $vcf += "TEL;TYPE=$vcfPhoneType`:$($phone.Number)"
    }

    # Email addresses - check for common email field names
    $emailFields = @(
        'email', 'email_address', 'mail', 'e_mail',
        'home_email', 'work_email', 'business_email',
        'profile.email', 'profile.email_address',
        'contact.email', 'contact.email_address'
    )

    foreach ($field in $emailFields) {
        if ($Contact.PSObject.Properties.Name -contains $field) {
            $emailValue = $Contact.$field
            if (-not [string]::IsNullOrWhiteSpace($emailValue)) {
                $emailValue = $emailValue.Trim()
                # Determine email type based on field name
                $emailType = "INTERNET"
                if ($field -match "home") {
                    $emailType = "HOME"
                } elseif ($field -match "work|business") {
                    $emailType = "WORK"
                }
                $vcf += "EMAIL;TYPE=$emailType`:$emailValue"
            }
        }
    }

    # Country/Location
    if (-not [string]::IsNullOrWhiteSpace($Contact.'profile.locations[0].country')) {
        $country = $Contact.'profile.locations[0].country'.Trim().ToUpper()
        $vcf += "ADR:;;;;;;$country"
    }

    # Creation time as revision
    if (-not [string]::IsNullOrWhiteSpace($Contact.creation_time)) {
        try {
            $creationDate = [DateTime]::Parse($Contact.creation_time)
            $revDate = $creationDate.ToString("yyyyMMddTHHmmssZ")
            $vcf += "REV:$revDate"
        } catch {
            # Ignore invalid dates
        }
    }
    
    $vcf += "END:VCARD"
    
    return $vcf -join "`r`n"
}

# Function to sanitize filename
function Get-SafeFileName {
    param([string]$Name)
    
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "Unknown_Contact"
    }
    
    # Remove or replace invalid filename characters
    $safeName = $Name -replace '[<>:"/\\|?*]', '_'
    $safeName = $safeName -replace '\s+', '_'
    $safeName = $safeName.Trim('._')
    
    # Limit length
    if ($safeName.Length -gt 50) {
        $safeName = $safeName.Substring(0, 50)
    }
    
    return $safeName
}

# Main script
try {
    Write-Host "Skype CSV to VCF Converter" -ForegroundColor Green
    Write-Host "=========================" -ForegroundColor Green
    Write-Host ""
    
    # Get input CSV file
    Write-Host "Step 1: Select the Skype CSV export file" -ForegroundColor Cyan
    $csvPath = Get-FilePath -Title "Select Skype CSV file" -Filter "CSV files (*.csv)|*.csv|All files (*.*)|*.*"

    if (-not $csvPath) {
        Write-Host "No file selected. Exiting." -ForegroundColor Red
        Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
    
    Write-Host "Selected CSV file: $csvPath" -ForegroundColor Green
    
    # Read and parse CSV
    Write-Host "`nReading CSV file..." -ForegroundColor Yellow
    try {
        $contacts = Import-Csv -Path $csvPath
        Write-Host "Found $($contacts.Count) contacts in CSV file." -ForegroundColor Green

        # Show available fields
        if ($contacts.Count -gt 0) {
            $fields = $contacts[0].PSObject.Properties.Name
            Write-Host "`nAvailable fields in CSV:" -ForegroundColor Cyan
            $fields | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

            # Check for phone and email fields
            $phoneFields = $fields | Where-Object { $_ -match "phone|mobile|cell|tel" }
            $emailFields = $fields | Where-Object { $_ -match "email|mail" }

            if ($phoneFields) {
                Write-Host "`nPhone fields detected: $($phoneFields -join ', ')" -ForegroundColor Green
            }
            if ($emailFields) {
                Write-Host "Email fields detected: $($emailFields -join ', ')" -ForegroundColor Green
            }
            if (-not $phoneFields -and -not $emailFields) {
                Write-Host "`nNo phone or email fields detected in this CSV." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "Error reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }

    if ($contacts.Count -eq 0) {
        Write-Host "No contacts found in CSV file." -ForegroundColor Red
        Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
    
    # Ask user for output preference
    Write-Host "`nStep 2: Choose output format" -ForegroundColor Cyan
    Write-Host "1. Single VCF file with all contacts"
    Write-Host "2. Individual VCF files for each contact"
    
    do {
        $choice = Read-Host "Enter your choice (1 or 2)"
    } while ($choice -notin @("1", "2"))
    
    if ($choice -eq "1") {
        # Single file output
        Write-Host "`nStep 3: Choose output file location" -ForegroundColor Cyan
        $outputPath = Get-FilePath -Title "Save VCF file as" -Filter "VCF files (*.vcf)|*.vcf|All files (*.*)|*.*" -Save
        
        if (-not $outputPath) {
            Write-Host "No output file selected. Exiting." -ForegroundColor Red
            Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 1
        }
        
        # Ensure .vcf extension
        if (-not $outputPath.EndsWith(".vcf", [StringComparison]::OrdinalIgnoreCase)) {
            $outputPath += ".vcf"
        }
        
        Write-Host "Converting contacts to single VCF file..." -ForegroundColor Yellow
        
        $allVcfContent = @()
        foreach ($contact in $contacts) {
            $vcfContent = New-VCardContent -Contact $contact
            $allVcfContent += $vcfContent
            $allVcfContent += ""  # Empty line between contacts
        }
        
        # Write to file
        $allVcfContent -join "`r`n" | Out-File -FilePath $outputPath -Encoding UTF8
        Write-Host "Successfully created VCF file: $outputPath" -ForegroundColor Green
        
    } else {
        # Multiple files output
        Write-Host "`nStep 3: Choose output directory" -ForegroundColor Cyan
        
        # Get output directory
        do {
            $outputDir = Read-Host "Enter output directory path"
            if ([string]::IsNullOrWhiteSpace($outputDir)) {
                Write-Host "Please enter a valid directory path." -ForegroundColor Yellow
                continue
            }
            
            $outputDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outputDir)
            
            if (-not (Test-Path $outputDir -PathType Container)) {
                $create = Read-Host "Directory doesn't exist. Create it? (y/n)"
                if ($create -eq "y" -or $create -eq "Y") {
                    try {
                        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
                        break
                    } catch {
                        Write-Host "Error creating directory: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            } else {
                break
            }
        } while ($true)
        
        Write-Host "Converting contacts to individual VCF files..." -ForegroundColor Yellow
        
        $successCount = 0
        foreach ($contact in $contacts) {
            try {
                $vcfContent = New-VCardContent -Contact $contact
                
                # Generate filename
                $contactName = ""
                if ($contact.display_name) {
                    $contactName = $contact.display_name
                } elseif ($contact.'profile.name.first' -or $contact.'profile.name.surname') {
                    $contactName = "$($contact.'profile.name.first') $($contact.'profile.name.surname')".Trim()
                } elseif ($contact.'profile.skype_handle') {
                    $contactName = $contact.'profile.skype_handle'
                } else {
                    $contactName = "Unknown_Contact_$($successCount + 1)"
                }
                
                $safeFileName = Get-SafeFileName -Name $contactName
                $vcfFilePath = Join-Path $outputDir "$safeFileName.vcf"
                
                # Handle duplicate filenames
                $counter = 1
                $originalPath = $vcfFilePath
                while (Test-Path $vcfFilePath) {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($originalPath)
                    $vcfFilePath = Join-Path $outputDir "$baseName`_$counter.vcf"
                    $counter++
                }
                
                # Write VCF file
                $vcfContent | Out-File -FilePath $vcfFilePath -Encoding UTF8
                $successCount++
                
                Write-Host "Created: $vcfFilePath" -ForegroundColor Gray
                
            } catch {
                Write-Host "Error processing contact '$contactName': $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        Write-Host "`nSuccessfully created $successCount VCF files in: $outputDir" -ForegroundColor Green
    }
    
    Write-Host "`nConversion completed successfully!" -ForegroundColor Green

} catch {
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Keep window open at the end
Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
