# CSV to VCF Converter

> **Cross-platform PowerShell script for converting CSV contacts to VCF format**

A powerful, user-friendly tool that converts CSV contact exports (especially from Skype) to VCF (vCard) format for easy import into contact management applications. Works seamlessly on Windows, Linux, and macOS with interactive file selection and comprehensive field mapping.

**Keywords:** `powershell` `csv` `vcf` `vcard` `contacts` `skype` `converter` `cross-platform`

## Features

- ✅ **Cross-platform compatibility** - Works on Windows, Linux, and macOS
- ✅ **Interactive file selection** - GUI file picker on Windows, console prompts on other platforms
- ✅ **Multiple output formats** - Single VCF file or individual files per contact
- ✅ **Comprehensive contact mapping** - Handles names, phone numbers, emails, websites, notes, and more
- ✅ **Skype CSV support** - Specifically designed for Skype contact exports
- ✅ **Phone number detection** - Automatically detects and converts various phone number field formats
- ✅ **Email support** - Handles multiple email field formats
- ✅ **Duplicate handling** - Removes duplicate phone numbers and handles duplicate filenames
- ✅ **Error handling** - Graceful fallbacks and user-friendly error messages

## Supported Contact Fields

The script converts the following contact information from CSV to VCF format:

| CSV Field Types | VCF Property | Description |
|----------------|--------------|-------------|
| `display_name`, `profile.name.first`, `profile.name.surname` | `FN`, `N` | Full name and structured name |
| `phones[x].number`, `profile.phones[x].number` | `TEL` | Phone numbers with type detection |
| `email`, `profile.email`, etc. | `EMAIL` | Email addresses |
| `profile.skype_handle` | `X-SKYPE`, `IMPP` | Skype username |
| `profile.website` | `URL` | Website/homepage |
| `profile.about` | `NOTE` | About/bio information |
| `profile.avatar_url` | `PHOTO` | Profile picture URL |
| `profile.locations[0].country` | `ADR` | Country/location |
| `creation_time` | `REV` | Contact creation date |

## Requirements

- **PowerShell 5.0+** (Windows PowerShell) or **PowerShell Core 6.0+** (cross-platform)
- No additional dependencies required

## Installation

1. Download or clone this repository
2. **For Windows users using the batch file**: No additional setup required
3. **For direct PowerShell execution**: Ensure PowerShell execution policy allows script execution:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

### Windows Users

**Option 1: Double-click the batch file (easiest, no setup required)**
```
Double-click Run-CSV2VCF.bat
```

**Option 2: Right-click PowerShell execution**
```
Right-click Convert-SkypeCSVToVCF.ps1 → "Run with PowerShell"
```

**Option 3: Command line**
```powershell
.\Convert-SkypeCSVToVCF.ps1
```

### Linux/macOS Users

```bash
pwsh ./Convert-SkypeCSVToVCF.ps1
```

### Interactive Process

1. **Select CSV file** - Choose your Skype CSV export file
2. **Choose output format**:
   - Single VCF file with all contacts
   - Individual VCF files for each contact
3. **Select output location** - Choose where to save the VCF file(s)
4. **Conversion** - The script processes all contacts and provides feedback

## CSV Format Support

The script is designed for Skype CSV exports but can handle various CSV formats. It automatically detects:

- **Phone number fields**: `phones[x].number`, `profile.phones[x].number`, `mobile`, `cell_phone`, `home_phone`, `work_phone`, etc.
- **Email fields**: `email`, `profile.email`, `home_email`, `work_email`, etc.
- **Name fields**: `display_name`, `profile.name.first`, `profile.name.surname`, etc.

### Example CSV Structure

```csv
type,id,display_name,phones[0].number,phones[0].type,profile.name.first,profile.name.surname,profile.skype_handle,creation_time
PhoneNumber,15551234567,John Smith,15551234567,mobile,John,Smith,john.smith123,2023-01-15 10:30:00Z
```

## Output Formats

### VCF 3.0 Format
The script generates VCF 3.0 compatible files with the following properties:
- `BEGIN:VCARD` / `END:VCARD`
- `VERSION:3.0`
- `FN` (Full Name)
- `N` (Structured Name)
- `TEL` (Phone numbers with type)
- `EMAIL` (Email addresses)
- `X-SKYPE` / `IMPP` (Skype handle)
- `URL` (Website)
- `NOTE` (About/bio)
- `PHOTO` (Avatar URL)
- `ADR` (Address/Country)
- `REV` (Revision date)

### Phone Number Types
- `mobile`/`cell` → `TEL;TYPE=CELL`
- `home` → `TEL;TYPE=HOME`
- `work`/`business` → `TEL;TYPE=WORK`
- Default → `TEL;TYPE=VOICE`

## Troubleshooting

### Windows PowerShell Issues
- **File dialog doesn't appear**: The script will fall back to console input automatically
- **Window closes immediately**: Use the batch file or the script now includes "Press any key" prompts
- **Execution policy error**: Use the batch file (bypasses this issue) or run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### General Issues
- **No contacts found**: Verify your CSV file has the correct format and headers
- **Missing phone numbers**: Check if your CSV uses different field names (the script will show detected fields)
- **Special characters**: The script handles most special characters in names and notes

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Thanks

Thanks to James at [mediamonarchy.com](https://www.mediamonarchy.com) for helping test and improve the script. 
