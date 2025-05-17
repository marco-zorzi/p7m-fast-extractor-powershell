# P7M Fast Extractor (PowerShell)
This PowerShell script extracts the content from P7M files and saves it to a specified folder. P7M files are a type of file that is digitally signed and encrypted using the S/MIME standard. This script is specifically designed for fast automated extraction of P7M file contents using PowerShell.

## Features
* Extracts content from P7M files using PowerShell for fast automated extraction.
* Decodes digitally signed and encrypted files and saves them to a specified folder.
* Uses the System.IO and System.Security assemblies to read P7M files and decode the content.

## Getting Started

### Prerequisites
Before running the script, you need to make sure that you have the following installed:
* PowerShell (version 5.0 or later)
* .NET Framework (version 4.6 or later)

### Installation
To use this script, you need to follow these steps:
1. Clone this repository or download the zip file and extract it to a local directory.
2. Open PowerShell and navigate to the directory where the script is located.
3. Run the script by calling the extract_content_from_p7m_file function with the appropriate parameters.

### Script Parameters
| Parameter         | Type   | Description                                      |
|-------------------|--------|--------------------------------------------------|
| `SourceFolder`    | String | Path to the folder containing `.p7m` files       |
| `DestinationFolder`| String| Path to the folder where extracted files are saved |

### Usage
Invoke the refactored cmdlet positionally:

```powershell
.\extract_content_from_p7m_file.ps1 "C:\P7M_Files" "C:\Extracted_Content"
```

Or call the function directly (e.g. from a module):

```powershell
Extract-P7MFileContent -SourceFolder "C:\P7M_Files" -DestinationFolder "C:\Extracted_Content" -Verbose
```

### Example
Suppose `C:\myfiles` contains `file1.p7m`, `file2.p7m`, and `file3.p7m`. To extract them into `C:\extracted_files`, run:
```
.\extract_content_from_p7m_file.ps1 "C:\myfiles" "C:\extracted_files"
```

## License
This script is licensed under the MIT License. You are free to use, modify, and distribute this script as long as you include the license file.

## Contributing
If you find any issues or have suggestions for improvements, feel free to create an issue or pull request. All contributions are welcome!
