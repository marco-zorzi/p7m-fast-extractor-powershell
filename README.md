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

### Usage
To use the extract_content_from_p7m_file function, you need to provide two parameters:
* `$folder_p7m`: the directory where the P7M files are located.
* `$folder_extract`: the directory where the extracted files should be saved.
```
extract_content_from_p7m_file -folder_p7m "C:\path\to\p7m\files" -folder_extract "C:\path\to\extract\folder"
```

### Example
Suppose you have a directory called `C:\myfiles` that contains three P7M files: `file1.p7m`, `file2.p7m`, and `file3.p7m`. You want to extract the contents of these files and save them to a directory called `C:\extracted_files`. To do this, you would run the following command:
```
extract_content_from_p7m_file -folder_p7m "C:\myfiles" -folder_extract "C:\extracted_files"
```

## License
This script is licensed under the MIT License. You are free to use, modify, and distribute this script as long as you include the license file.

## Contributing
If you find any issues or have suggestions for improvements, feel free to create an issue or pull request. All contributions are welcome!
