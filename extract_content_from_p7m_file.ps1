<#
.SYNOPSIS
Extracts the content from P7M files and saves it to a specified folder.

.DESCRIPTION
This PowerShell script is specifically designed for fast extraction of P7M file contents using PowerShell. P7M files are a type of file that is digitally signed and encrypted using S/MIME standard. The script uses the System.IO and System.Security assemblies to read the P7M files and decode the content. It then saves the extracted content to the specified folder. This script can be useful for anyone who needs to extract the contents of P7M files in an automated way.

.PARAMETER folder_p7m
The folder containing the P7M files to extract content from.

.PARAMETER folder_extract
The folder where the extracted content will be saved.

.EXAMPLE
.\extract_content_from_p7m_file.ps1 -folder_p7m "C:\P7M_Files" -folder_extract "C:\Extracted_Content"

.AUTHOR
Marco Zorzi (github.mz@gmail.com)

.DATE
2022-Q4

.NOTES
File: extract_content_from_p7m_file.ps1
#>

function extract_content_from_p7m_file($folder_p7m, $folder_extract)
{
    # Add the 'System.IO' and 'System.Security' assemblies
    Add-Type -AssemblyName System.IO
    Add-Type -AssemblyName System.Security
    
    # Get a list of P7M files in the '$folder_p7m' directory
    $files = Get-ChildItem $folder_p7m -Filter "*.p7m"
    
    # Initialize a counter variable to keep track of progress
    $i = 0
    
    # Iterate over each file in the directory
    foreach ($file in $files) {
        # Increment the counter variable
        $i++
        
        # Display a progress bar indicating the current file being processed
        Write-Progress -Activity "Content extraction from $file" -Status "File $($i) of $($($files).count)" -PercentComplete (($i / $($files).count) * 100)
        
        # Get the full path to the file
        $file_path = $file.FullName
       
        # Get the base name of the file
        $file_basename = $file.Basename
        
        # Read the contents of the file into a byte array
        $content = [IO.File]::ReadAllBytes("$file_path")
        
        # Create a new 'SignedCms' object and decode the contents of the file
        $signedCms = New-Object System.Security.Cryptography.Pkcs.SignedCms
        $signedCms.Decode($content)
        
        # Clear the 'content' variable to free up memory
        Clear-Variable content
        
        # Get the content of the signed CMS message
        $infoFile = $signedCms.ContentInfo.Content
        
        # Clear the 'signedCms' variable to free up memory
        Clear-Variable signedCms
        
        # Write the content of the message to a file in the '$folder_extract' directory
        [IO.File]::WriteAllBytes("$folder_extract\$file_basename", $infoFile)
        
        # Clear the 'infoFile' variable to free up memory
        Clear-Variable infoFile
    }
}
