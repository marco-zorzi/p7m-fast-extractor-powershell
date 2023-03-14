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
Marco Zorzi

.DATE
2022-Q4

.NOTES
File: extract_content_from_p7m_file.ps1
#>

function extract_content_from_p7m_file($folder_p7m, $folder_extract)
{
    Add-Type -AssemblyName System.IO
    Add-Type -AssemblyName System.Security
    $files = Get-ChildItem $folder_p7m
    $i = 0
    foreach ($file in $files) {
        $i++
        Write-Progress -Activity "Content extraction from $file" -Status "File $($i) of $($($files).count)" -PercentComplete (($i / $($files).count) * 100)
        $file_path = $file.FullName
        $file_basename = $file.Basename
        $content = [IO.File]::ReadAllBytes("$file_path")
        $signedCms = New-Object System.Security.Cryptography.Pkcs.SignedCms
        $signedCms.Decode($content)
        Clear-Variable content
        $infoFile = $signedCms.ContentInfo.Content
        Clear-Variable signedCms
        [IO.File]::WriteAllBytes("$folder_extract\$file_basename", $infoFile)
        Clear-Variable infoFile
    }
}
