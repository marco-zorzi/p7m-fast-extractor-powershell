<#
.SYNOPSIS
Extracts content from P7M files and saves them to a specified folder.

.DESCRIPTION
This PowerShell script processes S/MIME signed .p7m files and extracts the original content to the destination folder.
It uses .NET cryptography classes to decode the message and logs the process with detailed progress and error handling.

.PARAMETER SourceFolder
The folder containing the .p7m files.

.PARAMETER DestinationFolder
The folder where the extracted content will be saved.

.EXAMPLE
.\extract_content_from_p7m_file.ps1 "C:\P7M_Files" "C:\Extracted_Content"

.AUTHOR
Marco Zorzi (github.mz@gmail.com)

.DATE
2022-Q4

.NOTES
File: extract_content_from_p7m_file.ps1
#>

function Extract-P7MFileContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateScript({ Test-Path $_ -PathType Container })]
        [string]$SourceFolder,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()]
        [string]$DestinationFolder
    )

    # Get list of files once
    $p7mFiles = Get-ChildItem -Path $SourceFolder -Filter '*.p7m'
    $total = $p7mFiles.Count															 
    $count = 0
    
    foreach ($file in $p7mFiles) {										
        $count++																			
        Write-Progress -Activity "Extracting P7M files" -Status "$($file.Name) ($count of $total)" -PercentComplete (($count / $total) * 100)
        if ($file.Length -eq 0) {
            Write-Warning "$($file.Name) is empty. Skipping."
            continue
        }
        try {														 
            $bytes = [IO.File]::ReadAllBytes($file.FullName)																			 
            $cms = [System.Security.Cryptography.Pkcs.SignedCms]::new()
            $cms.Decode($bytes)
            if ($cms.Detached) {
                Write-Warning "$($file.Name) is detached. Skipping."
                continue
            }
            $outputPath = Join-Path $DestinationFolder ($file.BaseName)
            [IO.File]::WriteAllBytes($outputPath, $cms.ContentInfo.Content)
            Write-Verbose "Extracted to $outputPath"
        } catch {
            Write-Error "Error processing file $($file.Name): $_"
        }					   
    }
    Write-Host "Extraction complete. Processed $count file(s)."
}

# ——— Script Entrypoint Using $args ———

Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Security

if ($MyInvocation.MyCommand.Path -eq $PSCommandPath) {
    if ($args.Count -ne 2) {
        Write-Host "Usage: .\extract_content_from_p7m_file.ps1 <SourceFolder> <DestinationFolder>" -ForegroundColor Yellow
        exit 1
    }

    $source = $args[0]
    $destination = $args[1]
    Extract-P7MFileContent -SourceFolder $source -DestinationFolder $destination
}
