$exclude = "*.bak,*.bat,*.fxp,*.zip"
# Ensure no spaces before or after comma
$excludeFiles = $exclude.Split(',')
# Ensure no spaces before or after comma
$exclude = "*Documentation*"
$excludeFolders = $exclude.Split(',')

# Get the name of the zip file.
$zipFileName = "Source.zip"

try
{ 
    # Delete the zip file if it exists.
    $exists = Test-Path ($zipFileName)
    if ($exists)
    {
        del ($zipFileName)
    }

    # Loop through all the files in the project folder except those we don't want
    # and add them to a zip file.
    # See https://stackoverflow.com/questions/15294836/how-can-i-exclude-multiple-folders-using-get-childitem-exclude
    # for how to exclude folders when -Recurse is used

    $files = @(Get-ChildItem . -recurse -file -exclude $excludeFiles |
        %{ 
            $allowed = $true
            foreach ($exclude in $excludeFolders)
            { 
                if ((Split-Path $_.FullName -Parent) -ilike $exclude)
                { 
                    $allowed = $false
                    break
                }
            }
            if ($allowed)
            {
                $_
            }
        }
    );

    # See https://stackoverflow.com/questions/51392050/compress-archive-and-preserve-relative-paths to compress

    # exclude directory entries and generate fullpath list
    $filesFullPath = $files | Where-Object -Property Attributes -CContains Archive | ForEach-Object -Process {Write-Output -InputObject $_.FullName}

    #create zip file
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open((Join-Path -Path $(Resolve-Path -Path ".") -ChildPath $zipFileName), [System.IO.Compression.ZipArchiveMode]::Create)

    #write entries with relative paths as names
    foreach ($fname in $filesFullPath)
    {
        $rname = $(Resolve-Path -Path $fname -Relative) -replace '\.\\',''
        $zentry = $zip.CreateEntry($rname)
        $zentryWriter = New-Object -TypeName System.IO.BinaryWriter $zentry.Open()
        $zentryWriter.Write([System.IO.File]::ReadAllBytes($fname))
        $zentryWriter.Flush()
        $zentryWriter.Close()
    }

    # clean up
    Get-Variable -exclude Runspace | Where-Object {$_.Value -is [System.IDisposable]} | Foreach-Object {$_.Value.Dispose(); Remove-Variable $_.Name};
}

catch
{
    Write-Host "Error occurred at $(Get-Date): $($Error[0].Exception.Message)"
	pause
}
