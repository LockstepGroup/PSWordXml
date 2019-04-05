function Close-WordDocument {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 0)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $VerbosePrefix = "Close-WordDocument:"
    $OperatingSystem = Get-OsVersion

    #############################################################
    #region Output

    if (!($global:OpenWordDocument)) {
        Throw "$VerbosePrefix No open Word Document, use Open-WordDocument to get started."
    }

    $OutputDirectory = Split-Path -Path $global:OpenWordDocument
    $ExtractedContentPath = Join-Path -Path $global:OpenWordDocument -ChildPath '*'
    $TemporaryZipPath = Join-Path $OutputDirectory -ChildPath 'doc.zip'

    # zip docx
    switch -Regex ($OperatingSystem) {
        'MacOS' {
            # had to do this to stop zip from adding the root dir to the zip file
            Push-Location -Path $global:OpenWordDocument
            zip -r $TemporaryZipPath ./* | Out-Null
            #Compress-Archive -Path ./* -DestinationPath $TemporaryZipPath
            Pop-Location
            Move-Item -Path $TemporaryZipPath -Destination $OutputPath -Force:$Force
            break
        }
        default {
            Throw "need to add code for zipping on non-macos"
        }
    }


    Remove-Item -Path $global:OpenWordDocument -Recurse -Force:$Force
    Remove-Variable -Name OpenWordDocument -Scope global

    #endregion Output
    #############################################################
}