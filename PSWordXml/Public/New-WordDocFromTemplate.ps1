function New-WordDocFromTemplate {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 0)]
        [string]$TemplatePath,

        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 1)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$DocumentTitle,

        [Parameter(Mandatory = $true)]
        [string]$Company,

        [Parameter(Mandatory = $true)]
        [string]$AuthorName,

        [Parameter(Mandatory = $true)]
        [string]$AuthorEmail
    )

    $OperatingSystem = Get-OsVersion


    #############################################################
    #region LoadDocument

    # MacOS (and presumably linux) uses TMPDIR, windows uses TMP
    switch -Regex ($OperatingSystem) {
        'MacOS' {
            $TempDirectory = $env:TMPDIR
        }
        default {
            $TempDirectory = $env:TMP
        }
    }

    # extract document
    $ResolvedOutputDirectory = Split-Path -Path $OutputPath | Resolve-Path
    $ResolvedTemplatePath = Resolve-Path -Path $TemplatePath
    $ExtractDirectory = Join-Path -Path (Split-Path -Path $ResolvedOutputDirectory) -ChildPath 'WordXml'
    # cleanup old junk if needed
    if (Test-Path -Path $ExtractDirectory) {
        Remove-Item -Path $ExtractDirectory -Recurse -Force
    }

    # create temp dir for extracting docx file
    New-Item -Path $ExtractDirectory -ItemType Directory | Out-Null

    # unzip docx
    switch -Regex ($OperatingSystem) {
        'MacOS' {
            unzip $ResolvedTemplatePath -d $ExtractDirectory | Out-Null
            break
        }
        default {
            Throw "need to add code for unzipping on non-macos"
        }
    }

    # get main doc xml
    $MainDocumentPath = Join-Path -Path $ExtractDirectory -ChildPath "word/document.xml"
    $DocumentContents = [xml](Get-Content -Path $MainDocumentPath -Raw)
    $Body = $DocumentContents.document.body
    $Paragraphs = $Body.p

    #endregion LoadDocument
    #############################################################

    #############################################################
    #region TitlePage
    $ParagraphMappings = @{
        DocumentTitle = 6
        Company       = 7
        Author        = 17
    }

    $Paragraphs[$ParagraphMappings.DocumentTitle].r.t = $DocumentTitle
    $Paragraphs[$ParagraphMappings.Company].r.t = $Company
    $Paragraphs[$ParagraphMappings.Author].r[0].t = $AuthorName
    $Paragraphs[$ParagraphMappings.Author].hyperlink.r.t = $AuthorEmail

    #endregion TitlePage
    #############################################################

    #############################################################
    #region Output

    $ExtractedContentPath = Join-Path -Path $ExtractDirectory -ChildPath '*'
    $TemporaryZipPath = Join-Path $ExtractDirectory -ChildPath 'doc.zip'
    $DocumentContents.OuterXml | Out-File -FilePath $MainDocumentPath

    # zip docx
    switch -Regex ($OperatingSystem) {
        'MacOS' {
            # had to do this to stop zip from adding the root dir to the zip file
            Push-Location -Path $ExtractDirectory
            zip -r $TemporaryZipPath ./* | Out-Null
            Pop-Location
            Move-Item -Path $TemporaryZipPath -Destination $OutputPath -Force
            break
        }
        default {
            Throw "need to add code for zipping on non-macos"
        }
    }

    #endregion Output
    #############################################################
}