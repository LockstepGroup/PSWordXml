function New-WordRun {
    [cmdletbinding(DefaultParameterSetName = 'Text')]
    Param (
        [Parameter(ParameterSetName = 'Text', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 1)]
        [string]$Style,
        #TODO: need to figure out why hyperlinks are getting style quite right.

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$NoNewLine,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$Bold,

        [Parameter(ParameterSetName = 'LineBreak', Mandatory = $false)]
        [switch]$LineBreak,

        [Parameter(ParameterSetName = 'PageBreak', Mandatory = $false)]
        [switch]$PageBreak
    )

    $VerbosePrefix = "New-WordRun:"

    if ($NoNewLine) {
        $Preserve = ' xml:space="preserve"'
    } else {
        $Preserve = ''
    }

    $OutputXml = @()
    $OutputXml += '<doc>'
    $OutputXml += '<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    if ($Style -eq 'HyperLink') {
        $OutputXml += '    <w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:history="1">'
    }
    $OutputXml += '    <w:r>'
    if ($Style -or $Bold) {
        $OutputXml += '        <w:rPr>'
    }
    if ($Style) {
        $OutputXml += '             <w:rStyle w:val="' + $Style + '"/>'
    }
    if ($Bold) {
        $OutputXml += '             <w:b/>'
    }
    if ($Style -or $Bold) {
        $OutputXml += '        </w:rPr>'
    }

    #############################################################
    #region Text

    if (-not $LineBreak -and -not $PageBreak) {
        $OutputXml += '        <w:t' + $Preserve + '>' + $Text + '</w:t>'
    }

    #endregion Text
    #############################################################

    #############################################################
    #region PageBreak

    if ($LineBreak) {
        $OutputXml += '        <w:br/>'
    }

    #endregion PageBreak
    #############################################################

    #############################################################
    #region PageBreak

    if ($PageBreak) {
        $OutputXml += '        <w:br w:type="page"/>'
    }

    #endregion PageBreak
    #############################################################

    $OutputXml += '    </w:r>'
    if ($Style -eq 'HyperLink') {
        $OutputXml += '    </w:hyperlink>'
    }
    $OutputXml += '</w:p>'
    $OutputXml += '</doc>'

    $OutputXml = $OutputXml -join "`n"
    $OutputXml = [xml]$OutputXml

    if ($Style -eq 'HyperLink') {
        Write-Verbose "$VerbosePrefix Style is HyperLink, returning hyperlink node"
        $Output = $OutputXml.doc.p.hyperlink
    } else {
        Write-Verbose "$VerbosePrefix Style is not HyperLink, returning r node"
        $Output = $OutputXml.doc.p.r
    }

    $Output
}