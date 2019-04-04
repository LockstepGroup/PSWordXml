function New-WordRun {
    [cmdletbinding(DefaultParameterSetName = 'Text')]
    Param (
        [Parameter(ParameterSetName = 'Text', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 1)]
        [string]$Style,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$NoNewLine,

        [Parameter(ParameterSetName = 'LineBreak', Mandatory = $false)]
        [switch]$LineBreak,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$Bold
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
    if ($Style) {
        $OutputXml += '        <w:rPr>'
        $OutputXml += '             <w:rStyle w:val="' + $Style + '"/>'
        $OutputXml += '        </w:rPr>'
    }
    if ($LineBreak) {
        $OutputXml += '        <w:br/>'
    } else {
        $OutputXml += '        <w:t' + $Preserve + '>' + $Text + '</w:t>'
    }
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

<#

<w:rPr>
                    <w:b/>
                </w:rPr>
<w:p w14:paraId="346B3CB9" w14:textId="69F7CD0E" w:rsidR="00577F56" w:rsidRPr="00577F56" w:rsidRDefault="00867423" w:rsidP="000B2256"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:pPr>
        <w:pStyle w:val="Subtitle" />
    </w:pPr>
    <w:r>
        <w:t>Josh Sanders</w:t>
    </w:r>
    <w:r w:rsidR="00342F64">
        <w:br />
    </w:r>
    <w:hyperlink r:id="rId9" w:history="1"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:r w:rsidR="00577F56" w:rsidRPr="00FD02CE">
            <w:rPr>
                <w:rStyle w:val="Hyperlink" />
            </w:rPr>
            <w:t>jsanders@lockstepgroup.com</w:t>
        </w:r>
    </w:hyperlink>
</w:p>
#>