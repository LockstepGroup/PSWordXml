function New-WordTableCell {
    [cmdletbinding(DefaultParameterSetName = 'Text')]
    Param (
        [Parameter(ParameterSetName = 'Text', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        #[AllowEmptyString()]
        [string]$Text,

        [Parameter(ParameterSetName = 'Paragraph', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [System.Xml.XmlElement]$Paragraph
    )

    $VerbosePrefix = "New-WordTableCell:"

    # Start Cell
    $OutputXml = @()
    $OutputXml += '<doc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
    $OutputXml += '    <w:tc>'

    # Formatting
    $OutputXml += '<w:tcPr>'
    $OutputXml += '    <w:cnfStyle w:val="001000000000" w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>'
    $OutputXml += '    <w:tcW w:w="968" w:type="pct"/>'
    $OutputXml += '</w:tcPr>'

    # Close Cell
    $OutputXml += '    </w:tc>'
    $OutputXml += '</doc>'

    # Compile to xml
    $OutputXml = $OutputXml -join "`n"
    $OutputXml = [xml]$OutputXml

    # Content
    if ($Text) {
        $Paragraph = New-WordRun -Text $Text | New-WordParagraph
    }

    $ImportNode = $OutputXml.ImportNode($Paragraph, $true)
    $OutputXml.doc.tc.AppendChild($ImportNode) | Out-Null

    # Output
    $Output = $OutputXml.doc.tc
    $Output
}