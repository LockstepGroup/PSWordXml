function New-WordTableRow {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [System.Xml.XmlElement]$Cell
    )



    Begin {
        $VerbosePrefix = "New-WordTableRow:"

        # Start Cell
        $OutputXml = @()
        $OutputXml += '<doc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        $OutputXml += '    <w:tr>'

        # Formatting
        $OutputXml += '<w:trPr>'
        $OutputXml += '    <w:cnfStyle w:val="100000000000" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="0" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>'
        $OutputXml += '</w:trPr>'

        # Close Cell
        $OutputXml += '    </w:tr>'
        $OutputXml += '</doc>'

        # Compile to xml
        $OutputXml = $OutputXml -join "`n"
        $OutputXml = [xml]$OutputXml
    }

    Process {
        # Content
        Write-Verbose "$VerbosePrefix adding cell"
        $ImportNode = $OutputXml.ImportNode($Cell, $true)
        $OutputXml.doc.tr.AppendChild($ImportNode) | Out-Null
    }

    End {
        # Output
        $Output = $OutputXml.doc.tr
        $Output
    }
}