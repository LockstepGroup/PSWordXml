function New-WordTable {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [Array]$Data,

        [Parameter(Mandatory = $true, Position = 0)]
        [string[]]$Headers,

        [Parameter(Mandatory = $false)]
        [int[]]$ColumnWidths,

        [Parameter(Mandatory = $false)]
        [int]$TableWidth = 5000
    )

    $VerbosePrefix = "New-WordTable:"

    # StartTable
    $OutputXml = @()
    $OutputXml += '<doc>'
    $OutputXml += '    <w:tbl>'

    # Table Style
    $OutputXml += '<w:tblPr>'
    $OutputXml += '    <w:tblStyle w:val="GridTable2-Accent2"/>'
    $OutputXml += '    <w:tblW w:w="' + $TableWidth + '" w:type="pct"/>'
    $OutputXml += '    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
    $OutputXml += '</w:tblPr>'

    # Column Widths
    $OutputXml += '<w:tblGrid>'
    if ($ColumnWidths) {
        foreach ($width in $ColumnWidths) {
            $OutputXml += '<w:gridCol w:w="' + $width + '"/>'
        }
    } else {
        $AverageWidth = $TableWidth / $Headers.Count * 4
        for ($i = 1; $i++; $i -le $Headers.Count) {
            $OutputXml += '<w:gridCol w:w="' + $width + '"/>'
        }
    }
    $OutputXml += '</w:tblGrid>'

    # Close Table
    $OutputXml += '    </w:tbl>'
    $OutputXml += '<doc>'

    # Format Output
    $OutputXml = $OutputXml -join "`n"
    $OutputXml = [xml]$OutputXml
    $Output = $OutputXml.doc.tbl

    # Output
    $Output
}