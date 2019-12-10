function New-WordTable {
    [cmdletbinding(DefaultParameterSetName = 'row')]
    Param (
        [Parameter(ParameterSetName = 'array', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [psobject]$DataArray,

        [Parameter(ParameterSetName = 'row', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [System.Xml.XmlElement]$Row,

        [Parameter(Mandatory = $true, Position = 0)]
        [string[]]$Headers,

        [Parameter(Mandatory = $false)]
        [int[]]$ColumnWidths,

        [Parameter(Mandatory = $false)]
        [int]$TableWidth = 5000
    )

    Begin {
        $VerbosePrefix = "New-WordTable:"

        # StartTable
        $OutputXml = @()
        $OutputXml += '<doc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        $OutputXml += '    <w:tbl>'

        # Table Style
        Write-Verbose "$VerbosePrefix Create tblPr"
        $OutputXml += '<w:tblPr>'
        $OutputXml += '    <w:tblStyle w:val="GridTable2-Accent2"/>'
        $OutputXml += '    <w:tblW w:w="' + $TableWidth + '" w:type="pct"/>'
        $OutputXml += '    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        $OutputXml += '</w:tblPr>'

        # Column Widths
        Write-Verbose "$VerbosePrefix Create tblGrid"
        $OutputXml += '<w:tblGrid>'
        if ($ColumnWidths) {
            Write-Verbose "$VerbosePrefix Widths specified"
            foreach ($width in $ColumnWidths) {
                $OutputXml += '<w:gridCol w:w="' + $width + '"/>'
            }
        } else {
            $AverageWidth = $TableWidth / $Headers.Count * 4
            Write-Verbose "$VerbosePrefix Widths not specified, taking average: $AverageWidth"
            Write-Verbose "$VerbosePrefix Header Count: $($Headers.Count)"
            for ($i = 1; $i -le $Headers.Count; $i++) {
                Write-Verbose "$VerbosePrefix $i"
                $OutputXml += '<w:gridCol w:w="' + $AverageWidth + '"/>'
            }
        }
        $OutputXml += '</w:tblGrid>'
        $TheseRows = @()

        # Header Row
        $Cells = @()
        foreach ($header in $Headers) {
            $Cells += New-WordTableCell -Text $header
        }
        $TheseRows += $Cells | New-WordTableRow
    }

    Process {

        Write-Verbose "$VerbosePrefix Processing input for ParameterSet: $($PSCmdlet.ParameterSetName)"
        Switch ($PSCmdlet.ParameterSetName) {
            row {
                $TheseRows += $Row
            }
            array {
                foreach ($entry in $DataArray) {
                    $Cells = @()
                    foreach ($header in $Headers) {
                        Write-Verbose "$VerbosePrefix adding header: $header"
                        if ($entry.$header) {
                            if (($entry.$header).GetType().FullName -eq 'System.Xml.XmlElement') {
                                $Cells += New-WordTableCell -Paragraph $entry.$header
                            } else {
                                Write-Verbose "$VerbosePrefix provided value is plain text"
                                Write-Verbose "$VerbosePrefix adding cell with value: $($entry.header)"
                                $Cells += New-WordTableCell -Text $entry.$header
                            }
                        }
                    }
                    $TheseRows += $Cells | New-WordTableRow
                }
            }
        }
    }

    End {
        # Close Table
        $OutputXml += '    </w:tbl>'
        $OutputXml += '</doc>'

        # Format Output
        $OutputXml = $OutputXml -join "`n"
        $OutputXml = [xml]$OutputXml

        # Import Rows
        foreach ($row in $TheseRows) {
            $ImportNode = $OutputXml.ImportNode($row, $true)
            $OutputXml.doc.tbl.AppendChild($ImportNode) | Out-Null
        }

        # Output
        $Output = $OutputXml.doc.tbl
        $Output
    }
}