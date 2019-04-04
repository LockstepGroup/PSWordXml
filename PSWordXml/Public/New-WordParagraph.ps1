function New-WordParagraph {
    [cmdletbinding(DefaultParameterSetName = 'Run')]
    Param (
        [Parameter(ParameterSetName = 'Text', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$Bold,

        [Parameter(ParameterSetName = 'Run', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [System.Xml.XmlElement]$Run,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [Parameter(ParameterSetName = 'Run', Mandatory = $false)]
        [string]$Style
    )

    Begin {
        $VerbosePrefix = "New-WordParagraph:"
        Write-Verbose "$VerbosePrefix ParameterSetname: $($PSCmdlet.ParameterSetName)"

        $ParagraphXml = @()
        $ParagraphXml += '<doc>'
        $ParagraphXml += '<w:p w14:paraId="" w14:textId="" w:rsidR="" w:rsidRDefault="" w:rsidP=""'
        $ParagraphXml += '    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        $ParagraphXml += '    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        if ($Style) {
            $ParagraphXml += '    <w:pPr>'
            $ParagraphXml += '        <w:pStyle w:val="' + $Style + '" />'
            $ParagraphXml += '    </w:pPr>'
        }

        Switch ($PSCmdlet.ParameterSetName) {
            'Run' {
                # close out xml
                $ParagraphXml += '    </w:p>'
                $ParagraphXml += '</doc>'

                # convert text to xml
                $ParagraphXml = $ParagraphXml -join "`n"
                $ParagraphXml = [xml]$ParagraphXml
            }
        }
    }

    # begin process
    Process {
        Switch ($PSCmdlet.ParameterSetName) {
            'Text' {
                $SplitTextOnLineBreak = $Text.Split("`r`n")
                $i = 0
                foreach ($line in $SplitTextOnLineBreak) {
                    $i++
                    $ParagraphXml += '    <w:r>'
                    if ($Bold) {
                        $ParagraphXml += '    <w:rPr>'
                        $ParagraphXml += '        <w:b/>'
                        $ParagraphXml += '    </w:rPr>'
                    }
                    $ParagraphXml += '        <w:t>' + $line + '</w:t>'
                    $ParagraphXml += '    </w:r>'

                    if ($i -lt $SplitTextOnLineBreak.Count) {
                        # add line break for all except last line
                        $ParagraphXml += '    <w:r w:rsidR="">'
                        $ParagraphXml += '        <w:br />'
                        $ParagraphXml += '    </w:r>'
                    }
                }

                # close out xml
                $ParagraphXml += '    </w:p>'
                $ParagraphXml += '</doc>'

                # convert text to xml
                $ParagraphXml = $ParagraphXml -join "`n"
                $ParagraphXml = [xml]$ParagraphXml
            }
            'Run' {
                #import run
                Write-Verbose "$VerbosePrefix adding run"
                $global:testpara = $ParagraphXml
                $global:testrun = $Run
                $ImportNode = $ParagraphXml.ImportNode($Run, $true)
                $ParagraphXml.doc.p.AppendChild($ImportNode) | Out-Null
            }
        }
    }

    End {
        return $ParagraphXml.doc.p
    }
}