function New-WordParagraph {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [System.Xml.XmlElement]$Run,

        [Parameter(Mandatory = $false)]
        [string]$Style
    )

    Begin {
        $VerbosePrefix = "New-WordParagraph:"

        #############################################################
        #region XmlSetup

        # Create RootDocument
        [xml]$RootDocument = New-Object System.Xml.XmlDocument

        # NamespaceUris
        # These are needed for setting attributes.
        # Namespaces have to be added to the xml doc after it has contents.
        # So we do that in the output region.
        $WNamespaceUri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $XmlNamespaceUri = 'http://www.w3.org/XML/1998/namespace'

        # Add xml Declaration
        $Declaration = $RootDocument.CreateXmlDeclaration("1.0", "UTF-8", 'yes')
        $RootDocument.AppendChild($Declaration) | Out-Null

        # Stage Xml Nodes
        $ParagraphNode = $RootDocument.CreateNode('element', 'w', 'p', $WNamespaceUri)
        $ParagraphRunFormattingNode = $RootDocument.CreateNode('element', 'w', 'pPr', $WNamespaceUri)

        #endregion XmlSetup
        #############################################################

        Write-Verbose "$VerbosePrefix ParameterSetname: $($PSCmdlet.ParameterSetName)"

        $ParagraphXml = @()
        $ParagraphXml += '<doc xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        $ParagraphXml += '<w:p w14:paraId="" w14:textId="" w:rsidR="" w:rsidRDefault="" w:rsidP="">'
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
                $ImportNode = $ParagraphXml.ImportNode($Run, $true)
                $ParagraphXml.doc.p.AppendChild($ImportNode) | Out-Null
            }
        }
    }

    End {
        return $ParagraphXml.doc.p
    }
}