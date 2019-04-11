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
        $ParagraphFormattingNode = $RootDocument.CreateNode('element', 'w', 'pPr', $WNamespaceUri)
        $ParagraphStyleNode = $RootDocument.CreateNode('element', 'w', 'pStyle', $WNamespaceUri)

        #endregion XmlSetup
        #############################################################

        #############################################################
        #region Formatting

        # Style
        if ($Style) {
            $ParagraphStyleNode.SetAttribute('val', $WNamespaceUri, $Style) | Out-Null
            $ParagraphFormattingNode.AppendChild($ParagraphStyleNode) | Out-Null
            $ParagraphNode.AppendChild($ParagraphFormattingNode) | Out-Null
        }

        #endregion Formatting
        #############################################################
    }

    Process {
        #############################################################
        #region AddRun

        $ImportNode = $RootDocument.ImportNode($Run, $true)
        $ParagraphNode.AppendChild($ImportNode) | Out-Null

        #endregion AddRun
        #############################################################
    }

    End {
        #############################################################
        #region Output

        # Append to RootDocument so namespaces are summarized properly
        $RootDocument.AppendChild($ParagraphNode) | Out-Null

        # Add Namespaces to document now that there are some contents
        $RootDocument.DocumentElement.SetAttribute('xmlns:w', $WNamespaceUri)
        $RootDocument.DocumentElement.SetAttribute('xmlns:xml', $XmlNamespaceUri)

        # return just the runs
        $RootDocument.p

        #endregion Output
        #############################################################
    }
}