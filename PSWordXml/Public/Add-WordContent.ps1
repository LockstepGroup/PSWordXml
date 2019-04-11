function Add-WordContent {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 0)]
        [System.Xml.XmlElement]$Content
    )

    Begin {
        $VerbosePrefix = "Add-WordContent:"

        if (!($global:OpenWordDocument)) {
            Throw "$VerbosePrefix No open Word Document, use Open-WordDocument to get started."
        }

        #############################################################
        #region XmlSetup

        # Import Xml from $global:OpenWordDocument
        $DocumentXmlPath = Join-Path -Path $global:OpenWordDocument -ChildPath word/document.xml
        $RootDocument = [xml](Get-Content -Path $DocumentXmlPath)

        # Add Namespaces to document now that there are some contents
        $WNamespaceUri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $XmlNamespaceUri = 'http://www.w3.org/XML/1998/namespace'
        $RootDocument.DocumentElement.SetAttribute('xmlns:w', $WNamespaceUri)
        $RootDocument.DocumentElement.SetAttribute('xmlns:xml', $XmlNamespaceUri)

        #endregion XmlSetup
        #############################################################


    }
    Process {
        #############################################################
        #region AddNodes

        Write-Verbose "$VerbosePrefix adding content"
        $ImportNode = $RootDocument.ImportNode($Content, $true)
        $RootDocument.document.body.AppendChild($ImportNode) | Out-Null

        #endregion AddNodes
        #############################################################
    }

    End {
        #############################################################
        #region Output

        # Output back to document.xml
        $global:root = $RootDocument
        $RootDocument.OuterXml | Out-File -FilePath $DocumentXmlPath -Force

        #endregion Output
        #############################################################
    }
}