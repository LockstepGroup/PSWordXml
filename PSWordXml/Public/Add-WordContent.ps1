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

        $DocumentXmlPath = Join-Path -Path $global:OpenWordDocument -ChildPath word/document.xml
        $DocumentContents = [xml](Get-Content -Path $DocumentXmlPath)
    }
    Process {
        Write-Verbose "$VerbosePrefix adding content"
        $ImportNode = $DocumentContents.ImportNode($Content, $true)
        $DocumentContents.document.body.AppendChild($ImportNode) | Out-Null
    }

    End {
        $DocumentContents.OuterXml | Out-File -FilePath $DocumentXmlPath -Force
    }
}