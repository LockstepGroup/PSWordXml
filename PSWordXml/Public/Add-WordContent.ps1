function Add-WordContent {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 0)]
        [System.Xml.XmlElement]$Content
    )

    $VerbosePrefix = "Add-WordContent:"

    if (!($global:OpenWordDocument)) {
        Throw "$VerbosePrefix No open Word Document, use Open-WordDocument to get started."
    }

    $DocumentXmlPath = Join-Path -Path $global:OpenWordDocument -ChildPath word/document.xml
    $DocumentContents = [xml](Get-Content -Path $DocumentXmlPath)
    $ImportNode = $DocumentContents.ImportNode($Content, $true)
    $DocumentContents.document.body.AppendChild($ImportNode) | Out-Null
    $DocumentContents.OuterXml | Out-File -FilePath $DocumentXmlPath -Force
}