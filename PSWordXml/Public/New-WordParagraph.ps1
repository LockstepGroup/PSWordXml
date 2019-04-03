function New-WordParagraph {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 0)]
        [string]$Text
    )

    $VerbosePrefix = "New-WordParagraph:"

    $ParagraphXml = [xml]@"
<doc>
    <w:p w14:paraId="" w14:textId="" w:rsidR="" w:rsidRDefault="" w:rsidP=""
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>$Text</w:t>
        </w:r>
        <w:proofErr w:type="gramEnd" />
    </w:p>
</doc>
"@

    return $ParagraphXml.doc.p
}