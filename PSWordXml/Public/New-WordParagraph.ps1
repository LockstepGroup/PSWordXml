function New-WordParagraph {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text,

        [Parameter(Mandatory = $true, ValueFromPipeline = $false, Position = 1)]
        [string]$Style
    )

    $VerbosePrefix = "New-WordParagraph:"

    $ParagraphXml = [xml]@"
<doc>
    <w:p w14:paraId="" w14:textId="" w:rsidR="" w:rsidRDefault="" w:rsidP=""
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
            <w:pStyle w:val="$Style" />
        </w:pPr>
        <w:r>
            <w:t>$Text</w:t>
        </w:r>
        <w:proofErr w:type="gramEnd" />
    </w:p>
</doc>
"@

    return $ParagraphXml.doc.p
}