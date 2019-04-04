function New-WordLineBreak {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Style,

        [Parameter(Mandatory = $false, Position = 1)]
        [string]$Count = 1
    )

    $VerbosePrefix = "New-WordLineBreak:"

    $ParagraphXml = [xml]@"
<doc>
    <w:p w14:paraId="670E6D33" w14:textId="77777777" w:rsidR="004E14AF" w:rsidRPr="004E14AF" w:rsidRDefault="004E14AF" w:rsidP="004E14AF"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:pPr>
            <w:pStyle w:val="$Style" />
        </w:pPr>
        <w:bookmarkStart w:id="0" w:name="_GoBack" />
        <w:bookmarkEnd w:id="0" />
    </w:p>
</doc>
"@

    $ReturnObject = @()
    for ($i = 0; $i -lt $Count; $i++) {
        $ReturnObject += $ParagraphXml.doc.p
    }

    return $ReturnObject
}