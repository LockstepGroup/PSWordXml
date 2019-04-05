function New-WordTableOfContents {
    [cmdletbinding(DefaultParameterSetName = 'Text')]
    Param (
    )

    $VerbosePrefix = "New-WordTableOfContents:"

    $OutputXml = @"
<doc>
    <w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:sdtPr>
            <w:rPr>
                <w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
                <w:color w:val="auto"/>
                <w:sz w:val="22"/>
                <w:szCs w:val="22"/>
            </w:rPr>
            <w:id w:val="-499039493"/>
            <w:docPartObj>
                <w:docPartGallery w:val="Table of Contents"/>
                <w:docPartUnique/>
            </w:docPartObj>
        </w:sdtPr>
        <w:sdtEndPr>
            <w:rPr>
                <w:b/>
                <w:bCs/>
                <w:noProof/>
            </w:rPr>
        </w:sdtEndPr>
        <w:sdtContent>
            <w:p w14:paraId="63F80B50" w14:textId="77777777" w:rsidR="000230DC" w:rsidRDefault="000230DC">
                <w:pPr>
                    <w:pStyle w:val="TOCHeading"/>
                </w:pPr>
                <w:r>
                    <w:t>Table of Contents</w:t>
                </w:r>
            </w:p>
            <w:p w14:paraId="42EEE2A8" w14:textId="008D5EEE" w:rsidR="000230DC" w:rsidRDefault="007617E4">
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                        <w:noProof/>
                    </w:rPr>
                    <w:fldChar w:fldCharType="begin"/>
                </w:r>
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                        <w:noProof/>
                    </w:rPr>
                    <w:instrText xml:space="preserve"> TOC \o "1-2" \h \z \u </w:instrText>
                </w:r>
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                        <w:noProof/>
                    </w:rPr>
                    <w:fldChar w:fldCharType="separate"/>
                </w:r>
                <w:r w:rsidR="002002A8">
                    <w:rPr>
                        <w:noProof/>
                    </w:rPr>
                    <w:t>No table of contents entries found.</w:t>
                </w:r>
                <w:r>
                    <w:rPr>
                        <w:b/>
                        <w:bCs/>
                        <w:noProof/>
                    </w:rPr>
                    <w:fldChar w:fldCharType="end"/>
                </w:r>
            </w:p>
        </w:sdtContent>
    </w:sdt>
</doc>
"@

    $OutputXml = [xml]$OutputXml
    $OutputXml.doc.sdt
}