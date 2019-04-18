function New-WordRun {
    [cmdletbinding(DefaultParameterSetName = 'Text')]
    Param (
        [Parameter(ParameterSetName = 'Text', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text,

        [Parameter(Mandatory = $false, ValueFromPipeline = $false, Position = 1)]
        [string]$Style,
        #TODO: need to figure out why hyperlinks are getting style quite right.

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$NoNewLine,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$Bold,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$Italic,

        [Parameter(ParameterSetName = 'LineBreak', Mandatory = $false)]
        [switch]$LineBreak,

        [Parameter(ParameterSetName = 'PageBreak', Mandatory = $false)]
        [switch]$PageBreak
        ,

        [Parameter(ParameterSetName = 'Text', Mandatory = $false)]
        [switch]$NoMarkdown
    )

    $VerbosePrefix = "New-WordRun:"
    Write-Verbose "#############################################################"
    Write-Verbose "$VerbosePrefix Starting"


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
    $RunNode = $RootDocument.CreateNode('element', 'w', 'r', $WNamespaceUri)
    $RunFormattingNode = $RootDocument.CreateNode('element', 'w', 'rPr', $WNamespaceUri)
    $BoldNode = $RootDocument.CreateNode('element', 'w', 'b', $WNamespaceUri)
    $ItalicNode = $RootDocument.CreateNode('element', 'w', 'i', $WNamespaceUri)
    $StyleNode = $RootDocument.CreateNode('element', 'w', 'rStyle', $WNamespaceUri)
    $Hyperlink = $RootDocument.CreateNode('element', 'w', 'hyperlink', $WNamespaceUri)
    $TextNode = $RootDocument.CreateNode('element', 'w', 't', $WNamespaceUri)
    $BreakNode = $RootDocument.CreateNode('element', 'w', 'br', $WNamespaceUri)

    #endregion XmlSetup
    #############################################################

    #############################################################
    #region Formatting

    # Bold
    if ($Bold) { $RunFormattingNode.AppendChild($BoldNode) | Out-Null }

    # Italic
    if ($Italic) { $RunFormattingNode.AppendChild($ItalicNode) | Out-Null }

    # Style
    if ($Style) {
        $StyleNode.SetAttribute('val', $WNamespaceUri, $Style) | Out-Null
        $RunFormattingNode.AppendChild($StyleNode) | Out-Null
    }

    # Append RunFormattingNode if any styling is present
    if ($Bold -or $Italic -or $Style) {
        $RunNode.AppendChild($RunFormattingNode) | Out-Null
    }
    #endregion Formatting
    #############################################################

    #############################################################
    #region CreateRunNode
    Switch ($PSCmdlet.ParameterSetName) {
        'Text' {
            if ($NoMarkdown) {
                $ResolvedText = "" | Select-Object -Property Text, Bold, Italic
                $ResolvedText.Text = $Text
                $ResolvedText.Bold = $false
                $ResolvedText.Italic = $false
            } else {
                $ResolvedText = Resolve-WordText -Text $Text
            }
            Write-Verbose "$VerbosePrefix ResolvedText Count: $($ResolvedText.Count)"
            if ($ResolvedText.Count -gt 1) {
                $ResolvedRuns = @()
                $RunCount = 1
                foreach ($t in $ResolvedText) {
                    Write-Verbose "-------------------------------------------------------------"
                    Write-Verbose "$VerbosePrefix Getting ResolvedRun: $($t.Text)"
                    Write-Verbose "$VerbosePrefix Bold: $($t.Bold)"
                    Write-Verbose "$VerbosePrefix Italic: $($t.Italic)"
                    $RunParams = @{}
                    $RunParams.Text = $t.Text
                    $RunParams.Bold = $t.Bold
                    $RunParams.Italic = $t.Italic
                    if ($RunCount -lt $ResolvedText.Count) {
                        Write-Verbose "$VerbosePrefix Adding NoNewLine"
                        $RunParams.NoNewLine = $true
                    }
                    $ResolvedRuns += New-WordRun @RunParams
                    $RunCount++
                }
                Write-Verbose "$VerbosePrefix ResolvedRuns Count: $($ResolvedRuns.Count)"
            } else {
                Write-Verbose "$VerbosePrefix single text entry"
                # NoNewLine
                if ($NoNewLine) {
                    $TextNode.SetAttribute('space', $XmlNamespaceUri, 'preserve') | Out-Null
                }

                # add text
                $TextNode.InnerText = $ExecutionContext.InvokeCommand.ExpandString($ResolvedText.Text)

                # append to RunNode
                $RunNode.AppendChild($TextNode) | Out-Null
            }
        }
        'LineBreak' {
            $RunNode.AppendChild($BreakNode) | Out-Null
        }
        'PageBreak' {
            $BreakNode.SetAttribute('type', $WNamespaceUri, 'page') | Out-Null
            $RunNode.AppendChild($BreakNode) | Out-Null
        }
    }
    #endregion CreateRunNode
    #############################################################

    #############################################################
    #region Output

    # Append to paragraph node for output
    if ($ResolvedText.Count -gt 1) {
        #$Global:ResolvedRuns = $ResolvedRuns
        Write-Verbose "$VerbosePrefix adding multiple runs to paragraph: $($ResolvedRuns.Count)"
        foreach ($run in $ResolvedRuns) {

            $ImportNode = $RootDocument.ImportNode($run, $true)
            $ParagraphNode.AppendChild($ImportNode) | Out-Null
        }
    } else {
        Write-Verbose "$VerbosePrefix adding single run to paragraph: $Text"
        $ParagraphNode.AppendChild($RunNode) | Out-Null
    }

    # Append to RootDocument so namespaces are summarized properly
    $RootDocument.AppendChild($ParagraphNode) | Out-Null

    # Add Namespaces to document now that there are some contents
    $RootDocument.DocumentElement.SetAttribute('xmlns:w', $WNamespaceUri)
    $RootDocument.DocumentElement.SetAttribute('xmlns:xml', $XmlNamespaceUri)

    # return just the runs
    $ParagraphNode.ChildNodes

    #endregion Output
    #############################################################
}