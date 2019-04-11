function Resolve-WordText {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text
    )

    $VerbosePrefix = "Resolve-WordText:"

    $BoldOutput = @()

    function IsEven($i) {
        if (($i % 2) -eq 0) {
            $true
        } else {
            $false
        }
    }

    # process bold text
    $BoldSplit = $Text.Split('**')
    $i = 0
    foreach ($run in $BoldSplit) {
        $i++
        $new = "" | Select-Object -Property Text, Bold, Italic
        $new.Text = $run
        $new.Bold = $false
        $new.Italic = $false
        if (IsEven $i) {
            $new.Bold = $true
        }
        $BoldOutput += $new
    }

    # process italic text
    $ItalicOutput = @()
    foreach ($boldrun in $BoldOutput) {
        $ItalicSplit = $boldrun.Text.Split('_')
        $i = 0
        foreach ($run in $ItalicSplit) {
            $i++
            if ($run -eq "") { continue }
            $new = "" | Select-Object -Property Text, Bold, Italic
            $new.Text = $run
            $new.Bold = $boldrun.Bold
            $new.Italic = $false
            if (IsEven $i) {
                $new.Italic = $true
            }
            $ItalicOutput += $new
        }
    }

    $ItalicOutput
}