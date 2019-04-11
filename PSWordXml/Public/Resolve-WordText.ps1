function Resolve-WordText {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string]$Text
    )

    $VerbosePrefix = "Resolve-WordText:"



    $Output
}