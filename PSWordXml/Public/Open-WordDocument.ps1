function Open-WordDocument {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 0)]
        [string]$Path,

        [Parameter(Mandatory = $true, ValueFromPipeline = $True, Position = 1)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $VerbosePrefix = "Open-WordDocument:"
    $OperatingSystem = Get-OsVersion

    #############################################################
    #region LoadDocument

    # create output path
    if (!(Test-Path -Path $OutputPath)) {
        if ($Force) {
            New-Item -Path $OutputPath -ItemType Directory | Out-Null
        } else {
            Throw "$VerbosePrefix OutputPath does not exist, use -Force to create automatically."
        }
    } elseif ((Get-ChildItem -Path $OutputPath).Count -gt 0) {
        if ($Force) {
            Remove-Item -Path "$OutputPath/*" -Recurse -Force
        } else {
            Throw "$VerbosePrefix OutputPath is not empty, use -Force to overwrite."
        }
    }
    $ResolvedOutputPath = Resolve-Path -Path $OutputPath
    $ResolvedInputPath = Resolve-Path -Path $Path

    Expand-Archive -Path $ResolvedInputPath -DestinationPath $ResolvedOutputPath

    $global:OpenWordDocument = $ResolvedOutputPath
}