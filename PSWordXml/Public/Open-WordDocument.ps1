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
    $ResolvedZipInputPath = Join-Path -Path (Split-Path -Path $ResolvedInputPath) -ChildPath ((Get-ChildItem -Path $ResolvedInputPath).BaseName + '.zip')

    # have to do this for Expand-Archive to work
    Copy-Item -Path $ResolvedInputPath -Destination $ResolvedZipInputPath -Force:$Force
    Expand-Archive -Path $ResolvedZipInputPath -DestinationPath $ResolvedOutputPath -Force:$Force
    Remove-Item -Path $ResolvedZipInputPath -Force:$Force

    $global:OpenWordDocument = $ResolvedOutputPath
}