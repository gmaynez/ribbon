[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^1\.(0|[1-9]\d*)\.(0|[1-9]\d*)(-[0-9A-Za-z]+([.-][0-9A-Za-z]+)*)?$')]
    [string] $Version,

    [string] $CertificateThumbprint,

    [switch] $UseProjectSigningSettings,

    [ValidateSet('Release')]
    [string] $Configuration = 'Release',

    [string] $ArtifactsDirectory,

    [switch] $SkipInstaller
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if ([string]::IsNullOrWhiteSpace($ArtifactsDirectory)) {
    $ArtifactsDirectory = Join-Path $repositoryRoot 'artifacts\release'
}

$artifactsRoot = [System.IO.Path]::GetFullPath($ArtifactsDirectory)
$allowedArtifactsRoot = [System.IO.Path]::GetFullPath((Join-Path $repositoryRoot 'artifacts'))
if (-not $artifactsRoot.StartsWith($allowedArtifactsRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "ArtifactsDirectory must be a child of '$allowedArtifactsRoot'."
}

if ($Version -notmatch '^(?<numeric>1\.(0|[1-9]\d*)\.(0|[1-9]\d*))') {
    throw "Version '$Version' does not contain a valid numeric version."
}
$assemblyVersion = "$($Matches.numeric).0"

if (-not $UseProjectSigningSettings) {
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        throw 'CertificateThumbprint is required unless UseProjectSigningSettings is specified.'
    }

    $certificatePath = "Cert:\CurrentUser\My\$CertificateThumbprint"
    if (-not (Test-Path -LiteralPath $certificatePath)) {
        throw "Signing certificate '$CertificateThumbprint' was not found in Cert:\CurrentUser\My."
    }
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw "Visual Studio locator was not found at '$vswhere'."
}

# Probe every Visual Studio instance instead of trusting -latest: a Build Tools
# instance without the Office (VSTO) targets resolves first on machines that
# also have a full IDE, and cannot build this solution. SDK-style projects
# resolve Microsoft.NET.Sdk through the standalone dotnet resolvers, so only
# the VSTO targets discriminate here.
$msbuild = $null
$msbuildCandidates = @(& $vswhere -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe')
foreach ($candidate in $msbuildCandidates) {
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
    $msbuildRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $candidate))
    $hasVstoTargets = [bool](Get-ChildItem -LiteralPath (Join-Path $msbuildRoot 'Microsoft\VisualStudio') -Filter 'Microsoft.VisualStudio.Tools.Office.targets' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($hasVstoTargets) {
        $msbuild = $candidate
        break
    }
}
if ([string]::IsNullOrWhiteSpace($msbuild)) {
    throw 'No Visual Studio instance provides the Office (VSTO) build targets. Install Visual Studio with the Office/SharePoint development workload (VSTO tools).'
}

$buildProperties = @(
    "/p:Configuration=$Configuration",
    "/p:Version=$Version",
    "/p:AssemblyVersion=$assemblyVersion",
    "/p:FileVersion=$assemblyVersion",
    "/p:InformationalVersion=$Version",
    '/p:IncludeSourceRevisionInInformationalVersion=false',
    "/p:ApplicationVersion=$assemblyVersion",
    '/p:DefineConstants=VSTO40%3BTRACE%3BUSEOFFICEINTEROP'
)
if (-not $UseProjectSigningSettings) {
    $buildProperties += @(
        '/p:SignManifests=true',
        '/p:ManifestKeyFile=',
        "/p:ManifestCertificateThumbprint=$CertificateThumbprint"
    )
}

Push-Location $repositoryRoot
try {
    & $msbuild 'Ribbon.slnx' '/restore' '/t:Build' '/m' '/verbosity:minimal' @buildProperties
    if ($LASTEXITCODE -ne 0) { throw 'Release build failed.' }

    & dotnet test 'Ribbon.Broker.Tests\Ribbon.Broker.Tests.csproj' '--configuration' $Configuration '--no-build' '--no-restore'
    if ($LASTEXITCODE -ne 0) { throw 'Broker tests failed.' }
}
finally {
    Pop-Location
}

if (Test-Path -LiteralPath $artifactsRoot) {
    Remove-Item -LiteralPath $artifactsRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $artifactsRoot | Out-Null

$requiredBrokerFiles = @(
    'Ribbon.Broker.exe',
    'Ribbon.Broker.dll',
    'Ribbon.Broker.deps.json',
    'Ribbon.Broker.runtimeconfig.json'
)
$hosts = @('Grid', 'Quill', 'Deck', 'Post')
$contractsAssembly = Join-Path $repositoryRoot "Ribbon.Contracts\bin\$Configuration\netstandard2.0\Ribbon.Contracts.dll"
if (-not (Test-Path -LiteralPath $contractsAssembly -PathType Leaf)) {
    throw "Expected release file '$contractsAssembly' was not produced."
}
$reportedProductVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($contractsAssembly).ProductVersion
if ($reportedProductVersion -ne $Version) {
    throw "Ribbon.Contracts reports product version '$reportedProductVersion' instead of '$Version'."
}

foreach ($hostName in $hosts) {
    $outputDirectory = Join-Path $repositoryRoot "$hostName\bin\$Configuration"
    $deploymentManifest = Join-Path $outputDirectory "$hostName.vsto"
    $applicationManifest = Join-Path $outputDirectory "$hostName.dll.manifest"
    $requiredFiles = @(
        (Join-Path $outputDirectory "$hostName.dll"),
        $deploymentManifest,
        $applicationManifest
    ) + ($requiredBrokerFiles | ForEach-Object { Join-Path $outputDirectory $_ })

    foreach ($requiredFile in $requiredFiles) {
        if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
            throw "Expected release file '$requiredFile' was not produced."
        }
    }

    $applicationManifestContent = Get-Content -LiteralPath $applicationManifest -Raw
    $deploymentManifestContent = Get-Content -LiteralPath $deploymentManifest -Raw
    foreach ($brokerFile in $requiredBrokerFiles) {
        if ($applicationManifestContent -notmatch [regex]::Escape($brokerFile)) {
            throw "$hostName application manifest does not include '$brokerFile'."
        }
    }

    foreach ($manifestContent in @($deploymentManifestContent, $applicationManifestContent)) {
        if ($manifestContent -notmatch '<Signature\b') {
            throw "$hostName produced an unsigned deployment or application manifest."
        }

        $identityMatch = [regex]::Match(
            $manifestContent,
            '<(?:[A-Za-z0-9_.-]+:)?assemblyIdentity\s+name="[^"]+"\s+version="([^"]+)"')
        if (-not $identityMatch.Success -or $identityMatch.Groups[1].Value -ne $assemblyVersion) {
            throw "$hostName manifest identity does not use version '$assemblyVersion'."
        }

        if (-not $UseProjectSigningSettings) {
            $certificateMatch = [regex]::Match($manifestContent, '<X509Certificate>([^<]+)</X509Certificate>')
            if (-not $certificateMatch.Success) {
                throw "$hostName manifest does not embed its signing certificate."
            }
            $manifestCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                [Convert]::FromBase64String($certificateMatch.Groups[1].Value))
            try {
                if ($manifestCertificate.Thumbprint -ne $CertificateThumbprint) {
                    throw "$hostName manifest was not signed by certificate '$CertificateThumbprint'."
                }
            }
            finally {
                $manifestCertificate.Dispose()
            }
        }
    }

    $stagingDirectory = Join-Path $artifactsRoot $hostName
    New-Item -ItemType Directory -Path $stagingDirectory | Out-Null
    Get-ChildItem -LiteralPath $outputDirectory -File |
        Where-Object Extension -ne '.pdb' |
        Copy-Item -Destination $stagingDirectory
    Set-Content -LiteralPath (Join-Path $stagingDirectory 'VERSION.txt') -Value $Version -Encoding utf8NoBOM

    $archivePath = Join-Path $artifactsRoot "Ribbon-$hostName-v$Version.zip"
    Compress-Archive -Path (Join-Path $stagingDirectory '*') -DestinationPath $archivePath -CompressionLevel Optimal
    Remove-Item -LiteralPath $stagingDirectory -Recurse -Force
}

if (-not $SkipInstaller) {
    $installerRoot = Join-Path $artifactsRoot 'installer'
    New-Item -ItemType Directory -Path $installerRoot | Out-Null

    foreach ($hostName in $hosts) {
        $outputDirectory = Join-Path $repositoryRoot "$hostName\bin\$Configuration"
        $hostStaging = Join-Path $installerRoot $hostName
        New-Item -ItemType Directory -Path $hostStaging | Out-Null
        Get-ChildItem -LiteralPath $outputDirectory -File |
            Where-Object Extension -ne '.pdb' |
            Copy-Item -Destination $hostStaging
    }

    $brokerPublish = Join-Path $installerRoot 'Broker'
    & dotnet publish (Join-Path $repositoryRoot 'Ribbon.Broker\Ribbon.Broker.csproj') `
        '--configuration' $Configuration `
        '--runtime' 'win-x64' `
        '--self-contained' 'true' `
        '--output' $brokerPublish `
        "/p:Version=$Version" `
        "/p:AssemblyVersion=$assemblyVersion" `
        "/p:FileVersion=$assemblyVersion" `
        "/p:InformationalVersion=$Version" `
        '/p:IncludeSourceRevisionInInformationalVersion=false' `
        '/p:CopyOutputSymbolsToPublishDirectory=false'
    if ($LASTEXITCODE -ne 0) { throw 'Self-contained broker publish failed.' }

    Get-ChildItem -LiteralPath $brokerPublish -Filter '*.pdb' -File -Recurse |
        Remove-Item -Force

    $iscc = @(
        (Join-Path $env:ProgramFiles 'Inno Setup 7\ISCC.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 7\ISCC.exe')
    ) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $iscc) {
        $isccCommand = Get-Command iscc -ErrorAction SilentlyContinue
        if ($isccCommand) { $iscc = $isccCommand.Source }
    }
    if (-not $iscc) {
        throw "Inno Setup 7 compiler was not found. Install the 64-bit edition from https://jrsoftware.org/isinfo.php or pass -SkipInstaller."
    }

    $iss = Join-Path $PSScriptRoot 'Ribbon.iss'
    & $iscc '/Qp' $iss "/DAppVersion=$Version" "/DSourceDir=$installerRoot" "/O$artifactsRoot"
    if ($LASTEXITCODE -ne 0) { throw 'Inno Setup compilation failed.' }

    Remove-Item -LiteralPath $installerRoot -Recurse -Force
}

$checksumLines = Get-ChildItem -LiteralPath $artifactsRoot -File |
    Where-Object { $_.Name -ne 'SHA256SUMS.txt' } |
    Sort-Object Name |
    ForEach-Object {
        $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        "$hash  $($_.Name)"
    }
Set-Content -LiteralPath (Join-Path $artifactsRoot 'SHA256SUMS.txt') -Value $checksumLines -Encoding utf8NoBOM

Write-Host "Release artifacts are available in '$artifactsRoot'."
