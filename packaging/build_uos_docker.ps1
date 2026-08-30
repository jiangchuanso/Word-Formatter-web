param(
    [ValidateSet("aarch64", "x86_64")]
    [string]$Arch = "x86_64"
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VersionText = Get-Content (Join-Path $Root "wfp_version.py") -Raw
$VersionMatch = [regex]::Match($VersionText, '__version__\s*=\s*"([^"]+)"')
if (-not $VersionMatch.Success) {
    throw "Unable to read version from wfp_version.py"
}
$Version = $VersionMatch.Groups[1].Value

if ($Arch -eq "aarch64") {
    $Platform = "linux/arm64"
    $TargetArch = "arm64"
} else {
    $Platform = "linux/amd64"
    $TargetArch = "amd64"
}

$Image = if ($env:WFP_UOS_IMAGE) {
    $env:WFP_UOS_IMAGE
} else {
    "word-formatter-pro-uos-${TargetArch}:${Version}"
}
$Raw = "release/Word-Formatter-Pro.v${Version}.UOS-V20.${Arch}"
$AppImage = "${Raw}.AppImage"

$BuildArgs = @(
    "build", "--platform", $Platform,
    "--build-arg", "TARGETARCH=${TargetArch}"
)

# Set these environment variables when Docker Desktop must use a local UOS
# base image or a reachable Conda mirror on a restricted Windows network.
if ($env:WFP_UOS_BASE_IMAGE) {
    $BuildArgs += @("--build-arg", "UOS_BASE_IMAGE=$($env:WFP_UOS_BASE_IMAGE)")
}
if ($env:WFP_CONDA_CHANNEL) {
    $BuildArgs += @("--build-arg", "CONDA_CHANNEL=$($env:WFP_CONDA_CHANNEL)")
}
$BuildArgs += @(
    "--tag", $Image,
    "--file", (Join-Path $Root "packaging/Dockerfile.uos"),
    $Root
)

& docker @BuildArgs
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& docker run --rm `
    --platform $Platform `
    --volume "${Root}:/workspace" `
    --workdir /workspace `
    $Image `
    python packaging/build_release.py uos `
    --arch $Arch `
    --reuse-venv `
    --also-raw `
    --appimagetool /usr/local/bin/appimagetool
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$RunArgs = @(
    "run", "--rm", "--platform", $Platform,
    "--volume", "${Root}:/workspace:ro",
    "--workdir", "/workspace",
    "--env", "APPIMAGE_EXTRACT_AND_RUN=1",
    $Image
)

$RawVersion = (& docker @RunArgs $Raw "--version").Trim()
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
$AppImageVersion = (& docker @RunArgs $AppImage "--version").Trim()
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
if ($RawVersion -ne $Version -or $AppImageVersion -ne $Version) {
    throw "Built artifact version check failed"
}

$Validation = 'set -euo pipefail; test "$(uname -m)" = ''{0}''; ' +
    'file ''{1}'' ''{2}''; ''{1}'' --test; sha256sum ''{1}'' ''{2}'''
$Validation = $Validation -f $Arch, $Raw, $AppImage
& docker @RunArgs bash -lc $Validation
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "UOS ${Arch} artifacts:"
Write-Host "  $(Join-Path $Root $Raw)"
Write-Host "  $(Join-Path $Root $AppImage)"
