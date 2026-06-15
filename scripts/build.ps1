param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$source = Join-Path $repoRoot "src\noWPS.cs"
$manifest = Join-Path $repoRoot "src\noWPS.exe.manifest"
$dist = Join-Path $repoRoot "dist"
$output = Join-Path $dist "noWPS.exe"
$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"

if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}

if (-not (Test-Path $csc)) {
    throw "csc.exe was not found. Install or enable .NET Framework 4.x."
}

New-Item -ItemType Directory -Force -Path $dist | Out-Null

& $csc `
    /nologo `
    /target:winexe `
    /optimize+ `
    /out:$output `
    /win32manifest:$manifest `
    /r:System.Drawing.dll `
    /r:System.IO.Compression.dll `
    /r:System.IO.Compression.FileSystem.dll `
    /r:System.Xml.Linq.dll `
    /r:System.Windows.Forms.dll `
    $source

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$hash = Get-FileHash -Algorithm SHA256 $output
Write-Host "Built $output"
Write-Host "SHA-256: $($hash.Hash)"
