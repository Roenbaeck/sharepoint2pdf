param(
    [string]$Url = "https://amfpension.sharepoint.com/sites/DWplattformsteam-STM/SitePages/Kundvärdesmodellen.aspx",
    [string]$Output = "C:\KVM-PDF",
    [string]$UserDataDir = "",
    [int]$WaitSec = 120,
    [int]$ViewportWidth = 1200,
    [int]$WarmupSec = 0
)

$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edgePath)) { $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }

$profileDir = if ($UserDataDir -and (Test-Path $UserDataDir)) { $UserDataDir } else { Join-Path $env:TEMP "edge-remote-profile" }
if (-not (Test-Path $profileDir)) { New-Item -Path $profileDir -ItemType Directory | Out-Null }

$outputDir = $Output
if ([string]::IsNullOrWhiteSpace($outputDir)) { $outputDir = "C:\KVM-PDF" }
if ([IO.Path]::GetExtension($outputDir) -ieq '.pdf') {
    Write-Host "Output now expects a directory; using parent directory of provided file path."
    $outputDir = Split-Path $outputDir -Parent
}
if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }

try {
    $urlObj = [URI]$url
    $urlLeaf = [IO.Path]::GetFileName($urlObj.AbsolutePath)
    if (-not [string]::IsNullOrWhiteSpace($urlLeaf)) {
        $urlLeaf = [URI]::UnescapeDataString($urlLeaf)
    }
    if ([string]::IsNullOrWhiteSpace($urlLeaf)) { $urlLeaf = $urlObj.Host }
} catch {
    $urlLeaf = "output"
}

$urlBase = [IO.Path]::GetFileNameWithoutExtension($urlLeaf)
if ([string]::IsNullOrWhiteSpace($urlBase)) { $urlBase = "output" }
$invalidFileChars = [IO.Path]::GetInvalidFileNameChars()
$safeBase = -join ($urlBase.ToCharArray() | ForEach-Object { if ($invalidFileChars -contains $_) { '_' } else { $_ } })
$safeBase = $safeBase.Trim('.',' ')
if ([string]::IsNullOrWhiteSpace($safeBase)) { $safeBase = "output" }
$outPath = Join-Path $outputDir ("{0}.pdf" -f $safeBase)

if ($WaitSec -lt 1) { $WaitSec = 1 }
$effectiveWaitMs = $WaitSec * 1000
if ($ViewportWidth -lt 800) { $ViewportWidth = 800 }
if ($WarmupSec -lt 0) { $WarmupSec = 0 }

$a4Ratio = [Math]::Sqrt(2)
$viewportHeight = [int][Math]::Round($ViewportWidth * $a4Ratio)
$windowSize = "{0},{1}" -f $ViewportWidth, $viewportHeight

function Get-PdfPageCount {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    try {
        $bytes = [IO.File]::ReadAllBytes($Path)
        $text = [Text.Encoding]::ASCII.GetString($bytes)
        return ([regex]::Matches($text, '/Type\s*/Page\b')).Count
    } catch {
        return 0
    }
}

function Invoke-DirectHeadlessPrint {
    param(
        [string]$EdgePath,
        [string]$ProfileDir,
        [string]$OutputPath,
        [string]$TargetUrl,
        [int]$BudgetMs,
        [string]$WindowSize = "1920,1080"
    )

    $edgeArgs = @(
        '--headless=new',
        '--disable-gpu',
        "--user-data-dir=`"$ProfileDir`"",
        '--no-first-run',
        '--run-all-compositor-stages-before-draw',
        "--virtual-time-budget=$BudgetMs",
        "--timeout=$BudgetMs",
        "--window-size=$WindowSize",
        '--disable-lazy-image-loading',
        '--disable-lazy-frame-loading',
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--print-to-pdf-no-header',
        "--print-to-pdf=`"$OutputPath`"",
        $TargetUrl
    )

    $proc = Start-Process -FilePath $EdgePath -ArgumentList $edgeArgs -PassThru -Wait
    return $proc.ExitCode
}

Write-Host "Starting Edge in direct headless print mode..."

$running = Get-Process -Name msedge -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "Stopping existing Edge processes to ensure a fresh headless start..."
    $running | Stop-Process -Force
    Start-Sleep -Seconds 1
}

Write-Host "Wait budget: $WaitSec sec ($effectiveWaitMs ms)"
Write-Host "Viewport (A4 portrait ratio): $windowSize"

if ($WarmupSec -gt 0) {
    $warmupMs = $WarmupSec * 1000
    $warmupPath = [IO.Path]::ChangeExtension($outPath, $null) + ".warmup.pdf"
    Write-Host "Running warm-up pass for $WarmupSec sec before final output..."
    $warmupSw = [System.Diagnostics.Stopwatch]::StartNew()
    $null = Invoke-DirectHeadlessPrint -EdgePath $edgePath -ProfileDir $profileDir -OutputPath $warmupPath -TargetUrl $url -BudgetMs $warmupMs -WindowSize $windowSize
    $warmupSw.Stop()
    if (Test-Path $warmupPath) { Remove-Item -Path $warmupPath -Force -ErrorAction SilentlyContinue }
    Write-Host "Warm-up pass finished after $([Math]::Round($warmupSw.Elapsed.TotalSeconds,1))s"
}

$printSw = [System.Diagnostics.Stopwatch]::StartNew()
$exitCode = Invoke-DirectHeadlessPrint -EdgePath $edgePath -ProfileDir $profileDir -OutputPath $outPath -TargetUrl $url -BudgetMs $effectiveWaitMs -WindowSize $windowSize
$printSw.Stop()
Write-Host "Edge headless print process exited with code $exitCode after $([Math]::Round($printSw.Elapsed.TotalSeconds,1))s"

if (-not (Test-Path $outPath)) {
    throw "Headless print did not produce output file: $outPath"
}

$pdfInfo = Get-Item $outPath
$pageCount = Get-PdfPageCount $outPath
Write-Host "PDF saved to $outPath ($($pdfInfo.Length) bytes, approx pages=$pageCount)"

if ($pdfInfo.Length -lt 1024) {
    Write-Host "Warning: PDF is very small; this can happen if the page requires an authenticated profile."
}

Write-Host "Script finished."
