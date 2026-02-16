param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Url,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Output,
    [string]$UserDataDir = "",
    [int]$WaitSec = 120,
    [int]$ViewportWidth = 1200,
    [int]$WarmupSec = 0,
    [switch]$SecondPass,
    [int]$SecondPassWaitSec = 0,
    [int]$MaxDepth = 1,
    [int]$MaxPages = 30,
    [string]$BaseUrl = "",
    [int]$CdpRenderWaitSec = 15
)

$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edgePath)) { $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }

$profileDir = if ($UserDataDir -and (Test-Path $UserDataDir)) { $UserDataDir } else { Join-Path $env:TEMP "edge-remote-profile" }
if (-not (Test-Path $profileDir)) { New-Item -Path $profileDir -ItemType Directory | Out-Null }

$outputDir = $Output
if ([IO.Path]::GetExtension($outputDir) -ieq '.pdf') {
    Write-Host "Output expects a directory; using parent directory of provided file path."
    $outputDir = Split-Path $outputDir -Parent
}
if (-not (Test-Path $outputDir)) { New-Item -Path $outputDir -ItemType Directory | Out-Null }

if ($WaitSec -lt 1) { $WaitSec = 1 }
if ($ViewportWidth -lt 800) { $ViewportWidth = 800 }
if ($WarmupSec -lt 0) { $WarmupSec = 0 }
if ($SecondPassWaitSec -lt 0) { $SecondPassWaitSec = 0 }
if ($MaxDepth -lt 0) { $MaxDepth = 0 }
if ($MaxPages -lt 1) { $MaxPages = 1 }
if ($CdpRenderWaitSec -lt 0) { $CdpRenderWaitSec = 0 }

$effectiveWaitMs = $WaitSec * 1000
$a4Ratio = [Math]::Sqrt(2)
$viewportHeight = [int][Math]::Round($ViewportWidth * $a4Ratio)
$windowSize = "{0},{1}" -f $ViewportWidth, $viewportHeight

function Add-TrailingSlash {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value.EndsWith('/')) { return $Value }
    return "$Value/"
}

function Get-DefaultBasePrefix {
    param([string]$TargetUrl)

    $uri = [URI]$TargetUrl
    $segments = @($uri.Segments)
    $authority = Add-TrailingSlash ($uri.GetLeftPart([System.UriPartial]::Authority))

    if ($segments.Length -ge 3 -and ($segments[1] -ieq 'sites/' -or $segments[1] -ieq 'teams/')) {
        return Add-TrailingSlash "$authority$($segments[1])$($segments[2])"
    }

    if ($segments.Length -ge 2) {
        return Add-TrailingSlash "$authority$($segments[1])"
    }

    return Add-TrailingSlash "$authority/"
}

function Test-InScopeUrl {
    param(
        [string]$TargetUrl,
        [string]$BasePrefix
    )

    try {
        $targetUri = [URI]$TargetUrl
        $baseUri = [URI]$BasePrefix

        if ($targetUri.Scheme -ne $baseUri.Scheme) { return $false }
        if ($targetUri.Host -ne $baseUri.Host) { return $false }
        if ($targetUri.Port -ne $baseUri.Port) { return $false }

        $targetPath = Add-TrailingSlash $targetUri.AbsolutePath
        $basePath = Add-TrailingSlash $baseUri.AbsolutePath
        return $targetPath.StartsWith($basePath, [System.StringComparison]::OrdinalIgnoreCase)
    } catch {
        return $false
    }
}

function Get-ShortHash {
    param([string]$InputText)

    if ([string]::IsNullOrEmpty($InputText)) { return "" }

    $sha = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($InputText)
        $hashBytes = $sha.ComputeHash($bytes)
        $hex = [BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
        return $hex.Substring(0, 8)
    } finally {
        $sha.Dispose()
    }
}

function Get-SafeBaseNameFromUrl {
    param([string]$TargetUrl)

    try {
        $uri = [URI]$TargetUrl
        $path = [URI]::UnescapeDataString($uri.AbsolutePath)
        $trimmedPath = $path.Trim('/')

        if ([string]::IsNullOrWhiteSpace($trimmedPath)) {
            $candidate = $uri.Host
        } else {
            $candidate = $trimmedPath -replace '/', '_'
            $candidate = [IO.Path]::GetFileNameWithoutExtension($candidate)
            if ([string]::IsNullOrWhiteSpace($candidate)) {
                $candidate = $trimmedPath -replace '/', '_'
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($uri.Query)) {
            $candidate = "$candidate`_$(Get-ShortHash $uri.Query)"
        }
    } catch {
        $candidate = 'output'
    }

    if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = 'output' }

    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $safe = -join ($candidate.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '_' } else { $_ } })
    $safe = $safe.Trim('.',' ')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'output' }

    return $safe
}

function Get-UniquePdfPath {
    param(
        [string]$OutputDirectory,
        [string]$TargetUrl
    )

    $baseName = Get-SafeBaseNameFromUrl $TargetUrl
    $path = Join-Path $OutputDirectory ("{0}.pdf" -f $baseName)
    if (-not (Test-Path $path)) { return $path }

    $urlHash = Get-ShortHash $TargetUrl
    if ([string]::IsNullOrWhiteSpace($urlHash)) { $urlHash = (Get-Date -Format 'yyyyMMddHHmmss') }
    return Join-Path $OutputDirectory ("{0}_{1}.pdf" -f $baseName, $urlHash)
}

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
        [string]$WindowSize = '1920,1080'
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

function Get-PageDomHtml {
    param(
        [string]$EdgePath,
        [string]$ProfileDir,
        [string]$TargetUrl,
        [int]$BudgetMs,
        [string]$WindowSize
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
        '--dump-dom',
        $TargetUrl
    )

    $stdoutPath = Join-Path $env:TEMP ("sp2pdf-dom-{0}.txt" -f [Guid]::NewGuid().ToString('N'))
    $stderrPath = Join-Path $env:TEMP ("sp2pdf-dom-{0}.err.txt" -f [Guid]::NewGuid().ToString('N'))

    try {
        $proc = Start-Process -FilePath $EdgePath -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $timeoutMs = [Math]::Max(5000, $BudgetMs + 5000)

        if (-not $proc.WaitForExit($timeoutMs)) {
            Write-Host "Warning: DOM dump timed out after $timeoutMs ms for $TargetUrl. Skipping link discovery on this page."
            try { $proc.Kill() } catch {}
            return ""
        }

        if (Test-Path $stdoutPath) {
            return Get-Content -Path $stdoutPath -Raw -ErrorAction SilentlyContinue
        }

        return ""
    } catch {
        Write-Host "Warning: Could not dump DOM for $TargetUrl. $_"
        return ""
    } finally {
        if (Test-Path $stdoutPath) { Remove-Item -Path $stdoutPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path $stderrPath) { Remove-Item -Path $stderrPath -Force -ErrorAction SilentlyContinue }
    }
}

function Get-FreeTcpPort {
    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    try {
        $listener.Start()
        return ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
    } finally {
        $listener.Stop()
    }
}

function Receive-CdpMessage {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [int]$TimeoutMs
    )

    $buffer = New-Object byte[] 8192
    $segment = [ArraySegment[byte]]::new($buffer)
    $builder = New-Object System.Text.StringBuilder
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        $remaining = [Math]::Max(1, $TimeoutMs - [int]$sw.ElapsedMilliseconds)
        $cts = [System.Threading.CancellationTokenSource]::new($remaining)

        try {
            $recvTask = $Socket.ReceiveAsync($segment, $cts.Token)
            $recvTask.Wait()
            $recv = $recvTask.Result
        } catch {
            $cts.Dispose()
            return $null
        }

        $cts.Dispose()

        if ($recv.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
            return $null
        }

        if ($recv.Count -gt 0) {
            $chunk = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $recv.Count)
            $null = $builder.Append($chunk)
        }

        if ($recv.EndOfMessage) {
            return $builder.ToString()
        }
    }

    return $null
}

function Invoke-CdpCommand {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [int]$Id,
        [string]$Method,
        [object]$Params,
        [int]$TimeoutMs = 30000
    )

    $payload = @{ id = $Id; method = $Method }
    if ($null -ne $Params) { $payload.params = $Params }

    $json = $payload | ConvertTo-Json -Compress -Depth 20
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $segment = [ArraySegment[byte]]::new($bytes)

    $Socket.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [System.Threading.CancellationToken]::None).Wait()

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        $remaining = [Math]::Max(1, $TimeoutMs - [int]$sw.ElapsedMilliseconds)
        $raw = Receive-CdpMessage -Socket $Socket -TimeoutMs $remaining
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }

        try {
            $msg = $raw | ConvertFrom-Json -Depth 30
        } catch {
            continue
        }

        if ($null -ne $msg.id -and [int]$msg.id -eq $Id) {
            return $msg
        }
    }

    return $null
}

function Wait-CdpEvent {
    param(
        [System.Net.WebSockets.ClientWebSocket]$Socket,
        [string]$EventName,
        [int]$TimeoutMs
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
        $remaining = [Math]::Max(1, $TimeoutMs - [int]$sw.ElapsedMilliseconds)
        $raw = Receive-CdpMessage -Socket $Socket -TimeoutMs $remaining
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }

        try {
            $msg = $raw | ConvertFrom-Json -Depth 30
        } catch {
            continue
        }

        if ($msg.method -eq $EventName) {
            return $msg
        }
    }

    return $null
}

function Get-LinkedAspxUrlsViaCdp {
    param(
        [string]$EdgePath,
        [string]$ProfileDir,
        [string]$TargetUrl,
        [string]$BasePrefix,
        [int]$BudgetMs,
        [string]$WindowSize,
        [int]$RenderWaitSec,
        [string]$OutputPath = ''
    )

    $result = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $printed = $false
    $printError = ''
    $debugPort = Get-FreeTcpPort
    $edgeProc = $null
    $socket = $null

    try {
        $edgeArgs = @(
            '--disable-gpu',
            "--user-data-dir=`"$ProfileDir`"",
            '--no-first-run',
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-lazy-image-loading',
            '--disable-lazy-frame-loading',
            '--run-all-compositor-stages-before-draw',
            "--window-size=$WindowSize",
            "--remote-debugging-port=$debugPort",
            'about:blank'
        )

        $edgeArgs = @('--headless=new') + $edgeArgs
        $edgeProc = Start-Process -FilePath $EdgePath -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden

        $endpointBase = "http://127.0.0.1:$debugPort"
        $ready = $false
        for ($i = 0; $i -lt 100; $i++) {
            try {
                $null = Invoke-RestMethod -Uri "$endpointBase/json/version" -TimeoutSec 1
                $ready = $true
                break
            } catch {
                Start-Sleep -Milliseconds 100
            }
        }
        if (-not $ready) {
            Write-Host "Warning: CDP endpoint did not become ready for $TargetUrl"
            return @()
        }

        $escapedUrl = [Uri]::EscapeDataString($TargetUrl)
        $targetInfo = $null
        try {
            $targetInfo = Invoke-RestMethod -Method Put -Uri "$endpointBase/json/new?$escapedUrl" -TimeoutSec 5
        } catch {
            try {
                $targetInfo = Invoke-RestMethod -Method Get -Uri "$endpointBase/json/new?$escapedUrl" -TimeoutSec 5
            } catch {
                Write-Host "Warning: CDP could not create target for $TargetUrl"
                return @()
            }
        }

        $wsUrl = [string]$targetInfo.webSocketDebuggerUrl
        if ([string]::IsNullOrWhiteSpace($wsUrl)) {
            Write-Host "Warning: CDP target did not provide webSocketDebuggerUrl for $TargetUrl"
            return @()
        }

        $socket = [System.Net.WebSockets.ClientWebSocket]::new()
        $socket.ConnectAsync([Uri]$wsUrl, [System.Threading.CancellationToken]::None).Wait()

        $id = 1
        $null = Invoke-CdpCommand -Socket $socket -Id $id -Method 'Page.enable' -Params @{} -TimeoutMs 10000
        $id++
        $null = Invoke-CdpCommand -Socket $socket -Id $id -Method 'Runtime.enable' -Params @{} -TimeoutMs 10000
        $id++
        $null = Invoke-CdpCommand -Socket $socket -Id $id -Method 'Page.navigate' -Params @{ url = $TargetUrl } -TimeoutMs 15000
        $id++

        $loadTimeoutMs = [Math]::Max(5000, $BudgetMs + 5000)
        $loadEvent = Wait-CdpEvent -Socket $socket -EventName 'Page.loadEventFired' -TimeoutMs $loadTimeoutMs
        if ($null -eq $loadEvent) {
            Write-Host "Warning: CDP load event timeout for $TargetUrl"
            return @()
        }

        if ($RenderWaitSec -gt 0) {
            Start-Sleep -Seconds $RenderWaitSec
        }

                $extractExpression = @'
(async () => {
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    const getScrollableElements = () => {
        const candidates = Array.from(document.querySelectorAll('div,section,main,article,[role="list"],[class*="List"],[class*="list"]'));
        const scrollables = candidates.filter(el => {
            const style = window.getComputedStyle(el);
            const overflowY = style.overflowY || '';
            const canScroll = el.scrollHeight > (el.clientHeight + 50);
            const allowsScroll = /auto|scroll/i.test(overflowY);
            return canScroll && allowsScroll;
        });

        scrollables.sort((a, b) => (b.scrollHeight - b.clientHeight) - (a.scrollHeight - a.clientHeight));
        return scrollables.slice(0, 6);
    };

    const scrollDiagnostics = {
        windowMoves: 0,
        containerMoves: 0,
        maxWindowY: 0,
        containerCount: 0
    };

    const containers = getScrollableElements();
    scrollDiagnostics.containerCount = containers.length;

    for (let i = 0; i < 20; i++) {
        const beforeWindowY = window.scrollY || window.pageYOffset || 0;
        const documentHeight = Math.max(
            document.body ? document.body.scrollHeight : 0,
            document.documentElement ? document.documentElement.scrollHeight : 0
        );
        window.scrollTo(0, documentHeight);

        const afterWindowY = window.scrollY || window.pageYOffset || 0;
        if (afterWindowY > beforeWindowY) {
            scrollDiagnostics.windowMoves++;
            scrollDiagnostics.maxWindowY = Math.max(scrollDiagnostics.maxWindowY, afterWindowY);
        }

        for (const el of containers) {
            const beforeTop = el.scrollTop;
            el.scrollTop = el.scrollHeight;
            if (el.scrollTop > beforeTop) {
                scrollDiagnostics.containerMoves++;
            }
        }

        await sleep(900);
    }

    window.scrollTo(0, 0);
    for (const el of containers) {
        el.scrollTop = 0;
    }
    await sleep(500);

  const leftNavSelector = 'nav[class*="spReactLeftNav"], nav[role="navigation"][class*="spReactLeftNav"], nav[role="navigation"][aria-label="Webbplats"]';
    const clean = (href) => {
        if (!href) return '';
        if (href.startsWith('#')) return '';
        if (/^javascript:/i.test(href)) return '';
        if (/^mailto:/i.test(href)) return '';
        return href;
    };

    const tileCandidates = [];
    for (const card of Array.from(document.querySelectorAll('[data-automation-id="tile-card"]'))) {
        const a = card.closest('a[href]');
        if (a) tileCandidates.push(clean(a.getAttribute('href') || ''));
    }
    for (const title of Array.from(document.querySelectorAll('[data-automation-id="quick-links-item-title"]'))) {
        const a = title.closest('a[href]');
        if (a) tileCandidates.push(clean(a.getAttribute('href') || ''));
    }
    for (const a of Array.from(document.querySelectorAll('a[href][class*="tileCard"]'))) {
        tileCandidates.push(clean(a.getAttribute('href') || ''));
    }

    const tileHrefs = Array.from(new Set(tileCandidates.filter(Boolean).filter(h => /\.aspx(\?|$)/i.test(h))));

    const anchors = Array.from(document.querySelectorAll('a[href]'));
    const rows = anchors
        .map(a => ({
            href: clean(a.getAttribute('href') || ''),
            inLeftNav: !!a.closest(leftNavSelector)
        }))
        .filter(r => !!r.href)
        .filter(r => /\.aspx(\?|$)/i.test(r.href));

    const uniqueRows = [];
    const seenRows = new Set();
    for (const row of rows) {
        const key = `${row.href}::${row.inLeftNav}`;
        if (!seenRows.has(key)) {
            seenRows.add(key);
            uniqueRows.push(row);
        }
    }

    return {
        scrollDiagnostics,
        tileHrefs,
        aspxAnchorRows: uniqueRows
    };
})()
'@

        $evalResponse = Invoke-CdpCommand -Socket $socket -Id $id -Method 'Runtime.evaluate' -Params @{
            expression = $extractExpression
            returnByValue = $true
            awaitPromise = $true
        } -TimeoutMs ([Math]::Max(15000, $BudgetMs))

        if ($null -eq $evalResponse -or $null -eq $evalResponse.result -or $null -eq $evalResponse.result.result) {
            return @()
        }

        $value = $evalResponse.result.result.value
        $tileRows = @()
        $rows = @()
        $scrollInfo = $null
        if ($null -ne $value) {
            if ($null -ne $value.scrollDiagnostics) {
                $scrollInfo = $value.scrollDiagnostics
            }
            if ($null -ne $value.tileHrefs) {
                if ($value.tileHrefs -is [System.Array]) { $tileRows = @($value.tileHrefs) } else { $tileRows = @($value.tileHrefs) }
            }
            if ($null -ne $value.aspxAnchorRows) {
                if ($value.aspxAnchorRows -is [System.Array]) { $rows = @($value.aspxAnchorRows) } else { $rows = @($value.aspxAnchorRows) }
            }
        }

        if ($null -ne $scrollInfo) {
            Write-Host "CDP scroll diagnostics: containers=$($scrollInfo.containerCount), windowMoves=$($scrollInfo.windowMoves), containerMoves=$($scrollInfo.containerMoves), maxWindowY=$($scrollInfo.maxWindowY)"
        }

        if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
            try {
                $printResponse = Invoke-CdpCommand -Socket $socket -Id $id -Method 'Page.printToPDF' -Params @{
                    displayHeaderFooter = $false
                    printBackground = $true
                    preferCSSPageSize = $true
                } -TimeoutMs ([Math]::Max(15000, $BudgetMs))
                $id++

                $pdfData = ''
                if ($null -ne $printResponse -and $null -ne $printResponse.result -and $null -ne $printResponse.result.data) {
                    $pdfData = [string]$printResponse.result.data
                }

                if (-not [string]::IsNullOrWhiteSpace($pdfData)) {
                    $pdfBytes = [Convert]::FromBase64String($pdfData)
                    [IO.File]::WriteAllBytes($OutputPath, $pdfBytes)
                    $printed = $true
                } else {
                    $printError = 'CDP print returned no PDF data.'
                }
            } catch {
                $printError = [string]$_
            }
        }

        $tileDetected = 0
        $tileKept = 0
        $totalAspx = 0
        $leftNavAspx = 0
        $keptAspx = 0
        $sampleRejected = New-Object System.Collections.Generic.List[string]

        foreach ($tileHref in $tileRows) {
            $href = [string]$tileHref
            if ([string]::IsNullOrWhiteSpace($href)) { continue }
            $tileDetected++
            try {
                $abs = [URI]::new([URI]$TargetUrl, $href).AbsoluteUri
                $absUri = [URI]$abs

                if (-not (Test-InScopeUrl -TargetUrl $abs -BasePrefix $BasePrefix)) { continue }
                if ([IO.Path]::GetExtension($absUri.AbsolutePath) -ine '.aspx') { continue }

                if ($result.Add($abs)) {
                    $tileKept++
                }
            } catch {
                continue
            }
        }

        foreach ($row in $rows) {
            $href = [string]$row.href
            $inLeftNav = [bool]$row.inLeftNav
            if ([string]::IsNullOrWhiteSpace($href)) { continue }
            $totalAspx++
            if ($inLeftNav) {
                $leftNavAspx++
                continue
            }

            try {
                $abs = [URI]::new([URI]$TargetUrl, $href).AbsoluteUri
                $absUri = [URI]$abs

                if (-not (Test-InScopeUrl -TargetUrl $abs -BasePrefix $BasePrefix)) {
                    if ($sampleRejected.Count -lt 8) { $sampleRejected.Add("out-of-scope: $abs") | Out-Null }
                    continue
                }
                if ([IO.Path]::GetExtension($absUri.AbsolutePath) -ine '.aspx') {
                    if ($sampleRejected.Count -lt 8) { $sampleRejected.Add("not-aspx-path: $abs") | Out-Null }
                    continue
                }

                if ($result.Add($abs)) {
                    $keptAspx++
                }
            } catch {
                if ($sampleRejected.Count -lt 8) { $sampleRejected.Add("invalid-url: $href") | Out-Null }
                continue
            }
        }

        Write-Host "CDP tile scan: detected=$tileDetected, kept=$tileKept"
        Write-Host "CDP link scan: total .aspx anchors=$totalAspx, inside left-nav=$leftNavAspx, kept fallback candidates=$keptAspx"
        if ($keptAspx -eq 0 -and $sampleRejected.Count -gt 0) {
            Write-Host "CDP sample rejected links:"
            foreach ($line in $sampleRejected) {
                Write-Host "  - $line"
            }
        }
    } catch {
        Write-Host "Warning: CDP link discovery failed for $TargetUrl. $_"
    } finally {
        if ($null -ne $socket) {
            try { $socket.Dispose() } catch {}
        }
        if ($null -ne $edgeProc -and -not $edgeProc.HasExited) {
            try { $edgeProc.Kill() } catch {}
        }
    }

    return [PSCustomObject]@{
        Links = @($result)
        Printed = $printed
        PrintError = $printError
    }
}

function Get-LinkedAspxUrls {
    param(
        [string]$Html,
        [string]$CurrentUrl,
        [string]$BasePrefix
    )

    if ([string]::IsNullOrWhiteSpace($Html)) { return @() }

    $result = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $regexOptions = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline

    $leftNavRanges = New-Object System.Collections.Generic.List[object]
    $leftNavOpenPattern = '<nav\b[^>]*\bclass\s*=\s*[\x22\x27][^\x22\x27]*\bspReactLeftNav_[^\x22\x27]*[\x22\x27][^>]*>'
    $leftNavOpenMatches = [regex]::Matches($Html, $leftNavOpenPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($navMatch in $leftNavOpenMatches) {
        $navStart = $navMatch.Index
        $searchStart = $navMatch.Index + $navMatch.Length
        $navEndOpen = $Html.IndexOf('</nav>', $searchStart, [System.StringComparison]::OrdinalIgnoreCase)
        if ($navEndOpen -lt 0) {
            $navEnd = $navMatch.Index + $navMatch.Length
        } else {
            $navEnd = $navEndOpen + 6
        }

        $leftNavRanges.Add([PSCustomObject]@{
            Start = $navStart
            End = $navEnd
        }) | Out-Null
    }

    $anchorPattern = '<a\b([^>]*)>'
    $anchorMatches = [regex]::Matches($Html, $anchorPattern, $regexOptions)

    foreach ($anchor in $anchorMatches) {
        $anchorIndex = $anchor.Index
        $insideLeftNav = $false
        foreach ($range in $leftNavRanges) {
            if ($anchorIndex -ge $range.Start -and $anchorIndex -lt $range.End) {
                $insideLeftNav = $true
                break
            }
        }
        if ($insideLeftNav) { continue }

        $attributes = $anchor.Groups[1].Value

        $hrefMatch = [regex]::Match($attributes, '\bhref\s*=\s*[\x22\x27]([^\x22\x27#]+)[\x22\x27]', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $hrefMatch.Success) { continue }

        $href = $hrefMatch.Groups[1].Value.Trim()
        if ([string]::IsNullOrWhiteSpace($href)) { continue }
        if ($href.StartsWith('javascript:', [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        if ($href.StartsWith('mailto:', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        try {
            $abs = [URI]::new([URI]$CurrentUrl, $href).AbsoluteUri
            $absUri = [URI]$abs

            if (-not (Test-InScopeUrl -TargetUrl $abs -BasePrefix $BasePrefix)) { continue }
            if ([IO.Path]::GetExtension($absUri.AbsolutePath) -ine '.aspx') { continue }

            $null = $result.Add($abs)
        } catch {
            continue
        }
    }

    return @($result)
}

Write-Host "Starting Edge in direct headless print mode..."

$running = Get-Process -Name msedge -ErrorAction SilentlyContinue
if ($running) {
    Write-Host "Stopping existing Edge processes to ensure a fresh headless start..."
    $running | Stop-Process -Force
    Start-Sleep -Seconds 1
}

$basePrefix = if ([string]::IsNullOrWhiteSpace($BaseUrl)) { Get-DefaultBasePrefix $Url } else { Add-TrailingSlash $BaseUrl }
Write-Host "Crawl base prefix: $basePrefix"
Write-Host "Wait budget: $WaitSec sec ($effectiveWaitMs ms)"
Write-Host "Viewport (A4 portrait ratio): $windowSize"
Write-Host "Crawl depth: $MaxDepth (max pages: $MaxPages)"
Write-Host "Second pass: $($SecondPass.IsPresent)"
if ($SecondPass) {
    $secondPassBudgetSec = if ($SecondPassWaitSec -gt 0) { $SecondPassWaitSec } else { $WaitSec }
    Write-Host "Second-pass first-run budget: $secondPassBudgetSec sec"
}
Write-Host "CDP extra render wait: $CdpRenderWaitSec sec"

$queue = New-Object System.Collections.Generic.Queue[object]
$visited = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
$exported = New-Object System.Collections.Generic.List[string]

$queue.Enqueue([PSCustomObject]@{ Url = $Url; Depth = 0; Warmup = ($WarmupSec -gt 0) })

while ($queue.Count -gt 0) {
    if ($visited.Count -ge $MaxPages) {
        Write-Host "Reached MaxPages=$MaxPages, stopping crawl."
        break
    }

    $item = $queue.Dequeue()
    $currentUrl = [string]$item.Url
    $depth = [int]$item.Depth
    $doWarmup = [bool]$item.Warmup

    if ($visited.Contains($currentUrl)) { continue }

    if (-not (Test-InScopeUrl -TargetUrl $currentUrl -BasePrefix $basePrefix)) {
        Write-Host "Skipping out-of-scope URL: $currentUrl"
        continue
    }

    $null = $visited.Add($currentUrl)
    Write-Host "Exporting (depth $depth): $currentUrl"

    $pdfPath = Get-UniquePdfPath -OutputDirectory $outputDir -TargetUrl $currentUrl

    if ($doWarmup) {
        $warmupMs = $WarmupSec * 1000
        $warmupPath = [IO.Path]::ChangeExtension($pdfPath, $null) + '.warmup.pdf'
        Write-Host "Running warm-up pass for $WarmupSec sec..."
        $null = Invoke-DirectHeadlessPrint -EdgePath $edgePath -ProfileDir $profileDir -OutputPath $warmupPath -TargetUrl $currentUrl -BudgetMs $warmupMs -WindowSize $windowSize
        if (Test-Path $warmupPath) { Remove-Item -Path $warmupPath -Force -ErrorAction SilentlyContinue }
    }

    if ($SecondPass) {
        $firstRunBudgetSec = if ($SecondPassWaitSec -gt 0) { $SecondPassWaitSec } else { $WaitSec }
        $firstRunBudgetMs = $firstRunBudgetSec * 1000
        $firstRunPath = [IO.Path]::ChangeExtension($pdfPath, $null) + '.firstpass.pdf'
        Write-Host "Running second-pass first render for $firstRunBudgetSec sec..."
        $null = Invoke-DirectHeadlessPrint -EdgePath $edgePath -ProfileDir $profileDir -OutputPath $firstRunPath -TargetUrl $currentUrl -BudgetMs $firstRunBudgetMs -WindowSize $windowSize
        if (Test-Path $firstRunPath) { Remove-Item -Path $firstRunPath -Force -ErrorAction SilentlyContinue }
    }

    $links = @()
    $printSw = [System.Diagnostics.Stopwatch]::StartNew()
    $exitCode = 0
    $cdpResult = Get-LinkedAspxUrlsViaCdp -EdgePath $edgePath -ProfileDir $profileDir -TargetUrl $currentUrl -BasePrefix $basePrefix -BudgetMs $effectiveWaitMs -WindowSize $windowSize -RenderWaitSec $CdpRenderWaitSec -OutputPath $pdfPath
    if ($null -ne $cdpResult -and $null -ne $cdpResult.Links) {
        if ($cdpResult.Links -is [System.Array]) { $links = @($cdpResult.Links) } else { $links = @($cdpResult.Links) }
    }

    if (-not (Test-Path $pdfPath)) {
        $exitCode = 1
        if ($null -ne $cdpResult -and -not [string]::IsNullOrWhiteSpace([string]$cdpResult.PrintError)) {
            Write-Host "Warning: CDP print failed for $currentUrl. $($cdpResult.PrintError)"
        } else {
            Write-Host "Warning: CDP print did not produce output for $currentUrl"
        }
        continue
    }
    $printSw.Stop()

    $pdfInfo = Get-Item $pdfPath
    $pageCount = Get-PdfPageCount $pdfPath
    Write-Host "Saved: $pdfPath (exit=$exitCode, $([Math]::Round($printSw.Elapsed.TotalSeconds,1))s, bytes=$($pdfInfo.Length), pages≈$pageCount)"
    $null = $exported.Add($pdfPath)

    if ($depth -lt $MaxDepth) {
        Write-Host "Discovering linked .aspx URLs from: $currentUrl"
        foreach ($link in $links) {
            if (-not $visited.Contains($link)) {
                $queue.Enqueue([PSCustomObject]@{ Url = $link; Depth = ($depth + 1); Warmup = $false })
            }
        }

        Write-Host "Queued $($links.Count) in-scope .aspx links from depth $depth page."
    }
}

Write-Host "Done. Exported $($exported.Count) PDF file(s) to $outputDir"
