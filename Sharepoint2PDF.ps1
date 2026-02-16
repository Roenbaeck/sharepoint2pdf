param(
    [switch]$Headless,
    [string]$HeadlessOutput = "C:\KVM-PDF\output_headless.pdf",
    [string]$UserDataDir = "",
    [int]$MaxDepth = 1,
    [int]$MinContentChars = 50,
    [int]$HeadlessWaitSec = 120,
    [int]$HeadlessWaitMs = 0,
    [int]$HeadlessViewportWidth = 1200,
    [double]$HeadlessSecondPassWidthFactor = 0.85,
    [switch]$HeadlessSecondPass
)

$url = "https://amfpension.sharepoint.com/sites/DWplattformsteam-STM/SitePages/Kundvärdesmodellen.aspx"

# ==========================
# Connect to Edge DevTools
# ==========================

# Detect Edge binary path
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (-not (Test-Path $edgePath)) { $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe" }
$profileDir = if ($UserDataDir -and (Test-Path $UserDataDir)) { $UserDataDir } else { Join-Path $env:TEMP "edge-remote-profile" }
if (-not (Test-Path $profileDir)) { New-Item -Path $profileDir -ItemType Directory | Out-Null }
$outPath = $HeadlessOutput
if (-not (Test-Path (Split-Path $outPath -Parent))) { New-Item -Path (Split-Path $outPath -Parent) -ItemType Directory | Out-Null }

if ($HeadlessWaitMs -gt 0) {
    Write-Host "HeadlessWaitMs is deprecated; converting ${HeadlessWaitMs}ms to seconds. Use -HeadlessWaitSec instead."
    $HeadlessWaitSec = [int][Math]::Ceiling($HeadlessWaitMs / 1000.0)
}
if ($HeadlessWaitSec -lt 1) { $HeadlessWaitSec = 1 }
$effectiveHeadlessWaitMs = $HeadlessWaitSec * 1000
if (-not $PSBoundParameters.ContainsKey('HeadlessSecondPass')) { $HeadlessSecondPass = $true }
if ($HeadlessViewportWidth -lt 800) { $HeadlessViewportWidth = 800 }
if ($HeadlessSecondPassWidthFactor -le 0 -or $HeadlessSecondPassWidthFactor -gt 1) { $HeadlessSecondPassWidthFactor = 0.85 }
$a4Ratio = [Math]::Sqrt(2)
$viewportHeight = [int][Math]::Round($HeadlessViewportWidth * $a4Ratio)
$headlessWindowSize = "{0},{1}" -f $HeadlessViewportWidth, $viewportHeight
$secondPassViewportWidth = [int][Math]::Round($HeadlessViewportWidth * $HeadlessSecondPassWidthFactor)
if ($secondPassViewportWidth -lt 800) { $secondPassViewportWidth = 800 }
$secondPassViewportHeight = [int][Math]::Round($secondPassViewportWidth * $a4Ratio)
$secondPassWindowSize = "{0},{1}" -f $secondPassViewportWidth, $secondPassViewportHeight

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

function Invoke-HeadlessDirectPrint {
    param(
        [string]$EdgePath,
        [string]$ProfileDir,
        [string]$OutputPath,
        [string]$TargetUrl,
        [int]$WaitMs,
        [string]$WindowSize = "1920,1080"
    )

    $args = @(
        '--headless=new',
        '--disable-gpu',
        "--user-data-dir=`"$ProfileDir`"",
        '--no-first-run',
        '--run-all-compositor-stages-before-draw',
        "--virtual-time-budget=$WaitMs",
        "--timeout=$WaitMs",
        "--window-size=$WindowSize",
        '--disable-lazy-image-loading',
        '--disable-lazy-frame-loading',
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--print-to-pdf-no-header',
        "--print-to-pdf=`"$OutputPath`"",
        $TargetUrl
    )

    $proc = Start-Process -FilePath $EdgePath -ArgumentList $args -PassThru -Wait
    return $proc.ExitCode
}

# Headless mode shortcut: run Edge headless print and exit
if ($Headless) {
    Write-Host "Starting Edge in direct headless print mode (no CDP)..."

    # If any msedge instances are running, stop them so our headless flags are respected
    $running = Get-Process -Name msedge -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host "Stopping existing Edge processes to ensure a fresh headless start..."
        $running | Stop-Process -Force
        Start-Sleep -Seconds 1
    }

    Write-Host "Headless wait budget: $HeadlessWaitSec sec ($effectiveHeadlessWaitMs ms)"
    Write-Host "Headless viewport (A4 portrait ratio): $headlessWindowSize"
    $printSw = [System.Diagnostics.Stopwatch]::StartNew()
    $exitCode = Invoke-HeadlessDirectPrint -EdgePath $edgePath -ProfileDir $profileDir -OutputPath $outPath -TargetUrl $url -WaitMs $effectiveHeadlessWaitMs -WindowSize $headlessWindowSize
    $printSw.Stop()
    Write-Host "Edge headless print process exited with code $exitCode after $([Math]::Round($printSw.Elapsed.TotalSeconds,1))s"

    if (-not (Test-Path $outPath)) {
        throw "Headless print did not produce output file: $outPath"
    }

    $pdfInfo = Get-Item $outPath
    $pageCount = Get-PdfPageCount $outPath
    Write-Host "PDF saved to $outPath ($($pdfInfo.Length) bytes, approx pages=$pageCount)"

    # Optional second pass with longer wait and taller viewport to capture lazy/late-loading images.
    if ($HeadlessSecondPass) {
        $retryWaitMs = [Math]::Max($effectiveHeadlessWaitMs * 2, 240000)
        $retryWaitSec = [int][Math]::Ceiling($retryWaitMs / 1000.0)
        $secondPassPath = [IO.Path]::ChangeExtension($outPath, $null) + ".pass2.pdf"
        Write-Host "Running second pass for image completeness (${retryWaitSec}s / ${retryWaitMs}ms), narrower A4 portrait viewport $secondPassWindowSize..."
        $retrySw = [System.Diagnostics.Stopwatch]::StartNew()
        $exitCode2 = Invoke-HeadlessDirectPrint -EdgePath $edgePath -ProfileDir $profileDir -OutputPath $secondPassPath -TargetUrl $url -WaitMs $retryWaitMs -WindowSize $secondPassWindowSize
        $retrySw.Stop()
        Write-Host "Retry print process exited with code $exitCode2 after $([Math]::Round($retrySw.Elapsed.TotalSeconds,1))s"

        if (Test-Path $secondPassPath) {
            $firstInfo = Get-Item $outPath
            $firstPages = Get-PdfPageCount $outPath
            $secondInfo = Get-Item $secondPassPath
            $secondPages = Get-PdfPageCount $secondPassPath

            $useSecond = $false
            if ($secondPages -gt $firstPages) { $useSecond = $true }
            elseif ($secondPages -eq $firstPages -and $secondInfo.Length -gt $firstInfo.Length) { $useSecond = $true }

            if ($useSecond) {
                Move-Item -Path $secondPassPath -Destination $outPath -Force
                $pdfInfo = Get-Item $outPath
                $pageCount = Get-PdfPageCount $outPath
                Write-Host "Selected second-pass PDF: $outPath ($($pdfInfo.Length) bytes, approx pages=$pageCount)"
            } else {
                Remove-Item -Path $secondPassPath -Force -ErrorAction SilentlyContinue
                Write-Host "Kept first-pass PDF: $outPath ($($firstInfo.Length) bytes, approx pages=$firstPages)"
                $pdfInfo = $firstInfo
                $pageCount = $firstPages
            }
        } else {
            Write-Host "Second pass did not produce an output file; keeping first pass."
        }
    }

    if ($pdfInfo.Length -lt 1024) {
        Write-Host "Warning: PDF is very small; this can happen if the page requires an authenticated profile."
    }

    Write-Host "Script finished."
    return
} else {
    # If not in headless mode, ensure an Edge instance is running with remote debugging
    $existing = Get-Process -Name msedge -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Host "Launching Edge with remote debugging..."
        Start-Process -FilePath $edgePath -ArgumentList "--remote-debugging-port=9222 --user-data-dir=`"$profileDir`" --no-first-run $url"
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Edge is already running; assuming remote debugging is enabled on port 9222."
    }
}

# Wait for DevTools to become available (use IPv4 to avoid IPv6 localhost resolution issues)
$targets = $null
$maxAttempts = 20
for ($i = 0; $i -lt $maxAttempts; $i++) {
    try {
        $targets = Invoke-RestMethod "http://127.0.0.1:9222/json/list" -TimeoutSec 2
        if ($targets) { break }
    } catch {
        Start-Sleep -Seconds 1
    }
} 
if (-not $targets) {
    Write-Host "Could not connect to Edge DevTools on 9222 after waiting. Gathering diagnostics..."
    Write-Host "Edge binary exists: $([IO.File]::Exists($edgePath)) Path: $edgePath"

    Write-Host "Processes matching msedge.exe:"
    Get-Process -Name msedge -ErrorAction SilentlyContinue | Select-Object Id, ProcessName, StartTime | Format-Table -AutoSize

    Write-Host "`nEdge command lines (may require admin privileges):"
    try {
        Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Select-Object ProcessId, CommandLine | Format-List
    } catch {
        Write-Host "Could not enumerate process command lines (requires elevated privileges)."
    }

    Write-Host "`nListening ports containing 9222 (netstat):"
    try {
        & netstat -ano | Select-String "9222" | ForEach-Object { $_.ToString() }
    } catch {
        Write-Host "netstat not available or failed."
    }

    Write-Host "`nTest-NetConnection result:"
    try {
        Test-NetConnection -ComputerName 127.0.0.1 -Port 9222 | Format-List
    } catch {
        Write-Host "Test-NetConnection failed or is not available on this system."
    }

    # If netstat found a listener, try to resolve the PID to a process
    $listenerInfo = (& netstat -ano | Select-String "9222" | Select-Object -First 1).ToString()
    if ($listenerInfo) {
        $parts = $listenerInfo -split '\s+'
        $listenerPid = $parts[-1]
        Write-Host "`nProcess listening on 9222 (pid $listenerPid):"
        try { Get-Process -Id $listenerPid -ErrorAction SilentlyContinue | Select Id, ProcessName, StartTime | Format-Table -AutoSize } catch {}
    }

    throw "Could not connect to Edge DevTools on 9222 after waiting. See diagnostics above."
}

$target = $targets | Where-Object { $_.type -eq 'page' -and $_.url -and ($_.url -like "$url*") } | Select-Object -First 1
if (-not $target) { $target = $targets | Where-Object { $_.type -eq 'page' } | Select-Object -First 1 }
if (-not $target) {
    # Wait briefly for a page target to appear (sometimes the browser-level target is listed first)
    for ($j = 0; $j -lt 10; $j++) {
        Start-Sleep -Seconds 1
        try { $targets = Invoke-RestMethod "http://127.0.0.1:9222/json/list" -TimeoutSec 2 } catch {}
        $target = $targets | Where-Object { $_.type -eq 'page' } | Select-Object -First 1
        if ($target) { break }
    }
}
if (-not $target) { $target = $targets[0] }
$script:wsUrl = $target.webSocketDebuggerUrl

$script:ws = New-Object System.Net.WebSockets.ClientWebSocket
$script:ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(10)
$uri = [URI]$script:wsUrl
$script:ws.ConnectAsync($uri, [Threading.CancellationToken]::None).Wait()

            function Reconnect-CDP {
                param()
                Write-Host "Attempting to reconnect CDP WebSocket..."

                if (-not (Get-Variable -Name script:reconnectCount -Scope Script -ErrorAction SilentlyContinue)) { $script:reconnectCount = 0; $script:lastReconnect = (Get-Date).AddSeconds(-61) }
                if ((Get-Date) -lt $script:lastReconnect.AddSeconds(60)) { $script:reconnectCount++ } else { $script:reconnectCount = 1; $script:lastReconnect = Get-Date }
                if ($script:reconnectCount -gt 5) { Write-Host "Reconnect attempted $($script:reconnectCount) times within last minute; aborting."; return $false }

                try {
                    $targets = Invoke-RestMethod "http://127.0.0.1:9222/json/list" -TimeoutSec 5
                    if (-not $targets) { throw "No DevTools targets available" }
                    # Prefer the FIRST page target available in headless mode
                    $target = $targets | Where-Object { $_.type -eq 'page' } | Select-Object -First 1
                    if (-not $target) { $target = $targets[0] }
                    
                    $script:wsUrl = $target.webSocketDebuggerUrl
                    Write-Host "Reconnected to CDP at $($script:wsUrl)"
                    
                    if ($script:ws -ne $null) { 
                        try { $script:ws.Dispose() } catch {}
                    }
                    
                    $script:ws = New-Object System.Net.WebSockets.ClientWebSocket
                    $script:ws.Options.KeepAliveInterval = [TimeSpan]::FromSeconds(10)
                    $uri = [URI]$script:wsUrl
                    $script:ws.ConnectAsync($uri, [Threading.CancellationToken]::None).Wait()

                    # Re-enable events
                    $tmpId = [System.Threading.Interlocked]::Increment([ref]$script:cdpId)
                    Send-CDPRaw (@{ id=$tmpId; method="Page.enable"; params=@{} } | ConvertTo-Json -Compress)
                    $tmpId = [System.Threading.Interlocked]::Increment([ref]$script:cdpId)
                    Send-CDPRaw (@{ id=$tmpId; method="Runtime.enable"; params=@{} } | ConvertTo-Json -Compress)
                    return $true
                } catch {
                    Write-Host "Reconnect failed: $_"
                    return $false
                }
            }

function Send-CDPRaw {
    param($msgJson)
    if (-not $script:ws -or $script:ws.State -ne [System.Net.WebSockets.WebSocketState]::Open) {
        throw "WebSocket not open (state=$($script:ws.State))"
    }
    Write-Host "Send-CDPRaw: WebSocket state before send = $($script:ws.State)"
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($msgJson)
    $seg = New-Object ArraySegment[byte] -ArgumentList (,$buffer)
    # This low-level send does not attempt reconnects; callers handle reconnect logic
    try {
        $script:ws.SendAsync($seg, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, [Threading.CancellationToken]::None).Wait()
    } catch {
        # Surface the actual exception to caller
        throw
    }
}

function Send-CDP {
    param ($id, $method, $params)

    $msg = @{
        id = $id
        method = $method
        params = $params
    } | ConvertTo-Json -Depth 10

    try {
        Send-CDPRaw $msg
    } catch {
        Write-Host "Send-CDP failed for $method (id=$id): $_. Trying to reconnect..."
        $reconnected = Reconnect-CDP
        if ($reconnected) {
            Write-Host "Retrying Send-CDP $method (id=$id) after reconnect"
            try {
                Send-CDPRaw $msg
            } catch {
                throw "Could not send CDP message after reconnect: $_"
            }
        } else {
            throw "Could not reconnect CDP WebSocket: $_"
        }
    }
}

# Wait for the LoadEventFired event before proceeding (with progress)
function WaitForLoadEvent {
    param ($ws = $null, $timeoutSec = 30)

    if (-not $ws) { $ws = $script:ws }

    # Enable Network events so we can track inflight requests
    try { Send-CDP ([System.Threading.Interlocked]::Increment([ref]$script:cdpId)) "Network.enable" @{} } catch {}

    $isLoaded = $false
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    $start = Get-Date
    $nextLog = $start.AddSeconds(5)

    $reconnectAttempted = $false
    $inflight = 0
    $lastNetworkActivity = Get-Date

    while ((Get-Date) -lt $deadline) {
        $msgBuilder = New-Object System.Text.StringBuilder
        try {
            do {
                $rawBuffer = New-Object byte[] 65536
                $seg = New-Object System.ArraySegment[byte] (,$rawBuffer)
                $task = $script:ws.ReceiveAsync($seg, [Threading.CancellationToken]::None)
                if (-not $task.Wait(10000)) { throw "Receive timeout waiting for message" }
                $receive = $task.Result
                if ($receive.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { throw "WebSocket closed by remote" }
                if ($receive.Count -gt 0) { $msgBuilder.Append([Text.Encoding]::UTF8.GetString($rawBuffer, 0, $receive.Count)) | Out-Null }
            } while (-not $receive.EndOfMessage)
        } catch {
            $errMsg = $_.ToString()
            if ($errMsg -like '*Receive timeout*') {
                if (-not $script:consecTimeouts) { $script:consecTimeouts = 0 }
                $script:consecTimeouts++
                Write-Host "Receive timeout while waiting for Page.loadEventFired (consec=$script:consecTimeouts). Elapsed $([int]((Get-Date)-$start).TotalSeconds)s"

                # Do an active readyState probe when idle to detect completion without events
                try {
                    $ready = (SendAndWait 'Runtime.evaluate' @{ expression = "document.readyState" } 5).result.value
                    Write-Host "Active probe: document.readyState=$ready"
                    if ($ready -eq 'complete' -and $inflight -eq 0) { Write-Host "Active probe: ready & no inflight -> treating as loaded"; $isLoaded = $true; break }
                } catch {
                    Write-Host "Active probe failed: $_"
                }

                # Only attempt reconnect after a few consecutive timeouts
                if ($script:consecTimeouts -ge 3 -and -not $reconnectAttempted) {
                    Write-Host "Multiple consecutive receive timeouts; attempting reconnect..."
                    $reconnected = Reconnect-CDP
                    if ($reconnected) {
                        $reconnectAttempted = $true
                        # After reconnect, probe readyState
                        try {
                            $ready = (SendAndWait 'Runtime.evaluate' @{ expression = "document.readyState" } 5).result.value
                            if ($ready -eq 'complete') { Write-Host "Document readyState=complete after reconnect"; $isLoaded = $true; break }
                        } catch { Write-Host "Could not determine readyState after reconnect: $_" }
                        continue
                    } else {
                        throw "Could not reconnect CDP WebSocket while waiting for Page.loadEventFired"
                    }
                }

                Start-Sleep -Milliseconds 250
                continue
            }

            Write-Host "Receive failed while waiting for Page.loadEventFired: $_. Attempting reconnect..."
            if (-not $reconnectAttempted) {
                $reconnected = Reconnect-CDP
                if ($reconnected) {
                    # After reconnect, check readyState; if already 'complete', treat as loaded
                    try {
                        $ready = (SendAndWait 'Runtime.evaluate' @{ expression = "document.readyState" } 5).result.value
                        if ($ready -eq 'complete') {
                            Write-Host "Document readyState=complete after reconnect"
                            $isLoaded = $true
                            break
                        }
                    } catch {
                        Write-Host "Could not determine readyState after reconnect: $_"
                    }
                    $reconnectAttempted = $true
                    continue
                } else {
                    throw "Could not reconnect CDP WebSocket while waiting for Page.loadEventFired"
                }
            } else {
                throw "Receive failed again after reconnect attempt: $_"
            }
        }

        $msgText = $msgBuilder.ToString()
        try {
            $obj = $msgText | ConvertFrom-Json -ErrorAction Stop

            # Track network request lifecycle events
            switch ($obj.method) {
                'Network.requestWillBeSent' { $inflight++; $lastNetworkActivity = Get-Date }
                'Network.loadingFinished' { if ($inflight -gt 0) { $inflight-- }; $lastNetworkActivity = Get-Date }
                'Network.loadingFailed' { if ($inflight -gt 0) { $inflight-- }; $lastNetworkActivity = Get-Date }
            }

            if ($obj.method -eq "Page.loadEventFired") {
                $isLoaded = $true
                Write-Host "Event: Page.loadEventFired received after $([int]((Get-Date)-$start).TotalSeconds)s"
                break
            }

            if ($obj.method -eq 'Page.lifecycleEvent' -and $obj.params.name -eq 'networkIdle') {
                $isLoaded = $true
                Write-Host "Event: Page.lifecycleEvent networkIdle received after $([int]((Get-Date)-$start).TotalSeconds)s"
                break
            }

            # If no inflight requests for 2s and document.readyState is 'complete', consider loaded
            if ($inflight -eq 0 -and ((Get-Date) -gt $lastNetworkActivity.AddSeconds(2))) {
                try {
                    $ready = (SendAndWait 'Runtime.evaluate' @{ expression = "document.readyState" } 2).result.value
                    if ($ready -eq 'complete') {
                        Write-Host "No inflight requests and document.readyState=complete; treating as loaded"
                        $isLoaded = $true
                        break
                    }
                } catch {
                    # ignore evaluation errors and continue waiting
                }
            }

        } catch {
            # Ignore non-JSON or partial messages
            if ((Get-Date) -gt $nextLog) {
                Write-Host "Waiting for Page.loadEventFired... elapsed $([int]((Get-Date)-$start).TotalSeconds)s inflight=$inflight"
                $nextLog = (Get-Date).AddSeconds(5)
            }
            Start-Sleep -Milliseconds 100
        }
    }

    if (-not $isLoaded) {
        throw "Timed out waiting for LoadEventFired event after $timeoutSec seconds. inflight=$inflight"
    }
}  
            # Ensure counter exists in script scope for [ref]
            if (-not (Get-Variable -Name cdpId -Scope Script -ErrorAction SilentlyContinue)) { $script:cdpId = 1000 }
            function SendAndWait {
                param($method, $params, $timeoutSec = 30)

                $id = [System.Threading.Interlocked]::Increment([ref]$script:cdpId)
                Write-Host "→ CDP: $method (id=$id)"
                try {
                    Send-CDP $id $method $params
                } catch {
                    Write-Host "ERROR sending CDP ${method}: $_"
                    throw
                }

                $deadline = (Get-Date).AddSeconds($timeoutSec)
                $start = Get-Date
                $nextLog = $start.AddSeconds(5)
                $resendAttempted = $false
    while ((Get-Date) -lt $deadline) {
                    $msgBuilder = New-Object System.Text.StringBuilder
                    try {
                        do {
                            $rawBuffer = New-Object byte[] 65536
                            $seg = New-Object System.ArraySegment[byte] (,$rawBuffer)
                            $task = $script:ws.ReceiveAsync($seg, [Threading.CancellationToken]::None)
                            if (-not $task.Wait(12000)) { 
                                Write-Host "[Debug] ReceiveAsync wait exceeded 12s in SendAndWait"
                                throw "Receive timeout waiting for response to ${method}" 
                            }
                            $receive = $task.Result
                            if ($receive.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) { throw "WebSocket closed by remote" }
                            if ($receive.Count -gt 0) { $msgBuilder.Append([Text.Encoding]::UTF8.GetString($rawBuffer, 0, $receive.Count)) | Out-Null }
                        } while (-not $receive.EndOfMessage)
                    } catch {
                        Write-Host "Receive failed while waiting for response to ${method}: $_. Attempting reconnect..."
                        if (-not $resendAttempted) {
                            $reconnected = Reconnect-CDP
                            if ($reconnected) {
                                Write-Host "Resending $method (id=$id) after reconnect"
                                Send-CDP $id $method $params
                                $resendAttempted = $true
                                continue
                            } else {
                                throw "Could not reconnect CDP WebSocket while waiting for response to $method"
                            }
                        } else {
                            throw "Receive failed again after reconnect: $_"
                        }
                    }

                    $msgText = $msgBuilder.ToString()
                    try {
                        $obj = $msgText | ConvertFrom-Json -ErrorAction Stop
                        if ($obj.id -eq $id) { Write-Host "← CDP: response for id=$id ($method)"; return $obj }
                        # otherwise ignore events
                    } catch {
                        if ((Get-Date) -gt $nextLog) {
                            Write-Host "Waiting for response to ${method} (id=$id)... elapsed $([int]((Get-Date)-$start).TotalSeconds)s"
                            $nextLog = (Get-Date).AddSeconds(5)
                        }
                        Start-Sleep -Milliseconds 50
                    }
                }
                throw "Timed out waiting for response to $method (id=$id)"
            }

            # Crawler state
            $script:visited = @{}
            $script:pages = @()
            $script:savedPdfs = @()

            function SafeFilenameFromUrl {
                param($u)
                $u2 = $u -replace '^https?://','' -replace '[^a-zA-Z0-9_.-]','_'
                return $u2.Trim('_')
            }

            function Print-CurrentPageToPdf {
                param($u, $suffix = $null)
                try {
                    $printResp = SendAndWait "Page.printToPDF" @{
                        landscape = $false
                        printBackground = $false
                        preferCSSPageSize = $true
                        displayHeaderFooter = $false
                        marginTop = 0
                        marginBottom = 0
                        marginLeft = 0
                        marginRight = 0
                        scale = 1
                    } 30
                    if ($printResp -and $printResp.result -and $printResp.result.data) {
                        $pdfBytes = [Convert]::FromBase64String($printResp.result.data)
                        $safe = SafeFilenameFromUrl $u
                        $ts = (Get-Date -Format "yyyyMMdd_HHmmss")
                        $name = "page_{0}_{1}{2}.pdf" -f $ts, $safe, (if ($suffix) { "_{0}" -f $suffix } else { '' })
                        $outPath = Join-Path "C:\KVM-PDF" $name
                        [IO.File]::WriteAllBytes($outPath, $pdfBytes)
                        Write-Host "Saved page PDF to $outPath ($($pdfBytes.Length) bytes)"
                        $script:savedPdfs += $outPath
                        return $outPath
                    } else {
                        Write-Host "printToPDF returned no data for $u"
                        return $null
                    }
                } catch {
                    Write-Host "Print-CurrentPageToPdf failed for ${u}: $_"
                    return $null
                }
            }

            function WaitForContent {
                param($timeoutSec = 60, $minChars = $MinContentChars)

                $deadline = (Get-Date).AddSeconds($timeoutSec)
                $start = Get-Date
                $nextLog = $start.AddSeconds(5)
                $consecTimeouts = 0
                $screenshotTaken = $false

                $probeScript = @'
(function(){
  const candidates = [];
  const body = (document.body && document.body.innerText) ? document.body.innerText : '';
  candidates.push({name:'body', text: body});
  const sel = ['main','article','div[data-automationid="CanvasZone"]','div[class*="CanvasZone"]','div[id^="CanvasZone"]'];
  sel.forEach(s => { try { const e = document.querySelector(s); if (e) candidates.push({name:s, text: e.innerText||''}); } catch(e) {} });
  // visible text sample
  try {
    const visible = Array.from(document.querySelectorAll('body *')).filter(el => el.offsetParent !== null).map(el => el.innerText||'').join('\n');
    candidates.push({name:'visible', text: visible});
  } catch(e) {}
  return candidates.map(c=>({name:c.name, len: (c.text||'').length, sample: ((c.text||'').substr(0,200))}));
})();
'@

                while ((Get-Date) -lt $deadline) {
                    try {
                                $resp = SendAndWait 'Runtime.evaluate' @{ expression = $probeScript; awaitPromise = $true; returnByValue = $true } 5
                        $probes = $resp.result.result.value
                        $ok = $false
                        foreach ($p in $probes) {
                            $name = $p.name
                            $len = $p.len
                            if ($len -ge $minChars) {
                                Write-Host "WaitForContent: candidate '$name' has sufficient text (len=$len) after $([int]((Get-Date)-$start).TotalSeconds)s"
                                return $true
                            }
                        }

                        $consecTimeouts = 0
                    } catch {
                        $err = $_.ToString()
                        if ($err -like '*Timed out waiting for response*' -or $err -like '*Receive timeout*') {
                            $consecTimeouts++
                            Write-Host "WaitForContent: eval timeout (consec=$consecTimeouts). Elapsed $([int]((Get-Date)-$start).TotalSeconds)s"
                            if ($consecTimeouts -ge 3) {
                                Write-Host "WaitForContent: multiple eval timeouts, attempting Reconnect-CDP"
                                if (Reconnect-CDP) { $consecTimeouts = 0 } else { break }
                            }
                        } else {
                            Write-Host "WaitForContent: evaluation error: $_"
                        }
                    }

                    if (-not $screenshotTaken -and ((Get-Date) -gt $start.AddSeconds(8))) {
                        try {
                            $shot = SendAndWait 'Page.captureScreenshot' @{ format = 'png'; fromSurface = $true } 5
                            $bytes = [Convert]::FromBase64String($shot.result.data)
                            $path = Join-Path 'C:\KVM-PDF' ("debug_nav_{0}.png" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
                            [IO.File]::WriteAllBytes($path, $bytes)
                            Write-Host "Saved debug screenshot to $path"
                            $screenshotTaken = $true
                        } catch { Write-Host "Failed to take debug screenshot: $_" }
                    }

                    if ((Get-Date) -gt $nextLog) {
                        $best = ($probes | Sort-Object -Property { $_.len } -Descending | Select-Object -First 1)
                        if ($best) {
                            Write-Host "Waiting for page content... best so far: $($best.name) (len=$($best.len)). Need $minChars. Elapsed $([int]((Get-Date)-$start).TotalSeconds)s"
                        } else {
                            Write-Host "Waiting for page content... no data yet. Elapsed $([int]((Get-Date)-$start).TotalSeconds)s"
                        }
                        $nextLog = (Get-Date).AddSeconds(5)
                    }

                    # try to nudge lazy loading
                    try { SendAndWait 'Runtime.evaluate' @{ expression = "(function(){window.scrollTo(0, document.body.scrollHeight); return true; })()" } 2 | Out-Null } catch {}
                    Start-Sleep -Milliseconds 1000
                }

                # if we reach here, we timed out. try to return the best we have instead of throwing.
                try {
                   $finalProbe = SendAndWait "Runtime.evaluate" @{ expression = $probeScript; awaitPromise = $true; returnByValue = $true } 5
                   $res = $finalProbe.result.result.value
                   if ($res.probes) {
                       $best = ($res.probes | Sort-Object -Property { $_.len } -Descending | Select-Object -First 1)
                       if ($best -and $best.len -gt 0) {
                           Write-Host "WaitForContent: timed out but accepting best candidate '$($best.name)' with len=$($best.len)"
                           return $true
                       }
                   }
                } catch {}

                Write-Host "WaitForContent: timed out after $timeoutSec seconds with no significant text found."
                return $false
            }

            function Export-FirstPagePdf {
                param(
                    [string]$TargetUrl,
                    [string]$DestinationPath
                )

                Write-Host "Navigating to seed URL: $TargetUrl"
                $navResp = SendAndWait "Page.navigate" @{ url = $TargetUrl } 60
                if ($navResp.result.errorText) {
                    throw "Navigation failed: $($navResp.result.errorText)"
                }

                # Give modern SPAs time to render after initial navigation.
                Start-Sleep -Seconds 8

                # Verify what page we actually reached before printing.
                try {
                    $urlCheck = SendAndWait "Runtime.evaluate" @{ expression = "window.location.href"; returnByValue = $true } 20
                    $titleCheck = SendAndWait "Runtime.evaluate" @{ expression = "document.title"; returnByValue = $true } 20
                    $readyCheck = SendAndWait "Runtime.evaluate" @{ expression = "document.readyState"; returnByValue = $true } 20
                    $currentUrl = $urlCheck.result.result.value
                    $currentTitle = $titleCheck.result.result.value
                    $readyState = $readyCheck.result.result.value
                    Write-Host "Reached page URL: $currentUrl"
                    Write-Host "Page title: $currentTitle"
                    Write-Host "Ready state: $readyState"
                } catch {
                    Write-Host "Could not verify URL/title via Runtime.evaluate before print: $_"
                }

                # Extra settle time for delayed content (SharePoint script hydration).
                Start-Sleep -Seconds 6

                try {
                    SendAndWait "Emulation.setEmulatedMedia" @{ media = "screen" } 5 | Out-Null
                } catch {}

                $printResp = SendAndWait "Page.printToPDF" @{
                    landscape = $false
                    printBackground = $true
                    preferCSSPageSize = $true
                    displayHeaderFooter = $false
                    marginTop = 0.4
                    marginBottom = 0.4
                    marginLeft = 0.4
                    marginRight = 0.4
                } 60

                if (-not ($printResp -and $printResp.result -and $printResp.result.data)) {
                    throw "Page.printToPDF returned no data for $TargetUrl"
                }

                $pdfBytes = [Convert]::FromBase64String($printResp.result.data)
                [IO.File]::WriteAllBytes($DestinationPath, $pdfBytes)
                Write-Host "First-page PDF saved to $DestinationPath ($($pdfBytes.Length) bytes)"
                if ($pdfBytes.Length -lt 1024) {
                    Write-Host "Warning: PDF is unusually small; page may still be rendering or authentication may be required."
                }
            }

            $startUri = [URI]$url

            function Normalize-Url {
                param($base, $href)
                try { return ([URI]::new([URI]$base, $href)).AbsoluteUri } catch { return $null }
            }

            function Is-SameOrigin($u1, $u2) {
                try {
                    $a = [URI]$u1; $b = [URI]$u2
                    return ($a.Scheme -eq $b.Scheme -and $a.Host -eq $b.Host -and $a.Port -eq $b.Port)
                } catch { return $false }
            }

            function Crawl-Url {
                param($u, $depth)
                if ($script:visited.ContainsKey($u)) { return }
                Write-Host "Crawling $u (depth $depth)"
                $script:visited[$u] = $true

                # Navigate and wait for content or load
                Write-Host "Navigating to $u..."
                try {
                    $null = SendAndWait "Page.navigate" @{ url = $u } 20
                } catch {
                    Write-Host "Warning: Navigation to $u timed out or failed. Skipping this branch: $_"
                    return
                }

                # Demand at least 200 chars for the main page, maybe less for linked ones
                $threshold = if ($depth -eq $MaxDepth) { 200 } else { 100 }
                try {
                    $null = WaitForContent -timeoutSec 30 -minChars $threshold
                } catch {
                    Write-Host "Warning: Content wait for $u failed/timed out ($threshold chars needed). Continuing with best available content."
                }

                # Print individual page PDF
                Write-Host "Triggering Print-CurrentPageToPdf for $u"
                Print-CurrentPageToPdf $u | Out-Null

                # Slight pause to allow dynamic text to settle
                Start-Sleep -Milliseconds 500

                $extractScript = 
@'
            (function(){
              // Prefer main content areas for text extraction
              const selectors = ['main', 'article', 'div[data-automationid="CanvasZone"]', '.CanvasZone', '#main-content'];
              let text = '';
              for (const s of selectors) {
                const e = document.querySelector(s);
                if (e && e.innerText && e.innerText.length > text.length) text = e.innerText;
              }
              if (!text || text.length < 200) text = document.body ? document.body.innerText : '';
              
              const anchors = Array.from(document.querySelectorAll('a[href]')).map(a => a.getAttribute('href'));
              return { text, anchors };
            })();
'@

                $evalResp = SendAndWait "Runtime.evaluate" @{ expression = $extractScript; awaitPromise = $true; returnByValue = $true }
                $val = $evalResp.result.result.value
                
                if ($val.text) {
                    $cleanedText = $val.text -replace "\r?\n{2,}", "`n`n"
                    $script:pages += @{ url = $u; text = $cleanedText }
                }

                if ($depth -gt 0 -and $val.anchors) {
                    $anchors = $val.anchors | Where-Object { $_ -and ($_ -ne '') }
                    $anchorCount = ($anchors | Measure-Object).Count
                    Write-Host "Found $anchorCount anchors on $u"
                    foreach ($href in $anchors) {
                        $abs = Normalize-Url $u $href
                        if (-not $abs) { continue }
                        # ignore mailto: and javascript: and external origins
                        if ($abs.StartsWith('mailto:') -or $abs.StartsWith('javascript:')) { continue }
                        if (-not (Is-SameOrigin $startUri.AbsoluteUri $abs)) { continue }
                        # Only follow .aspx pages (skip images and other extensions)
                        try {
                            $uriObj = [Uri]$abs
                            $ext = [IO.Path]::GetExtension($uriObj.AbsolutePath).ToLower()
                        } catch {
                            $ext = ''
                        }
                        if ($ext -and $ext -ne '.aspx') {
                            continue
                        }
                        if (-not $script:visited.ContainsKey($abs)) {
                            Crawl-Url $abs ($depth - 1)
                        }
                    }
                }
            }

if ($Headless) {
    Write-Host "Headless single-page mode active. Target PDF: $outPath"
    try {
        # Establish initial connection to a page target
        $reconnected = Reconnect-CDP
        if (-not $reconnected) { throw "Initial CDP connection failed." }

        Export-FirstPagePdf -TargetUrl $url -DestinationPath $outPath
    } catch {
        Write-Host "ERROR during headless first-page export: $_"
    }
} else {
    # Non-headless/Single-page flow
    Write-Host "Single-page PDF flow..."
    try {
        Export-FirstPagePdf -TargetUrl $url -DestinationPath $outPath
    } catch {
        Write-Host "ERROR: $_"
    }
}

# Close WebSocket if open
if ($script:ws -and $script:ws.State -eq 'Open') {
    try { $script:ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, "Finished", [System.Threading.CancellationToken]::None).Wait() } catch {}
}
Write-Host "Script finished."
