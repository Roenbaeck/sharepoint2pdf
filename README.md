# sharepoint2pdf

PowerShell script that exports a SharePoint page URL to PDF using Microsoft Edge in headless mode.

## What it does

- Opens the provided URL in headless Edge
- Prints the page to PDF
- Writes output to a directory you choose
- Names the PDF from the URL path (for example, `YourPage.pdf`)
- Optionally crawls linked `.aspx` pages and exports each to PDF

## Requirements

- Windows
- Microsoft Edge installed
- PowerShell 5.1+ (or PowerShell 7+)

## Script

- `Sharepoint2PDF.ps1`

## Usage

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\Exports"
```

## Parameters

- `-Url` (required)
	- SharePoint page URL to export
- `-Output` (required)
	- Output directory
	- If a `.pdf` path is provided, the script uses its parent directory
- `-UserDataDir` (optional)
	- Edge user data directory
	- If omitted, a temporary profile directory is used
- `-WaitSec` (optional, default: `120`)
	- Virtual-time/timeout budget passed to Edge (seconds)
- `-ViewportWidth` (optional, default: `1200`)
	- Viewport width in pixels
	- Height is auto-calculated to A4 portrait ratio
- `-WarmupSec` (optional, default: `0`)
	- Optional warm-up headless print pass before final export
- `-MaxDepth` (optional, default: `1`)
	- Crawl depth for linked `.aspx` pages
	- `0` exports only the initial `-Url`
- `-MaxPages` (optional, default: `30`)
	- Maximum number of pages to export during crawl
- `-BaseUrl` (optional)
	- Base URL/prefix boundary for crawl scope
	- If omitted, the script derives a base prefix from `-Url`
- `-CdpRenderWaitSec` (optional, default: `15`)
	- Extra post-load wait for CDP discovery/print

## Examples

### Basic export

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\Exports"
```

### Save to a specific directory with longer wait

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\Exports" -WaitSec 240
```

### Use custom viewport width

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\Exports" -ViewportWidth 1100
```

### Use warm-up pass

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\Exports" -WarmupSec 30
```

### Export only the initial page (no crawling)

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\Exports" -MaxDepth 0
```

### Crawl linked `.aspx` pages one level deep

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/Home.aspx" -Output "C:\Exports" -MaxDepth 1 -MaxPages 50
```

### Use explicit crawl scope

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/Home.aspx" -Output "C:\Exports" -BaseUrl "https://your-tenant.sharepoint.com/sites/YourSite/"
```

### Set extra render wait

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/Home.aspx" -Output "C:\Exports" -MaxDepth 1 -CdpRenderWaitSec 25
```

## Notes

- Some Chromium/Edge warning logs (GPU/USB/task manager) can appear in console output and are typically non-fatal.
- `WaitSec` is an upper-bound budget for headless rendering/print timing, not a guaranteed wall-clock delay.
- Crawling follows only in-scope links ending in `.aspx`.
- Some pages may render non-navigation cards without standard `<a href="...">` elements; in such cases, extra selector tuning may be required.

## Troubleshooting

- **Auth/session issues (blank output, tiny PDF, or missing protected content)**
	- Headless Edge may reuse an existing session/profile state.
	- If export fails due to authentication, open the same `-Url` interactively in Edge first and complete sign-in/MFA.
	- Running the script shortly after visiting the target URL can help ensure the session is still valid.
	- If needed, pass `-UserDataDir` to point to a profile that already has valid SharePoint auth.
