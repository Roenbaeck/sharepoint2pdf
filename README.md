# sharepoint2pdf

PowerShell script that exports a SharePoint page URL to PDF using Microsoft Edge in headless mode.

## What it does

- Opens the provided URL in headless Edge
- Prints the page to PDF
- Writes output to a directory you choose
- Names the PDF from the URL path (for example, `YourPage.pdf`)

## Requirements

- Windows
- Microsoft Edge installed
- PowerShell 5.1+ (or PowerShell 7+)

## Script

- `Sharepoint2PDF.ps1`

## Usage

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx"
```

## Parameters

- `-Url` (required)
	- SharePoint page URL to export
- `-Output` (optional, default: `C:\KVM-PDF`)
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

## Examples

### Basic export

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx"
```

### Save to a specific directory with longer wait

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -Output "C:\KVM-PDF" -WaitSec 240
```

### Use custom viewport width

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -ViewportWidth 1100
```

### Use warm-up pass

```powershell
.\Sharepoint2PDF.ps1 -Url "https://your-tenant.sharepoint.com/sites/YourSite/SitePages/YourPage.aspx" -WarmupSec 30
```

## Notes

- Some Chromium/Edge warning logs (GPU/USB/task manager) can appear in console output and are typically non-fatal.
- `WaitSec` is an upper-bound budget for headless rendering/print timing, not a guaranteed wall-clock delay.

## Troubleshooting

- **Auth/session issues (blank output, tiny PDF, or missing protected content)**
	- Headless Edge may reuse an existing session/profile state.
	- If export fails due to authentication, open the same `-Url` interactively in Edge first and complete sign-in/MFA.
	- Running the script shortly after visiting the target URL can help ensure the session is still valid.
	- If needed, pass `-UserDataDir` to point to a profile that already has valid SharePoint auth.
