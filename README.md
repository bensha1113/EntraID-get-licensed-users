# Microsoft 365 Licensed Users Report

`Get-licensed-users.ps1` exports a polished Microsoft 365 licensed-user report (HTML + optional PDF) directly from Microsoft Graph. It is designed for helpdesk and licensing teams that need an at-a-glance breakdown of licensed accounts, SKU usage, and recommended lifecycle actions.

## Highlights
- **Graph-powered inventory**: pulls all users, license assignments, and subscribed SKUs via Microsoft Graph.
- **Sign-in awareness**: correlates recent sign-ins from the audit log to flag inactive accounts (default 90 days).
- **Keep/Delete workflow**: color-coded badges plus summary cards show which users to keep or remove. Badges are clickable in the HTML so you can tweak decisions before printing.
- **Manual overrides**: optional CSV overrides let you enforce decisions per UPN/email.
- **Beautiful outputs**:
  - Modern web view with print button.
  - Legacy-print mode that mirrors classic corporate export styling and fits neatly on A4.
  - Optional PDF conversion via `wkhtmltopdf` or Edge/Chrome headless printing.
- **Automatic archiving**: every run stores HTML/PDF copies under `User Reports/<tenant-domain>/<timestamp>` inside the active user’s Documents (OneDrive-aware).

## Prerequisites
- PowerShell 7+ (recommended) with internet access.
- Microsoft Graph PowerShell SDK (`Microsoft.Graph`, installed automatically if missing).
- Azure AD account with delegated permissions: `User.Read.All`, `Directory.Read.All`, `AuditLog.Read.All`.
- Optional PDF conversion tools:
  - [`wkhtmltopdf`](https://wkhtmltopdf.org/) (preferred) **or**
  - Microsoft Edge / Google Chrome (headless print-to-pdf).

## Parameters (abridged)
| Parameter | Description | Default |
|-----------|-------------|---------|
| `OutputPdfPath` | Custom PDF destination. When omitted, output lands in the timestamped archive folder. | auto-generated |
| `Language` | `English`, `Swedish`, or `Bilingual` labels. | `English` |
| `UseWkHtml` | Forces wkhtmltopdf conversion instead of Edge/Chrome. | `false` |
| `InactiveThresholdDays` | Days without sign-ins before a user is marked “Delete”. | `90` |
| `DecisionOverrideCsvPath` | CSV with per-user keep/delete overrides (columns: `UPN`, `Action`). | none |

## Running the script
```powershell
# Basic run (HTML + PDF archived under Documents\User Reports)
pwsh ./Get-licensed-users.ps1

# Specify a custom PDF path
pwsh ./Get-licensed-users.ps1 -OutputPdfPath C:\Reports\M365-LicensedUsers.pdf

# Supply overrides and stricter inactivity rule
pwsh ./Get-licensed-users.ps1 -InactiveThresholdDays 60 -DecisionOverrideCsvPath .\overrides.csv
```
The first run will prompt for Microsoft Graph consent. Make sure the signed-in account has the required scopes.

## Override CSV format
```csv
UPN,Action
user1@contoso.com,Keep
user2@contoso.com,Delete
```
Supported synonyms include `Keep/Retain/Green/Stay` and `Delete/Remove/Drop/Red`.

## Output location and structure
Each execution creates:
```
Documents\\User Reports\\<tenant-domain>\\<yyyyMMdd-HHmmss>\\
    ├── M365-LicensedUsers.html
    └── M365-LicensedUsers.pdf   (if conversion succeeds or custom path copy)
```
If you pass a custom `OutputPdfPath`, the PDF is also copied into the timestamp folder for auditing.

## Customizing the report
- Modify the CSS inside the script to match your brand colors.
- Translate or adjust localized strings in `Get-LocalizedText`.
- Update the legacy layout if you need different print margins.

## Troubleshooting
- **No PDF produced**: ensure Edge/Chrome or wkhtmltopdf is installed and accessible in `PATH`. The script falls back to HTML and prints instructions if conversion fails.
- **Sign-ins missing**: audit logs only retain ~30 days by default; adjust `InactiveThresholdDays` or provide manual overrides.
- **Permission errors**: re-run `Connect-MgGraph` with an account that has the listed scopes.

## License
Copyright (c) 2025 Bensha1113.

Released under the [MIT License](LICENSE):

```
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

Feel free to adapt the script for your organization. Consider removing tenant-specific details before publishing publicly.
