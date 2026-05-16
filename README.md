# expecto-comparo

Easy comparison of docx suites using Word.

## First implementation: PowerShell WinForms prototype

The repository now includes `/home/runner/work/expecto-comparo/expecto-comparo/Expecto-Comparo.ps1`, a Windows PowerShell GUI tool focused on non-technical users.

### What it does

- Select Previous, Current, and Output folders
- Scan `.docx` files from both folders
- Suggest conservative pairings using:
  - exact filename match
  - normalized filename match (ignoring common version/date markers)
  - fuzzy suggestions (clearly marked)
- Show pairings in a side-by-side editable view
- Allow manual pair and unpair actions before execution
- Run Microsoft Word COM comparisons sequentially
- Save comparison documents with readable, overwrite-safe names
- Write a timestamped run log in the output folder
- Show progress and completion summary (success/skipped/failed)

### Requirements

- Windows desktop
- Microsoft Word installed locally
- Local or locally-synced access to `.docx` files (OneDrive/SharePoint synced folders supported)
- PowerShell with WinForms support

### Run

From Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\Expecto-Comparo.ps1
```

### Notes

- Comparisons run one at a time (sequentially).
- The tool opens source files read-only where possible.
- If output filenames already exist, a numeric suffix is appended.
- If OneDrive/SharePoint-style paths are detected, the tool logs a reminder to ensure files are locally available.
