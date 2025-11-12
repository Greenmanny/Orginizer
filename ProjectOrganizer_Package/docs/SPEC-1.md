# SPEC‑1 — Project Folder Organizer & Template Manager

## Background
- Generate new projects with a consistent structure & vendor seeds (TBC, Civil 3D).
- Organize existing projects inside a shared, synced root. No admin rights.
- Windows 10/11; Syncthing integration (auto‑pause/rescan; `.stignore` read‑only).

## Canonical Layout
`PROJECTS/FY<YY>/<Mission>/<Job>/` with:
- `20_TBC/`, `30_CAD/`, `30_CAD/_Templates/`, `40_PDF/`, `50_Import/`, `60_Soil/ (later)`, `90_Archive/`, `manifest.json`

## Seeds
- `_Organizer/Seeds/TBC/*` → `20_TBC/`
- `_Organizer/Seeds/Civil3D/*` → `30_CAD/_Templates/`
- USACE Civil 3D 2022 additions (Survey + Site Design discipline templates; Plan Production & Sheet templates).
- Organizer never edits Civil 3D profiles; standards root is configured once per workstation.

## Run
```
pwsh -STA -ExecutionPolicy Bypass -File "%USERPROFILE%/PROJECTS/_Organizer/Organizer.ps1"
```

## Notes
- WinForms GUI (PowerShell). Recommended runtime: PowerShell 7.4 LTS.
- Dry‑run preview, then apply; logs + per‑project manifest.
