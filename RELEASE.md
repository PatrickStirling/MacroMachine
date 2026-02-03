# Release Checklist

Use this flow for every update.

## 1) Bump version
Update the version in:
- `FusionMacroReorderer-electron/package.json`
  - Ensure the file is UTF-8 **without BOM** (macOS CI fails if a BOM is present).

## 2) Commit
```powershell
git add FusionMacroReorderer-electron/package.json
git commit -m "Bump version to X.Y.Z"
```

## 3) Tag + push
```powershell
git tag vX.Y.Z
git push origin main --tags
```

## 4) Verify CI
GitHub → Actions → **Build and Publish**:
- Both **windows-latest** and **macos-latest** should be green.

## 5) Verify release assets
GitHub → Releases → `vX.Y.Z`:
- Windows: `.exe`, `.exe.blockmap`, `latest.yml`
- macOS: `.dmg`, `.zip`, `.zip.blockmap`, `latest-mac.yml`
