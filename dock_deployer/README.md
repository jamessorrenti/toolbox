# Dock Deployer

A lightweight `zsh` utility for exporting, editing, and deploying **macOS Dock configurations**.  
No external dependencies, just a single `.zsh` script and macOS built-in tools.

---

## Features
- Export your current Dock into a `.mobileconfig` file
- Remove apps by name (e.g., `Maps`, `News`)
- Add apps by path (e.g., `/Applications/Google Chrome.app`)
- Toggle Dock options like `show-recents`
- Deploy directly to your Dock (default, one-time set-once mode)
- Optionally enforce Dock layout with a persistent managed profile (`--forced`)
- Export Dock changes to a `.mobileconfig` file (`--export`)
- Export only (generate `.mobileconfig`, then restore Dock) (`--export-only`)

---

## Installation

### Option 1: Quick try-out (Downloads folder)
Download the script to your Downloads folder and make it executable:
```bash
curl -L -o ~/Downloads/dock-deployer.zsh https://raw.githubusercontent.com/jamessorrenti/toolbox/main/dock_deployer/dock-deployer.zsh && chmod +x ~/Downloads/dock-deployer.zsh
```

Run it:
```bash
# See available options
~/Downloads/dock-deployer.zsh --help

# Quick export of your current Dock (no edits, just writes a .mobileconfig)
~/Downloads/dock-deployer.zsh --export-only
```

### Option 2: Install globally (recommended for reuse)
Download directly into `/usr/local/bin` so it’s available system-wide:
```bash
curl -L -o /usr/local/bin/dock-deployer https://raw.githubusercontent.com/jamessorrenti/toolbox/main/dock_deployer/dock-deployer.zsh && chmod +x /usr/local/bin/dock-deployer
```

Now you can run:
```bash
dock-deployer --help
```

---

## Usage

### Export current Dock only (no edits applied)
```bash
dock-deployer --export-only
```
This writes your current Dock configuration to `~/Desktop/Dock_export-<hostname>.mobileconfig` without changing your Dock at all.

---

### Edit current Dock (set once) *and* export to file
```bash
dock-deployer --remove "Maps,News" --add "/Applications/Slack.app" --hide-recents --export
```
This applies changes to your Dock immediately (non-persistent) **and** writes a `.mobileconfig` to `~/Desktop/Dock_export-<hostname>.mobileconfig`.

---

### Reset Dock to defaults, then apply edits
```bash
dock-deployer --reset --remove "Maps,News" --add "/Applications/Slack.app" --hide-recents
```

### Force a persistent managed profile
```bash
dock-deployer --remove "Maps" --add "/Applications/Slack.app" --hide-recents --forced
```

### Uninstall profile
```bash
dock-deployer --uninstall
```

### Custom output path for .mobileconfig
```bash
dock-deployer --forced --out ~/Desktop/MyDock.mobileconfig
```

---

## Deploying an exported profile manually

If you already have a profile file such as `Dock_export-MyMac.mobileconfig` copied to another Mac, you can install or remove it without using the script:

**Install the profile:**
```bash
sudo profiles -I -F ~/Desktop/Dock_export-MyMac.mobileconfig
```

**Remove the profile:**
```bash
sudo profiles -R -F ~/Desktop/Dock_export-MyMac.mobileconfig
```

---

## Options
- `--out <path>` → custom output location for `.mobileconfig`  
- `--remove <comma list>` → remove apps by label (case-insensitive)  
- `--add <comma list>` → add apps by full `.app` paths  
- `--hide-recents` → disable Recents section in Dock  
- `--reset` → reset Dock to macOS defaults before applying edits  
- `--export` → apply changes once and also write a `.mobileconfig`  
- `--export-only` → generate `.mobileconfig` and restore Dock to original state  
- `--forced` → install the profile as persistent managed Dock  
- `--uninstall` → remove the installed profile file at `--out`  

---

## Notes
- By default, Dock Deployer makes one-time edits to your Dock (set-once mode).  
- Use `--forced` for a persistent, managed Dock profile that remains enforced until you uninstall it.  
- Use `--export` or `--export-only` to generate `.mobileconfig` files without needing persistent management.  
- Finder and Trash are always present and cannot be removed.  
- For additional options, import your profile (e.g., `Dock_export-MyMac.mobileconfig`) into a tool such as **iMazing Profile Editor** for further editing.
