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

---

## Installation

### Option 1: Quick try-out (Downloads folder)
Download the script to your Downloads folder and make it executable:
```bash
curl -o ~/Downloads/dock-deployer.zsh https://github.com/jamessorrenti/toolbox/docker_deploy/raw/main/dock-deployer.zsh && chmod +x ~/Downloads/dock-deployer.zsh
```

Run it:
```bash
~/Downloads/dock-deployer.zsh --help
```

### Option 2: Install globally (recommended for reuse)
Download directly into `/usr/local/bin` so it’s available system-wide:
```bash
curl -o /usr/local/bin/dock-deployer https://github.com/jamessorrenti/toolbox/docker_deploy/raw/main/dock-deployer.zsh && chmod +x /usr/local/bin/dock-deployer
```

Now you can run:
```bash
dock-deployer --help
```

---

## Usage

### Export current Dock and set once (default, non-persistent)
```bash
dock-deployer --remove "Maps,News" --add "/Applications/Slack.app" --hide-recents
```

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
- `--forced` → install the profile as persistent managed Dock  
- `--uninstall` → remove the installed profile file at `--out`  

---

## Notes
- By default, Dock Deployer makes one-time edits to your Dock (set-once mode).  
- Use `--forced` for a persistent, managed Dock profile that remains enforced until you uninstall it.  
- Finder and Trash are always present and cannot be removed.  
- For additional options, import your profile (e.g., `Dock_export-MyMac.mobileconfig`) into a tool such as **iMazing Profile Editor** for further editing.

