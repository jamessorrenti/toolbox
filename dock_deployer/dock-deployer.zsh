#!/bin/zsh
# Dock Deployer
# Default: apply Dock edits once to the current user's Dock (non-persistent)
# Optional: --forced to install as a managed profile (persistent)
# Optional: --reset to reset Dock to macOS defaults before applying edits
# New: --export writes a .mobileconfig alongside set-once changes
# New: --export-only writes a .mobileconfig but restores your Dock afterward

set -e
set -o pipefail
set -u

# -------- Helpers --------
get_hostname() {
  local name
  name=$(scutil --get ComputerName 2>/dev/null || true)
  [[ -z "$name" ]] && name=$(hostname -s)
  echo "$name"
}

print_usage() {
  cat <<'USAGE'
Dock Deployer

Default: edits Dock once (non-persistent).
Use --forced for a persistent managed profile.
Use --reset to start from macOS default Dock.
Use --export to also write a .mobileconfig.
Use --export-only to write a .mobileconfig and restore your Dock afterward.

Options:
  --out <path>                       Output .mobileconfig path (default: ~/Desktop/Dock_export-{hostname}.mobileconfig)
  --remove <comma-separated labels>  Remove Dock items by label (e.g., "Mail,Maps,Photos,Apple TV,News,Freeform")
  --add <comma-separated app paths>  Add apps by full .app paths (e.g., "/Applications/Google Chrome.app,/Applications/Slack.app")
  --hide-recents                     Set show-recents=false
  --reset                            Reset Dock to macOS defaults before applying changes
  --export                           Also write .mobileconfig (no install unless --forced)
  --export-only                      Write .mobileconfig, then restore original Dock (ignores --forced)
  --forced                           Install as a persistent managed profile
  --uninstall                        Uninstall the profile file at --out
  --help                             Show this help
USAGE
}

# -------- Defaults --------
HOSTNAME_SAFE="$(get_hostname | tr ' ' '-')"
OUT_PATH="$HOME/Desktop/Dock_export-${HOSTNAME_SAFE}.mobileconfig"
REMOVE_LABELS=()
ADD_APPS=()
HIDE_RECENTS=false
FORCED=false
UNINSTALL=false
RESET=false
EXPORT=false
EXPORT_ONLY=false

# -------- Parse args --------
while [[ $# -gt 0 ]]; do
  case "$1" in
    --out) OUT_PATH="$2"; shift 2 ;;
    --remove) IFS=, read -rA REMOVE_LABELS <<< "$2"; shift 2 ;;
    --add) IFS=, read -rA ADD_APPS <<< "$2"; shift 2 ;;
    --hide-recents) HIDE_RECENTS=true; shift ;;
    --forced) FORCED=true; shift ;;
    --uninstall) UNINSTALL=true; shift ;;
    --reset) RESET=true; shift ;;
    --export) EXPORT=true; shift ;;
    --export-only) EXPORT_ONLY=true; shift ;;
    --help|-h) print_usage; exit 0 ;;
    *) echo "Unknown option: $1" >&2; print_usage; exit 1 ;;
  esac
done

# -------- Portable array join --------
{ IFS=,; REMOVE_CSV="${REMOVE_LABELS[*]}"; }
{ IFS=,; ADD_CSV="${ADD_APPS[*]}"; }

# -------- Reset Dock --------
if $RESET; then
  echo "Resetting Dock to macOS defaults..."
  defaults delete com.apple.dock || true
  killall Dock || true
  sleep 2
fi

# -------- Uninstall --------
if $UNINSTALL; then
  echo "Uninstalling profile at $OUT_PATH (sudo)..."
  sudo profiles -R -F "$OUT_PATH"
  exit 0
fi

# -------- Shared Python helpers (embedded) --------
gen_mobileconfig_py='
import os, sys, uuid, plistlib, datetime, subprocess, re

mode = sys.argv[1]  # gen_mobileconfig or set_once
out_path = sys.argv[2]
hide_recents = sys.argv[3].lower() == "true"
remove_csv = sys.argv[4]
add_csv = sys.argv[5]

remove_labels = [s.strip() for s in (remove_csv.split(",") if remove_csv else []) if s.strip()]
add_apps = [s.strip() for s in (add_csv.split(",") if add_csv else []) if s.strip()]

src = os.path.expanduser("~/Library/Preferences/com.apple.dock.plist")
if not os.path.exists(src):
    raise SystemExit(f"Dock plist not found: {src}")

with open(src, "rb") as f:
    dock = plistlib.loads(f.read())

labels_lower = {l.lower() for l in remove_labels}

def is_match(item):
    td = item.get("tile-data", {})
    lbl = (td.get("file-label") or "").lower()
    if lbl in labels_lower:
        return True
    bid = (td.get("bundle-identifier") or "").lower()
    if "maps" in labels_lower and bid == "com.apple.maps":
        return True
    if "news" in labels_lower and bid == "com.apple.news":
        return True
    if "photos" in labels_lower and bid == "com.apple.photos":
        return True
    if ("apple tv" in labels_lower or "tv" in labels_lower) and bid == "com.apple.tv":
        return True
    path = (td.get("file-data", {}).get("_CFURLString") or "").lower()
    m = re.search(r"/applications/([^/]+)\.app", path or "")
    if m and m.group(1).replace("-", " ").lower() in labels_lower:
        return True
    return False

def normalize_app_path(p):
    p = p or ""
    p = p.replace("file://", "")
    try:
        p = os.path.realpath(p)
    except Exception:
        pass
    return p.lower().rstrip("/")

def make_tile_for_app(app_path):
    url = "file://" + app_path
    label = os.path.splitext(os.path.basename(app_path))[0]
    tile = {
        "tile-data": {
            "file-data": {"_CFURLString": url, "_CFURLStringType": 15},
            "file-label": label
        },
        "tile-type": "file-tile"
    }
    try:
        bid = subprocess.check_output(
            ["mdls", "-name", "kMDItemCFBundleIdentifier", "-r", app_path],
            stderr=subprocess.DEVNULL, text=True
        ).strip()
        if bid and bid != "(null)":
            tile["tile-data"]["bundle-identifier"] = bid
    except Exception:
        bid = None
    return tile

def any_match(existing_item, label, bid, norm_path):
    td = existing_item.get("tile-data", {})
    ex_bid = (td.get("bundle-identifier") or "").lower()
    if bid and ex_bid and ex_bid == bid.lower():
        return True
    ex_label = (td.get("file-label") or "").strip().lower()
    if ex_label and label.lower() == ex_label:
        return True
    ex_url = (td.get("file-data", {}).get("_CFURLString") or "")
    ex_norm = normalize_app_path(ex_url)
    if ex_norm and ex_norm == norm_path:
        return True
    return False

def already_present(d, app_path, label, bid):
    norm = normalize_app_path(app_path)
    for it in d.get("persistent-apps", []):
        if any_match(it, label, bid, norm):
            return True
    return False

# Apply removals
if "persistent-apps" in dock and labels_lower:
    dock["persistent-apps"] = [i for i in dock["persistent-apps"] if not is_match(i)]

# Apply additions
for p in add_apps:
    if not p.endswith(".app"):
        continue
    if not os.path.exists(p):
        candidates = [f"/Applications/{p}", f"/System/Applications/{p}"]
        p = next((c for c in candidates if os.path.exists(c)), p)
        if not os.path.exists(p):
            continue

    label = os.path.splitext(os.path.basename(p))[0]
    try:
        bid = subprocess.check_output(
            ["mdls", "-name", "kMDItemCFBundleIdentifier", "-r", p],
            stderr=subprocess.DEVNULL, text=True
        ).strip()
        if bid == "(null)":
            bid = ""
    except Exception:
        bid = ""

    dock.setdefault("persistent-apps", [])
    if not already_present(dock, p, label, bid):
        dock["persistent-apps"].append(make_tile_for_app(p))

if hide_recents:
    dock["show-recents"] = False

if mode == "set_once":
    with open(src, "wb") as f:
        f.write(plistlib.dumps(dock, fmt=plistlib.FMT_BINARY))
    print("SET_ONCE_DONE")
    sys.exit(0)

wanted = ["persistent-apps","persistent-others","show-recents","autohide","magnification","tilesize","largesize","mru-spaces"]
dock_subset = {k: dock[k] for k in wanted if k in dock}

profile_uuid = str(uuid.uuid4()).upper()
payload_uuid = str(uuid.uuid4()).upper()

payload = {
    "PayloadType": "com.apple.ManagedClient.preferences",
    "PayloadVersion": 1,
    "PayloadIdentifier": f"com.example.dock.custom.{payload_uuid}",
    "PayloadUUID": payload_uuid,
    "PayloadEnabled": True,
    "PayloadDisplayName": "Dock (Custom Settings)",
    "PayloadContent": {
        "com.apple.dock": {
            "Forced": [ {"mcx_preference_settings": dock_subset} ]
        }
    }
}

mobileconfig = {
    "PayloadType": "Configuration",
    "PayloadVersion": 1,
    "PayloadIdentifier": f"com.example.dock.export.{profile_uuid}",
    "PayloadUUID": profile_uuid,
    "PayloadDisplayName": "Dock Deploy " + str(uuid.uuid4()),
    "PayloadRemovalDisallowed": False,
    "PayloadContent": [payload],
}

os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, "wb") as f:
    plistlib.dump(mobileconfig, f, fmt=plistlib.FMT_XML)
print(out_path)
'

# -------- EXPORT-ONLY: backup, export, restore --------
if $EXPORT_ONLY; then
  echo "Note: --export-only ignores --forced. Creating export without changing your Dock..."
  ts="$(date +%Y%m%d-%H%M%S)"
  backup="/tmp/Dock_backup-${ts}.plist"
  plist="$HOME/Library/Preferences/com.apple.dock.plist"
  if [[ ! -f "$plist" ]]; then
    echo "Dock plist not found at $plist" >&2
    exit 1
  fi
  cp "$plist" "$backup"
  /usr/bin/python3 - <<PY gen_mobileconfig "$OUT_PATH" "$HIDE_RECENTS" "$REMOVE_CSV" "$ADD_CSV"
$gen_mobileconfig_py
PY
  cp "$backup" "$plist"
  rm -f "$backup"
  killall Dock || true
  echo "Exported to: $OUT_PATH (Dock restored to original state)"
  exit 0
fi

# -------- FORCED (persistent) path --------
if $FORCED; then
  /usr/bin/python3 - <<PY gen_mobileconfig "$OUT_PATH" "$HIDE_RECENTS" "$REMOVE_CSV" "$ADD_CSV"
$gen_mobileconfig_py
PY
  echo "Installing profile (sudo)..."
  sudo profiles -I -F "$OUT_PATH"
  killall Dock || true
  exit 0
fi

# -------- Default SET-ONCE path (optionally with --export) --------
/usr/bin/python3 - <<PY set_once "$OUT_PATH" "$HIDE_RECENTS" "$REMOVE_CSV" "$ADD_CSV"
$gen_mobileconfig_py
PY
killall Dock || true

if $EXPORT; then
  /usr/bin/python3 - <<PY gen_mobileconfig "$OUT_PATH" "$HIDE_RECENTS" "$REMOVE_CSV" "$ADD_CSV"
$gen_mobileconfig_py
PY
  echo "Wrote export: $OUT_PATH"
fi

echo "Applied one-time Dock changes."
