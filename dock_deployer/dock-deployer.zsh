#!/bin/zsh
# Dock Deployer
# Default: apply Dock edits once to the current user's Dock (non-persistent)
# Optional: --forced to install as a managed profile (persistent)
# Optional: --reset to reset Dock to macOS defaults before applying edits

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

Options:
  --out <path>                       Output .mobileconfig path (default: ~/Desktop/Dock_export-{hostname}.mobileconfig)
  --remove <comma-separated labels>  Remove Dock items by label (e.g., "Maps,News")
  --add <comma-separated app paths>  Add apps by full .app paths (e.g., "/Applications/Google Chrome.app,/Applications/Slack.app")
  --hide-recents                     Set show-recents=false
  --forced                           Install as a persistent managed profile
  --uninstall                        Uninstall the profile file at --out
  --reset                            Reset Dock to macOS defaults before applying changes
  --help                             Show this help

Examples:
  # One-time edit (default):
  ./dock-deployer.zsh --remove "Maps,News" --add "/Applications/Slack.app" --hide-recents

  # Reset Dock to defaults, then edit:
  ./dock-deployer.zsh --reset --remove "Maps,News" --add "/Applications/Slack.app" --hide-recents

  # Force a persistent managed profile:
  ./dock-deployer.zsh --remove "Maps" --add "/Applications/Slack.app" --hide-recents --forced
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

# -------- Parse args --------
while [[ $# -gt 0 ]]; do
  case "$1" in
    --out) OUT_PATH="$2"; shift 2 ;;
    --remove) IFS=',' read -rA REMOVE_LABELS <<< "$2"; shift 2 ;;
    --add) IFS=',' read -rA ADD_APPS <<< "$2"; shift 2 ;;
    --hide-recents) HIDE_RECENTS=true; shift ;;
    --forced) FORCED=true; shift ;;
    --uninstall) UNINSTALL=true; shift ;;
    --reset) RESET=true; shift ;;
    --help|-h) print_usage; exit 0 ;;
    *) echo "Unknown option: $1" >&2; print_usage; exit 1 ;;
  esac
done

REMOVE_CSV="${(j:,: )REMOVE_LABELS}"
ADD_CSV="${(j:,: )ADD_APPS}"

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

# -------- Mode: Forced (persistent) --------
if $FORCED; then
  /usr/bin/python3 - <<'PY' "$OUT_PATH" "$HIDE_RECENTS" "$REMOVE_CSV" "$ADD_CSV"
import os, sys, uuid, plistlib, datetime, subprocess, re

out_path, hide_recents_s, remove_csv, add_csv = sys.argv[1:5]
hide_recents = hide_recents_s.lower() == "true"
remove_labels = [s.strip() for s in (remove_csv.split(",") if remove_csv else []) if s.strip()]
add_apps = [s.strip() for s in (add_csv.split(",") if add_csv else []) if s.strip()]

src = os.path.expanduser("~/Library/Preferences/com.apple.dock.plist")
with open(src, "rb") as f: dock = plistlib.loads(f.read())

labels_lower = {l.lower() for l in remove_labels}
def is_match(item):
    td = item.get("tile-data", {})
    lbl = (td.get("file-label") or "").lower()
    if lbl in labels_lower: return True
    bid = (td.get("bundle-identifier") or "").lower()
    if "maps" in labels_lower and bid == "com.apple.maps": return True
    if "news" in labels_lower and bid == "com.apple.news": return True
    path = (td.get("file-data", {}).get("_CFURLString") or "").lower()
    m = re.search(r"/applications/([^/]+)\.app", path or "")
    return m and m.group(1).replace("-", " ").lower() in labels_lower

if "persistent-apps" in dock and labels_lower:
    dock["persistent-apps"] = [i for i in dock["persistent-apps"] if not is_match(i)]

def make_tile(app_path):
    url = "file://" + app_path
    label = os.path.splitext(os.path.basename(app_path))[0]
    tile = {"tile-data": {"file-data": {"_CFURLString": url, "_CFURLStringType": 15},"file-label": label}, "tile-type": "file-tile"}
    try:
        bid = subprocess.check_output(["mdls","-name","kMDItemCFBundleIdentifier","-r",app_path],stderr=subprocess.DEVNULL,text=True).strip()
        if bid and bid!="(null)": tile["tile-data"]["bundle-identifier"]=bid
    except: pass
    return tile

def already(app_path):
    ap=("file://"+app_path).lower()
    for i in dock.get("persistent-apps", []):
        path=(i.get("tile-data", {}).get("file-data", {}).get("_CFURLString") or "").lower()
        if path==ap: return True
    return False

for p in add_apps:
    if not p.endswith(".app"): continue
    if not os.path.exists(p):
        for c in [f"/Applications/{p}", f"/System/Applications/{p}"]:
            if os.path.exists(c): p=c; break
        else: continue
    dock.setdefault("persistent-apps", [])
    if not already(p): dock["persistent-apps"].append(make_tile(p))

if hide_recents: dock["show-recents"]=False

wanted=["persistent-apps","persistent-others","show-recents","autohide","magnification","tilesize","largesize","mru-spaces"]
dock_subset={k:dock[k] for k in wanted if k in dock}

profile_uuid=str(uuid.uuid4()).upper(); payload_uuid=str(uuid.uuid4()).upper()
payload={"PayloadType":"com.apple.ManagedClient.preferences","PayloadVersion":1,"PayloadIdentifier":f"com.example.dock.custom.{payload_uuid}","PayloadUUID":payload_uuid,"PayloadEnabled":True,"PayloadDisplayName":"Dock (Custom Settings)","PayloadContent":{"com.apple.dock":{"Forced":[{"mcx_preference_settings":dock_subset}]}}}
mobileconfig={"PayloadType":"Configuration","PayloadVersion":1,"PayloadIdentifier":f"com.example.dock.export.{profile_uuid}","PayloadUUID":profile_uuid,"PayloadDisplayName":f"Dock Deploy ({datetime.datetime.now().strftime('%Y-%m-%d')})","PayloadRemovalDisallowed":False,"PayloadContent":[payload]}
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path,"wb") as f: plistlib.dump(mobileconfig,f,fmt=plistlib.FMT_XML)
print(out_path)
PY
  echo "Installing profile (sudo)..."
  sudo profiles -I -F "$OUT_PATH"
  killall Dock || true
  exit 0
fi

# -------- Mode: Default (set once, non-persistent) --------
/usr/bin/python3 - <<'PY' "$HIDE_RECENTS" "$REMOVE_CSV" "$ADD_CSV"
import os, sys, plistlib, subprocess, re

hide_recents=sys.argv[1].lower()=="true"
remove_csv=sys.argv[2]; add_csv=sys.argv[3]
remove_labels=[s.strip() for s in (remove_csv.split(",") if remove_csv else []) if s.strip()]
add_apps=[s.strip() for s in (add_csv.split(",") if add_csv else []) if s.strip()]

plist_path=os.path.expanduser("~/Library/Preferences/com.apple.dock.plist")
with open(plist_path,"rb") as f: d=plistlib.loads(f.read())

labels_lower={l.lower() for l in remove_labels}
def is_match(item):
    td=item.get("tile-data",{})
    lbl=(td.get("file-label") or "").lower()
    if lbl in labels_lower: return True
    bid=(td.get("bundle-identifier") or "").lower()
    if "maps" in labels_lower and bid=="com.apple.maps": return True
    if "news" in labels_lower and bid=="com.apple.news": return True
    path=(td.get("file-data",{}).get("_CFURLString") or "").lower()
    m=re.search(r"/applications/([^/]+)\.app", path or "")
    return m and m.group(1).replace("-", " ").lower() in labels_lower

if "persistent-apps" in d and labels_lower:
    d["persistent-apps"]=[i for i in d["persistent-apps"] if not is_match(i)]

def make_tile(app_path):
    url="file://"+app_path; label=os.path.splitext(os.path.basename(app_path))[0]
    tile={"tile-data":{"file-data":{"_CFURLString":url,"_CFURLStringType":15},"file-label":label},"tile-type":"file-tile"}
    try:
        bid=subprocess.check_output(["mdls","-name","kMDItemCFBundleIdentifier","-r",app_path],stderr=subprocess.DEVNULL,text=True).strip()
        if bid and bid!="(null)": tile["tile-data"]["bundle-identifier"]=bid
    except: pass
    return tile

def already(app_path):
    ap=("file://"+app_path).lower()
    for i in d.get("persistent-apps",[]):
        path=(i.get("tile-data",{}).get("file-data",{}).get("_CFURLString") or "").lower()
        if path==ap: return True
    return False

for p in add_apps:
    if not p.endswith(".app"): continue
    if not os.path.exists(p):
        for c in [f"/Applications/{p}", f"/System/Applications/{p}"]:
            if os.path.exists(c): p=c; break
        else: continue
    d.setdefault("persistent-apps",[])
    if not already(p): d["persistent-apps"].append(make_tile(p))

if hide_recents: d["show-recents"]=False

with open(plist_path,"wb") as f: f.write(plistlib.dumps(d,fmt=plistlib.FMT_BINARY))
subprocess.run(["killall","Dock"],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
print("Applied one-time Dock changes and restarted Dock.")
PY
