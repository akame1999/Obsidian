"""
OpsCenter — Windows Patch Management Dashboard
Simple, direct patch deployment for Windows Servers.
"""

import streamlit as st
import subprocess, os, json, re, tempfile, datetime, logging, threading
import requests, urllib3
import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OpsCenter | Patch Management",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.main { background: #080b12; }
section[data-testid="stSidebar"] { background: #0a0d18 !important; border-right: 1px solid #161d30; }

.block-container { padding-top: 0.4rem !important; padding-bottom: 0.3rem !important; }
div[data-testid="stTabContent"] { padding-top: 8px !important; }
hr { margin: 8px 0 !important; }
.stCaption { margin-bottom: 2px !important; }

.section-title {
    font-size: 9px; font-weight: 700; letter-spacing: 0.12em;
    text-transform: uppercase; color: #3b82f6;
    padding: 4px 0; border-bottom: 1px solid #161d30;
    margin-top: 5px; margin-bottom: 7px;
}
.info-box  { background:#0a1628; border:1px solid #1e3a5f; border-left:3px solid #3b82f6; border-radius:7px; padding:8px 13px; font-size:12px; color:#93c5fd; margin:4px 0 8px 0; }
.warn-box  { background:#1e1608; border:1px solid #78350f; border-left:3px solid #f59e0b; border-radius:7px; padding:8px 13px; font-size:12px; color:#fcd34d; margin:4px 0 8px 0; }
.ok-box    { background:#081a10; border:1px solid #065f46; border-left:3px solid #10b981; border-radius:7px; padding:8px 13px; font-size:12px; color:#6ee7b7; margin:4px 0 8px 0; }
.err-box   { background:#1a0808; border:1px solid #7f1d1d; border-left:3px solid #ef4444; border-radius:7px; padding:8px 13px; font-size:12px; color:#fca5a5; margin:4px 0 8px 0; }

div[data-testid="metric-container"] { background:#0d1120; border:1px solid #161d30; border-radius:8px; padding:6px 10px !important; }
div[data-testid="stMetricValue"] { font-size:1.2rem !important; font-weight:800 !important; color:#e2e8f0 !important; }
div[data-testid="stMetricLabel"] { font-size:9px !important; letter-spacing:0.1em !important; text-transform:uppercase !important; color:#4a5568 !important; }

button[data-baseweb="tab"] { font-size:11px !important; font-weight:600 !important; padding:4px 9px !important; }
.stButton > button { border-radius:6px !important; font-weight:600 !important; font-size:11px !important; padding: 3px 8px !important; }
div[data-testid="stDataFrame"] { margin-top:4px; margin-bottom:4px; border-radius:8px; overflow:hidden; }

/* Server cards */
.srv-card { background:#0d1120; border:1.5px solid #161d30; border-radius:9px; padding:8px 11px; margin-bottom:5px; }
.srv-card-online  { border-left:3px solid #10b981 !important; }
.srv-card-offline { border-left:3px solid #ef4444 !important; opacity:0.75; }
.srv-card-warn    { border-left:3px solid #f59e0b !important; }
.srv-card-header  { display:flex; align-items:center; gap:7px; margin-bottom:5px; }
.srv-dot-on  { width:6px; height:6px; border-radius:50%; background:#10b981; box-shadow:0 0 5px #10b981; flex-shrink:0; }
.srv-dot-off { width:6px; height:6px; border-radius:50%; background:#ef4444; flex-shrink:0; }
.srv-dot-unk { width:6px; height:6px; border-radius:50%; background:#4a5568; flex-shrink:0; }
.srv-hostname { font-family:'IBM Plex Mono',monospace; font-size:12px; font-weight:700; color:#e2e8f0; flex:1; }
.srv-time  { font-size:9px; color:#4a5568; }
.srv-grid  { display:grid; grid-template-columns:1fr 1fr; gap:4px 12px; }
.srv-kv    { display:flex; flex-direction:column; }
.srv-k     { font-size:8px; color:#4a5568; text-transform:uppercase; letter-spacing:0.08em; }
.srv-v     { font-size:10px; color:#94a3b8; font-weight:600; margin-top:1px; }
.srv-v-warn { color:#fbbf24 !important; }
.srv-v-ok   { color:#34d399 !important; }
.srv-badge  { display:inline-block; font-size:9px; font-weight:700; padding:2px 7px; border-radius:20px; text-transform:uppercase; letter-spacing:0.06em; }
.sb-on   { background:#081a10; color:#34d399; border:1px solid #065f46; }
.sb-off  { background:#1a0808; color:#fca5a5; border:1px solid #7f1d1d; }
.sb-warn { background:#1e1608; color:#fbbf24; border:1px solid #78350f; }
.sb-gray { background:#111827; color:#6b7280; border:1px solid #1e2a45; }

/* Patch cards */
.patch-card { background:#0d1120; border:1.5px solid #161d30; border-radius:8px; padding:6px 10px; margin:2px 0 3px 0; display:flex; align-items:center; gap:9px; }
.patch-card-sel { border-color:#3b82f6 !important; background:#08122a !important; }
.patch-kb   { font-size:12px; font-weight:800; color:#60a5fa; font-family:'IBM Plex Mono',monospace; min-width:75px; }
.patch-fname { font-size:9px; color:#64748b; margin-top:1px; }
.badge-pill { font-size:9px; font-weight:700; padding:2px 6px; border-radius:12px; text-transform:uppercase; }
.badge-sz   { background:#081a10; color:#34d399; border:1px solid #065f46; }
.badge-tp   { background:#0a1628; color:#60a5fa; border:1px solid #1e3a5f; }

/* Pipeline */
.pipe-hdr { display:grid; grid-template-columns:150px 100px 100px 100px 1fr; gap:4px; padding:5px 0 7px 0; border-bottom:2px solid #161d30; font-size:9px; font-weight:700; color:#4a5568; text-transform:uppercase; letter-spacing:0.1em; }
.pipe-row { display:grid; grid-template-columns:150px 100px 100px 100px 1fr; gap:4px; align-items:center; padding:6px 0; border-bottom:1px solid #0f1520; }
.pipe-srv  { font-family:'IBM Plex Mono',monospace; font-size:12px; font-weight:700; color:#cbd5e1; }
.pc { text-align:center; border-radius:5px; padding:3px 5px; font-size:10px; font-weight:700; }
.pc-ok   { background:#081a10; color:#34d399; }
.pc-fail { background:#1a0808; color:#fca5a5; }
.pc-skip { background:#1e1608; color:#fbbf24; }
.pc-wait { background:#0d1120; color:#4a5568; }
.pc-run  { background:#08122a; color:#60a5fa; }
.pipe-detail { font-size:10px; color:#64748b; overflow:hidden; white-space:nowrap; text-overflow:ellipsis; }

/* Deploy summary */
.dsb { background:#08122a; border:1.5px solid #1e3a5f; border-radius:9px; padding:10px 14px; display:flex; align-items:center; gap:8px; margin:6px 0 8px 0; flex-wrap:wrap; }
.dsb-kb  { font-weight:800; color:#60a5fa; font-family:'IBM Plex Mono',monospace; font-size:13px; }
.dsb-arrow { color:#3b82f6; font-size:15px; }
.dsb-srv { font-weight:600; color:#e2e8f0; font-size:12px; }
.dsb-sep { color:#1e3a5f; }
.dsb-meta { color:#4a5568; font-size:11px; }

    /* Server tiles for deploy tab */
    .srv-tile { border-radius:8px; padding:8px 12px; margin-bottom:4px; font-family:'IBM Plex Mono',monospace; }
    .srv-tile-on  { background:#08122a; border:1.5px solid #3b82f6; }
    .srv-tile-off { background:#0d1120; border:1.5px solid #161d30; opacity:0.55; }
    .srv-name { font-size:12px; font-weight:700; color:#e2e8f0; }
    .srv-sub  { font-size:10px; color:#6b7280; margin-top:2px; }
    .srv-sub-on { color:#34d399 !important; }

</style>
""", unsafe_allow_html=True)

# ─── PATHS & LOGGING ──────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
LOG_DIR  = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    filename=LOG_DIR / f'opcenter_{datetime.datetime.now():%Y%m%d}.log',
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
audit = logging.getLogger("audit")
audit.addHandler(logging.FileHandler(LOG_DIR / "audit.log"))
audit.setLevel(logging.INFO)

# ─── PORTABLE CONFIG (opcenter_config.json) ───────────────────────────────────
CONFIG_FILE = BASE_DIR / "opcenter_config.json"

DEFAULT_CONFIG = {
    "domain":             "ZL",
    "ps_timeout":         300,
    "patch_share":        r"\\10.87.60.2\IT Asset Detail\Naseer\Patch",
    "scan_interval_mins": 0,   # 0 = disabled; set to e.g. 60 for hourly auto-scans
    "servers": [
        "L11SGRIFP001",
        "L11SGRIFP002",
        "L11SGRIFP003",
        "L11SGRIFP005",
        "L11SGRIVMHDC01",
        "L11SGRIWEB001"
    ]
}

def load_config() -> dict:
    """Load config from JSON file, creating it with defaults if missing."""
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            # Ensure all keys exist (handles older config files)
            for k, v in DEFAULT_CONFIG.items():
                data.setdefault(k, v)
            return data
        except Exception as e:
            logging.warning(f"Config load error: {e}, using defaults")
    save_config(DEFAULT_CONFIG)
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    """Persist config to JSON file."""
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        logging.info("Config saved")
    except Exception as e:
        logging.error(f"Config save error: {e}")

# Load live config (re-read on every rerun so settings tab changes take effect)
_cfg = load_config()

class CFG:
    DOMAIN             = _cfg["domain"]
    PS_TIMEOUT         = _cfg["ps_timeout"]
    PATCH_SHARE        = _cfg["patch_share"]
    SERVERS            = _cfg["servers"]
    SCAN_INTERVAL_MINS = _cfg.get("scan_interval_mins", 0)
    SCRIPT_DIR         = BASE_DIR
    PS1_SCRIPT         = BASE_DIR / "Invoke-ModernPatch.ps1"

if not CFG.PS1_SCRIPT.exists():
    st.error(f"❌ Script not found: {CFG.PS1_SCRIPT}")
    st.stop()

# ─── POWERSHELL RUNNER ────────────────────────────────────────────────────────
def run_ps(action: str, username: str, password: str,
           server: str = None, patch: str = None,
           reboot_time: str = None, kb_number: str = None) -> Dict:
    if not username or not password:
        return {"success": False, "error": "Credentials required"}

    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.tmp') as f:
        f.write(password)
        cred_file = f.name

    tmp_script = None
    try:
        clean_user  = username.replace(f'{CFG.DOMAIN}\\', '').strip()
        full_user   = f"{CFG.DOMAIN}\\{clean_user}"
        script_path = str(CFG.PS1_SCRIPT).replace('\\', '\\\\')
        cred_ps     = cred_file.replace('\\', '\\\\')

        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.ps1', encoding='utf-8') as ps:
            ps.write(f"""
$password = Get-Content '{cred_ps}' -Raw
$securePass = ConvertTo-SecureString $password -AsPlainText -Force
Remove-Item '{cred_ps}' -Force -ErrorAction SilentlyContinue
$credential = New-Object System.Management.Automation.PSCredential('{full_user}', $securePass)
""")
            call = f"& '{script_path}' -Action '{action}' -Cred $credential"
            if server:      call += f" -TargetServer '{server}'"
            if patch:       call += f" -FileName '{patch.replace(chr(39), chr(39)*2)}'"
            if reboot_time: call += f" -RebootTime '{reboot_time}'"
            if kb_number:   call += f" -KBNumber '{kb_number}'"
            ps.write(call + "\nexit $LASTEXITCODE\n")
            tmp_script = ps.name

        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmp_script],
            capture_output=True, text=True, timeout=CFG.PS_TIMEOUT, cwd=str(CFG.SCRIPT_DIR)
        )

        out = r.stdout.strip()
        # Try JSON array first (e.g. GetInstalledPatches returns [...])
        a_s, a_e = out.find('['), out.rfind(']') + 1
        o_s, o_e = out.find('{'), out.rfind('}') + 1
        if a_s >= 0 and a_e > a_s and (o_s < 0 or a_s < o_s):
            data = json.loads(out[a_s:a_e])
            if isinstance(data, list):
                audit.info(f"{action} | Server={server} | User={username} | OK=True")
                return data
        if o_s >= 0 and o_e > o_s:
            data = json.loads(out[o_s:o_e])
            data['success'] = r.returncode == 0
            audit.info(f"{action} | Server={server} | User={username} | OK={data['success']}")
            return data

        return {"success": False, "error": r.stderr or "No response from server"}

    except subprocess.TimeoutExpired:
        return {"success": False, "error": f"Timed out after {CFG.PS_TIMEOUT}s — server may be busy"}
    except Exception as e:
        logging.error(f"PS error: {e}")
        return {"success": False, "error": str(e)}
    finally:
        for f in [cred_file, tmp_script]:
            if f and os.path.exists(f):
                try: os.remove(f)
                except: pass

def scan_patches() -> List[Dict]:
    patches = []
    try:
        for root, _, files in os.walk(CFG.PATCH_SHARE):
            for f in files:
                if f.lower().endswith(('.msu', '.exe', '.cab')):
                    path = os.path.join(root, f)
                    kb   = re.search(r'KB(\d+)', f, re.IGNORECASE)
                    patches.append({
                        'filename': f,
                        'kb':       f"KB{kb.group(1)}" if kb else "Unknown",
                        'type':     f[-4:].upper(),
                        'size_mb':  round(os.path.getsize(path) / 1024 / 1024, 2)
                    })
    except Exception as e:
        logging.error(f"Patch scan error: {e}")
    return sorted(patches, key=lambda x: x['filename'])

# ─── MS UPDATE CATALOG ENGINE ─────────────────────────────────────────────────
_CATALOG_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
_CATALOG_SESSION = requests.Session()
_CATALOG_SESSION.headers.update(_CATALOG_HEADERS)

def catalog_search(query: str) -> List[Dict]:
    """Search Microsoft Update Catalog and return list of update entries."""
    try:
        url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={requests.utils.quote(query)}"
        r = _CATALOG_SESSION.get(url, timeout=20, verify=False)
        r.raise_for_status()
        html = r.text

        # Extract each result row — rows have id="..." matching a GUID pattern
        rows = re.findall(
            r'<tr\s+id="([0-9a-f\-]{36}_[^"]+)"[^>]*>(.*?)</tr>',
            html, re.DOTALL | re.IGNORECASE
        )

        results = []
        for row_id, row_html in rows:
            cells = re.findall(r'<td[^>]*>(.*?)</td>', row_html, re.DOTALL)
            if len(cells) < 6:
                continue

            def clean(s):
                s = re.sub(r'<[^>]+>', '', s)
                return re.sub(r'\s+', ' ', s).strip()

            title       = clean(cells[1]) if len(cells) > 1 else ""
            products    = clean(cells[2]) if len(cells) > 2 else ""
            classif     = clean(cells[3]) if len(cells) > 3 else ""
            last_upd    = clean(cells[4]) if len(cells) > 4 else ""
            version     = clean(cells[5]) if len(cells) > 5 else ""
            # Clean size — Microsoft returns "X MB YYYYYYY" (MB + raw bytes)
            # Extract just the MB portion
            size_raw = clean(cells[6]) if len(cells) > 6 else ""
            size_match = re.match(r'([\d,\.]+\s*MB)', size_raw, re.IGNORECASE)
            size_str = size_match.group(1).strip() if size_match else size_raw.split()[0] if size_raw else ""

            # Extract update UID (the part before the underscore)
            uid = row_id.split('_')[0] if '_' in row_id else row_id

            # Extract KB number from title
            kb_match = re.search(r'KB\d+', title, re.IGNORECASE)
            kb = kb_match.group(0).upper() if kb_match else ""

            if title:
                results.append({
                    "uid":        uid,
                    "title":      title,
                    "kb":         kb,
                    "products":   products,
                    "classif":    classif,
                    "date":       last_upd,
                    "version":    version,
                    "size":       size_str,
                })

        return results

    except requests.exceptions.ConnectionError:
        return [{"error": "Cannot reach catalog.update.microsoft.com — check internet connection"}]
    except Exception as e:
        logging.error(f"Catalog search error: {e}")
        return [{"error": str(e)}]


def catalog_get_download_url(uid: str) -> Optional[str]:
    """Get the direct download URL for a given update UID."""
    try:
        payload = (
            f'updateIDs=[{{"size":0,"languages":"","uidInfo":"{uid}","updateID":"{uid}"}}]'
            f'&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku='
            f'&serverName=&ssl=&portNumber=&version='
        )
        r = _CATALOG_SESSION.post(
            "https://www.catalog.update.microsoft.com/DownloadDialog.aspx",
            data=payload,
            headers={**_CATALOG_HEADERS, "Content-Type": "application/x-www-form-urlencoded"},
            timeout=15,
            verify=False
        )
        # Extract download URL from the returned JS
        match = re.search(r"downloadInformation\[0\]\.files\[0\]\.url\s*=\s*'([^']+)'", r.text)
        if match:
            return match.group(1)
        # Fallback pattern
        match2 = re.search(r'https://[^\s\'"]+\.(?:msu|exe|cab)', r.text, re.IGNORECASE)
        return match2.group(0) if match2 else None
    except Exception as e:
        logging.error(f"Get download URL error: {e}")
        return None


def download_kb_to_share(uid: str, filename: str, dest_folder: str,
                          progress_callback=None) -> Dict:
    """Download a KB file from Microsoft to the patch share folder."""
    try:
        url = catalog_get_download_url(uid)
        if not url:
            return {"success": False, "error": "Could not retrieve download URL from Microsoft"}

        os.makedirs(dest_folder, exist_ok=True)
        dest_path = os.path.join(dest_folder, filename)

        r = _CATALOG_SESSION.get(url, stream=True, timeout=30, verify=False)
        r.raise_for_status()

        total = int(r.headers.get("content-length", 0))
        downloaded = 0
        with open(dest_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 256):  # 256 KB chunks
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total:
                        progress_callback(downloaded / total)

        size_mb = round(os.path.getsize(dest_path) / 1024 / 1024, 2)
        logging.info(f"Downloaded {filename} ({size_mb} MB) to {dest_folder}")
        return {"success": True, "path": dest_path, "size_mb": size_mb}

    except Exception as e:
        logging.error(f"Download error: {e}")
        return {"success": False, "error": str(e)}

# ─── SESSION STATE ────────────────────────────────────────────────────────────
for k, v in {
    'authenticated': False, 'username': '', 'password': '',
    'login_time': None, 'history': [], 'patches': [], 'scan_results': [],
    'installed_cache': {},
    'kb_results': [],
    'kb_raw_results': [],
    'kb_downloads': [],
    'kb_queue': [],
    'auto_scan_pending': False,
    'last_auto_scan': None,
    'next_auto_scan': None,
    'pipeline_results': [],
    'deploy_patch_sel': {},
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🛡️ OpsCenter")
    st.markdown('<div style="font-size:11px;color:#4a5568;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:16px;">Patch Management</div>', unsafe_allow_html=True)
    st.divider()

    if not st.session_state.authenticated:
        st.markdown("**Sign in to continue**")
        st.caption("Use your ZL domain account (e.g. oine3994100)")
        with st.form("login"):
            user = st.text_input("Username", placeholder="oine3994100")
            pwd  = st.text_input("Password", type="password")
            if st.form_submit_button("🔐 Sign In", type="primary", use_container_width=True):
                if user and pwd:
                    with st.spinner("Checking credentials…"):
                        result = run_ps("PreCheck", user, pwd, CFG.SERVERS[0])
                    if result.get('success'):
                        st.session_state.update(
                            authenticated=True, username=user,
                            password=pwd, login_time=datetime.datetime.now(),
                            auto_scan_pending=True   # trigger auto-scan on first render
                        )
                        audit.info(f"Login: {user}")
                        st.rerun()
                    else:
                        st.error("Wrong username or password")
                else:
                    st.warning("Please fill in both fields")
    else:
        st.success(f"Signed in\n**{CFG.DOMAIN}\\{st.session_state.username}**")
        mins = (datetime.datetime.now() - st.session_state.login_time).seconds // 60
        st.caption(f"Session active for {mins} min")
        st.divider()

        col1, col2 = st.columns(2)
        with col1: st.metric("Servers", len(CFG.SERVERS))
        with col2:
            ok = sum(1 for h in st.session_state.history if h.get('success'))
            st.metric("Deployed", ok)

        st.divider()
        st.markdown('<div class="section-title">How to use</div>', unsafe_allow_html=True)
        st.caption("📊 **Server Status** — Check which servers are online")
        st.caption("🚀 **Deploy a Patch** — Push a Windows update to servers")
        st.caption("📜 **History & Logs** — Review past actions")
        st.divider()

        if st.button("🚪 Sign Out", use_container_width=True):
            audit.info(f"Logout: {st.session_state.username}")
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

# ─── LOGIN GATE ───────────────────────────────────────────────────────────────
if not st.session_state.authenticated:
    st.markdown("""
    <div style="text-align:center; padding: 80px 0 20px 0;">
        <div style="font-size: 56px; margin-bottom: 16px;">🛡️</div>
        <div style="font-size: 30px; font-weight: 700; color: #e2e8f0; margin-bottom: 8px;">OpsCenter</div>
        <div style="font-size: 15px; color: #6b7280;">Windows Patch Management</div>
        <div style="margin-top: 32px; padding: 14px 24px; background:#0d1f3a;
                    border:1px solid #1e3a5f; border-radius:10px; display:inline-block;
                    font-size:13px; color:#93c5fd;">
            👈 Sign in with your ZL domain account to get started
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─── AUTO-SCAN ENGINE ─────────────────────────────────────────────────────────
def run_server_scan() -> list:
    """Run PreCheck on all servers and return rows list."""
    rows = []
    for srv in CFG.SERVERS:
        res = run_ps("PreCheck", st.session_state.username, st.session_state.password, srv)
        if res.get('success'):
            rows.append({
                "Server":          srv,
                "Status":          "🟢 Online",
                "Windows Version": res.get('OS', 'Unknown'),
                "Needs Restart?":  "⚠️ Yes — restart pending" if res.get('RebootRequired') == 'Yes' else "✅ No",
                "Last Checked":    datetime.datetime.now().strftime("%H:%M:%S")
            })
        else:
            rows.append({
                "Server":          srv,
                "Status":          "🔴 Offline",
                "Windows Version": "Could not connect",
                "Needs Restart?":  "—",
                "Last Checked":    datetime.datetime.now().strftime("%H:%M:%S")
            })
    return rows

# ── Auto-scan on first login ──────────────────────────────────────────────────
if st.session_state.get('auto_scan_pending'):
    st.session_state.auto_scan_pending = False
    with st.spinner("🔍 Auto-scanning all servers…"):
        st.session_state.scan_results = run_server_scan()
    st.session_state.last_auto_scan = datetime.datetime.now()
    # Set next scan time if interval configured
    if CFG.SCAN_INTERVAL_MINS > 0:
        st.session_state.next_auto_scan = (
            datetime.datetime.now() + datetime.timedelta(minutes=CFG.SCAN_INTERVAL_MINS)
        )

# ── Recurring scheduled scan ──────────────────────────────────────────────────
if (CFG.SCAN_INTERVAL_MINS > 0
        and st.session_state.next_auto_scan is not None
        and datetime.datetime.now() >= st.session_state.next_auto_scan):
    with st.spinner(f"🔄 Scheduled scan running (every {CFG.SCAN_INTERVAL_MINS} min)…"):
        st.session_state.scan_results = run_server_scan()
    st.session_state.last_auto_scan = datetime.datetime.now()
    st.session_state.next_auto_scan = (
        datetime.datetime.now() + datetime.timedelta(minutes=CFG.SCAN_INTERVAL_MINS)
    )
    logging.info(f"Scheduled scan completed. Next at {st.session_state.next_auto_scan:%H:%M}")

# ─── HEADER ───────────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([6, 1])
with col_h1:
    st.markdown(f'<div style="padding:2px 0 4px 0;"><span style="font-size:18px;font-weight:700;color:#e2e8f0;">🛡️ Windows Patch Management</span> <span style="font-size:11px;color:#4a5568;margin-left:10px;">{datetime.datetime.now():%a, %d %b %Y}</span></div>', unsafe_allow_html=True)
with col_h2:
    if st.button("↺ Refresh", use_container_width=True):
        st.rerun()

# ─── TABS ─────────────────────────────────────────────────────────────────────
tab_dash, tab_deploy, tab_history, tab_kb, tab_rollback, tab_settings = st.tabs([
    "📊  Server Status",
    "🚀  Deploy a Patch",
    "📜  History & Logs",
    "⬇️  KB Download",
    "↩️  Rollback",
    "⚙️  Settings"
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — SERVER STATUS
# ══════════════════════════════════════════════════════════════════════════════
with tab_dash:

    scan    = st.session_state.scan_results
    total   = len(CFG.SERVERS)
    online  = sum(1 for r in scan if "Online"  in r.get("Status",""))
    offline = sum(1 for r in scan if "Offline" in r.get("Status",""))
    reboot  = sum(1 for r in scan if "Yes"     in r.get("Needs Restart?",""))

    # ── Metric bar ────────────────────────────────────────────────────────────
    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
    with mc1: st.metric("Servers",  total)
    with mc2: st.metric("Online",   f"{online}/{total}" if scan else "—")
    with mc3: st.metric("Offline",  offline if scan else "—")
    with mc4: st.metric("Restart ↑",reboot  if scan else "—")
    with mc5: st.metric("Health",   f"{int(online/total*100)}%" if (scan and total) else "—")

    # Timing info
    info_parts = []
    if st.session_state.last_auto_scan:
        info_parts.append(f"🕐 Last scan: **{st.session_state.last_auto_scan:%H:%M:%S}**")
    if st.session_state.next_auto_scan and CFG.SCAN_INTERVAL_MINS > 0:
        delta     = st.session_state.next_auto_scan - datetime.datetime.now()
        mins_left = max(0, int(delta.total_seconds()//60))
        info_parts.append(f"🔄 Next in **{mins_left}m**")
    if info_parts:
        st.caption("  ·  ".join(info_parts))

    if st.button("🔍 Scan All Servers", type="primary", use_container_width=True, key="scan_all_btn"):
        bar  = st.progress(0, text="Scanning…")
        rows = []
        for i, srv in enumerate(CFG.SERVERS):
            bar.progress((i+1)/total, text=f"Checking {srv}… ({i+1}/{total})")
            res = run_ps("PreCheck", st.session_state.username, st.session_state.password, srv)
            if res.get("success"):
                rows.append({"Server": srv, "Status": "🟢 Online",
                             "Windows Version": res.get("OS","Unknown"),
                             "Needs Restart?":  "⚠️ Yes — restart pending" if res.get("RebootRequired")=="Yes" else "✅ No",
                             "Last Checked":    datetime.datetime.now().strftime("%H:%M:%S")})
            else:
                rows.append({"Server": srv, "Status": "🔴 Offline",
                             "Windows Version": "Could not connect",
                             "Needs Restart?":  "—",
                             "Last Checked":    datetime.datetime.now().strftime("%H:%M:%S")})
        bar.empty()
        st.session_state.scan_results     = rows
        st.session_state.last_auto_scan   = datetime.datetime.now()
        st.rerun()

    if not st.session_state.scan_results:
        st.markdown('<div class="info-box">👆 Click <strong>Scan All Servers</strong> to view live status, OS versions and patch state for every server.</div>', unsafe_allow_html=True)
    else:
        # Status banner
        if offline > 0 and reboot > 0:
            st.markdown(f'<div class="warn-box">⚠️ <strong>{offline} offline</strong> &nbsp;·&nbsp; <strong>{reboot} need restart</strong></div>', unsafe_allow_html=True)
        elif reboot > 0:
            st.markdown(f'<div class="warn-box">⚠️ <strong>{reboot} server(s)</strong> need a restart before patches take effect.</div>', unsafe_allow_html=True)
        elif offline > 0:
            st.markdown(f'<div class="err-box">❌ <strong>{offline} server(s) offline</strong> — check connectivity.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="ok-box">✅ All servers online · No restarts pending.</div>', unsafe_allow_html=True)

        # ── Two-column layout ─────────────────────────────────────────────────
        srv_col, patch_col = st.columns([5, 7], gap="medium")

        with srv_col:
            st.markdown('<div class="section-title">🖥️ Server Details</div>', unsafe_allow_html=True)
            df_exp = pd.DataFrame(st.session_state.scan_results)
            st.download_button("📥 Export CSV", df_exp.to_csv(index=False),
                               f"servers_{datetime.datetime.now():%Y%m%d_%H%M}.csv",
                               "text/csv", use_container_width=True, key="srv_exp")

            for row in st.session_state.scan_results:
                srv      = row["Server"]
                is_on    = "Online"  in row.get("Status","")
                is_off   = "Offline" in row.get("Status","")
                needs_rb = "Yes"     in row.get("Needs Restart?","")
                os_ver   = row.get("Windows Version","—")[:34]
                chk_t    = row.get("Last Checked","—")

                if is_off:
                    card_cls, dot_cls, badge_cls, badge_lbl = "srv-card srv-card-offline","srv-dot-off","sb-off","Offline"
                elif needs_rb:
                    card_cls, dot_cls, badge_cls, badge_lbl = "srv-card srv-card-warn",  "srv-dot-on", "sb-warn","Restart ↑"
                elif is_on:
                    card_cls, dot_cls, badge_cls, badge_lbl = "srv-card srv-card-online","srv-dot-on", "sb-on",  "Online"
                else:
                    card_cls, dot_cls, badge_cls, badge_lbl = "srv-card",                "srv-dot-unk","sb-gray","Unknown"

                rb_val_cls = "srv-v-warn" if needs_rb else "srv-v-ok"
                rb_val     = "⚠️ Pending" if needs_rb else "✅ Clear"

                st.markdown(f"""
                <div class="{card_cls}">
                    <div class="srv-card-header">
                        <div class="{dot_cls}"></div>
                        <span class="srv-hostname">{srv}</span>
                        <span class="srv-badge {badge_cls}">{badge_lbl}</span>
                        <span class="srv-time">{chk_t}</span>
                    </div>
                    <div class="srv-grid">
                        <div class="srv-kv">
                            <span class="srv-k">OS</span>
                            <span class="srv-v">{os_ver}</span>
                        </div>
                        <div class="srv-kv">
                            <span class="srv-k">Restart</span>
                            <span class="srv-v {rb_val_cls}">{rb_val}</span>
                        </div>
                    </div>
                </div>""", unsafe_allow_html=True)

        with patch_col:
            st.markdown('<div class="section-title">🔎 Installed Patches</div>', unsafe_allow_html=True)
            online_servers = [r["Server"] for r in scan if "Online" in r.get("Status","")]

            if not online_servers:
                st.caption("No online servers to query.")
            else:
                c_sel, c_btn = st.columns([4, 1])
                with c_sel:
                    chosen_server = st.selectbox("Server", online_servers,
                                                 label_visibility="collapsed", key="patch_inspect_server")
                with c_btn:
                    load_btn = st.button("📋 Load", use_container_width=True, key="load_installed")

                if load_btn and chosen_server:
                    with st.spinner(f"Loading {chosen_server}…"):
                        res = run_ps("GetInstalledPatches", st.session_state.username,
                                     st.session_state.password, server=chosen_server)
                    st.session_state.installed_cache[chosen_server] = res

                cached = st.session_state.installed_cache.get(chosen_server)
                if cached is not None:
                    if isinstance(cached, list) and len(cached) > 0:
                        df_inst = pd.DataFrame(cached)
                        df_inst = df_inst.rename(columns={"HotFixID":"KB","InstalledOn":"Installed","Description":"Type"})
                        if "Installed" in df_inst.columns:
                            df_inst = df_inst.sort_values("Installed", ascending=False)

                        srch = st.text_input("🔍 Filter", placeholder="KB5034441 or Security",
                                             key="patch_filter", label_visibility="collapsed")
                        if srch:
                            mask = df_inst.apply(lambda r: r.astype(str).str.contains(srch, case=False).any(), axis=1)
                            df_inst = df_inst[mask]

                        st.markdown(f'<div class="ok-box">✅ <strong>{len(df_inst)}</strong> patch(es) on <strong>{chosen_server}</strong></div>', unsafe_allow_html=True)
                        st.dataframe(df_inst, use_container_width=True, hide_index=True,
                                     column_config={"KB": st.column_config.TextColumn("KB", width="small"),
                                                    "Type": st.column_config.TextColumn("Type", width="medium"),
                                                    "Installed": st.column_config.TextColumn("Installed", width="small")})
                        st.download_button(f"📥 Export {chosen_server}",
                            df_inst.to_csv(index=False),
                            f"patches_{chosen_server}_{datetime.datetime.now():%Y%m%d}.csv",
                            "text/csv", key="dl_installed", use_container_width=True)
                    elif isinstance(cached, dict) and not cached.get("success", True):
                        st.markdown(f'<div class="err-box">❌ {cached.get("error","Unknown error")}</div>', unsafe_allow_html=True)
                    else:
                        st.caption("No patches found on this server.")
                else:
                    st.caption("Select a server above and click **Load**.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — DEPLOY  (Bulk Patches · Pre-Check · Pipeline)
# ══════════════════════════════════════════════════════════════════════════════
with tab_deploy:

    # init state
    if 'deploy_patch_sel'  not in st.session_state: st.session_state.deploy_patch_sel  = {}
    if 'deploy_srv_sel'    not in st.session_state: st.session_state.deploy_srv_sel    = {s:True for s in CFG.SERVERS}
    if 'pipeline_results'  not in st.session_state: st.session_state.pipeline_results  = []

    # auto-load patches
    if not st.session_state.patches:
        with st.spinner("Loading patch files…"):
            st.session_state.patches = scan_patches()

    patch_names = [p['filename'] for p in st.session_state.patches]
    patch_map   = {p['filename']: p for p in st.session_state.patches}

    # sync patch selection state with current file list
    for n in patch_names:
        if n not in st.session_state.deploy_patch_sel:
            st.session_state.deploy_patch_sel[n] = False
    if patch_names and not any(st.session_state.deploy_patch_sel.values()):
        st.session_state.deploy_patch_sel[patch_names[0]] = True

    status_map = {r["Server"]: r for r in st.session_state.scan_results}

    # ── Two-column layout: LEFT = config  RIGHT = live pipeline ──────────────
    cfg_col, pipe_col = st.columns([4, 6], gap="large")

    # ════════════════════════════════
    # LEFT PANEL — Patch + Server pick
    # ════════════════════════════════
    with cfg_col:

        # — PATCHES —
        hc1, hc2 = st.columns([4,1])
        with hc1: st.markdown('<div class="section-title">📦 Patches to Deploy</div>', unsafe_allow_html=True)
        with hc2:
            if st.button("🔄", key="dep_ref", help="Refresh patch list"):
                st.session_state.patches = scan_patches()
                st.session_state.deploy_patch_sel = {}
                st.rerun()

        if not patch_names:
            st.markdown('<div class="warn-box">⚠️ No patches found in share — use <strong>KB Download</strong> tab first.</div>', unsafe_allow_html=True)
        else:
            pa, pn, _ = st.columns([1,1,3])
            with pa:
                if st.button("☑ All", key="dep_pall", use_container_width=True):
                    st.session_state.deploy_patch_sel = {n:True  for n in patch_names}; st.rerun()
            with pn:
                if st.button("✕ None", key="dep_pnone", use_container_width=True):
                    st.session_state.deploy_patch_sel = {n:False for n in patch_names}; st.rerun()

            for fname in patch_names:
                p   = patch_map[fname]
                sel = st.session_state.deploy_patch_sel.get(fname, False)
                cc  = "patch-card patch-card-sel" if sel else "patch-card"
                st.markdown(f"""
                <div class="{cc}">
                    <div class="patch-kb">{p['kb']}</div>
                    <div>
                        <div class="patch-fname">{fname}</div>
                        <div style="display:flex;gap:5px;margin-top:3px;">
                            <span class="badge-pill badge-sz">{p['size_mb']} MB</span>
                            <span class="badge-pill badge-tp">{p['type']}</span>
                        </div>
                    </div>
                </div>""", unsafe_allow_html=True)
                nv = st.checkbox("Include", value=sel, key=f"dp_{fname}", label_visibility="collapsed")
                if nv != sel:
                    st.session_state.deploy_patch_sel[fname] = nv; st.rerun()

        # — SERVERS —
        st.markdown('<div class="section-title">🖥️ Target Servers</div>', unsafe_allow_html=True)

        sa, sn, _ = st.columns([1,1,3])
        with sa:
            if st.button("☑ All", key="dep_sall", use_container_width=True):
                st.session_state.deploy_srv_sel = {s:True  for s in CFG.SERVERS}; st.rerun()
        with sn:
            if st.button("✕ None", key="dep_snone", use_container_width=True):
                st.session_state.deploy_srv_sel = {s:False for s in CFG.SERVERS}; st.rerun()

        srv_chunks = [CFG.SERVERS[i:i+2] for i in range(0, len(CFG.SERVERS), 2)]
        for chunk in srv_chunks:
            cols = st.columns(2)
            for ci, srv in enumerate(chunk):
                with cols[ci]:
                    row      = status_map.get(srv, {})
                    is_on    = "Online"  in row.get("Status","")
                    is_off   = "Offline" in row.get("Status","")
                    needs_rb = "Yes"     in row.get("Needs Restart?","")
                    dot      = "🟢" if is_on else ("🔴" if is_off else "⚪")
                    sub      = ("Restart ↑" if needs_rb else "Online") if is_on else ("Offline" if is_off else "Not scanned")
                    state    = st.session_state.deploy_srv_sel.get(srv, True)
                    tc       = "srv-tile srv-tile-on"  if state else "srv-tile srv-tile-off"
                    sc       = "srv-sub srv-sub-on"    if is_on else "srv-sub"
                    st.markdown(f"""<div class="{tc}">
                        <div class="srv-name">{dot} {srv}</div>
                        <div class="{sc}">{sub}</div>
                    </div>""", unsafe_allow_html=True)
                    nv = st.checkbox("Include", value=state, key=f"ds_{srv}", label_visibility="collapsed")
                    if nv != state:
                        st.session_state.deploy_srv_sel[srv] = nv; st.rerun()

    # ════════════════════════════════
    # RIGHT PANEL — Pipeline
    # ════════════════════════════════
    with pipe_col:
        st.markdown('<div class="section-title">🤖 Automated Pipeline</div>', unsafe_allow_html=True)

        selected_patches = [n for n, v in st.session_state.deploy_patch_sel.items() if v]
        selected_servers = [s for s, v in st.session_state.deploy_srv_sel.items()    if v]
        can_run          = bool(selected_patches and selected_servers)

        # Summary bar
        if can_run:
            kbs       = " + ".join(patch_map[n]['kb'] for n in selected_patches)
            total_ops = len(selected_patches) * len(selected_servers)
            st.markdown(f"""
            <div class="dsb">
                <span class="dsb-kb">{kbs}</span>
                <span class="dsb-arrow">→</span>
                <span class="dsb-srv">{len(selected_servers)} server(s)</span>
                <span class="dsb-sep">|</span>
                <span class="dsb-meta">{len(selected_patches)} patch(es) · {total_ops} op(s) total</span>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown('<div class="info-box">👈 Select patches and servers on the left to begin.</div>', unsafe_allow_html=True)

        # Mode + options
        mode = st.radio("Mode", ["🔍 Pre-Check Only", "🚀 Full Auto  (Pre-Check → Deploy → Verify)"],
                        horizontal=True, label_visibility="collapsed", key="deploy_mode")
        full_auto = "Full Auto" in mode
        skip_reboot = False
        if full_auto:
            skip_reboot = st.toggle("⛔ Skip servers with pending reboot", value=True, key="dep_skip_rb")

        run_btn = st.button("▶  Run Pipeline", type="primary", use_container_width=True,
                            disabled=not can_run, key="dep_run")

        # Pipeline table header
        st.markdown("""
        <div class="pipe-hdr">
            <span>Server</span><span style="text-align:center">Pre-Check</span>
            <span style="text-align:center">Deploy</span><span style="text-align:center">Verify</span>
            <span>Detail</span>
        </div>""", unsafe_allow_html=True)

        # Build live row placeholders for each server
        row_ph = {}
        def blank_row(srv):
            return f"""<div class="pipe-row">
                <span class="pipe-srv">{srv}</span>
                <span class="pc pc-wait">—</span><span class="pc pc-wait">—</span>
                <span class="pc pc-wait">—</span><span class="pipe-detail">Waiting…</span>
            </div>"""

        for srv in CFG.SERVERS:
            ph = st.empty()
            row_ph[srv] = ph
            ph.markdown(blank_row(srv), unsafe_allow_html=True)

        overall_ph = st.empty()

        # ── Execute pipeline ──────────────────────────────────────────────────
        if run_btn and can_run:
            st.session_state.pipeline_results = []

            def rrow(srv, pre, dep, ver, detail, pc="pc-wait", dc="pc-wait", vc="pc-wait"):
                row_ph[srv].markdown(f"""<div class="pipe-row">
                    <span class="pipe-srv">{srv}</span>
                    <span class="pc {pc}">{pre}</span><span class="pc {dc}">{dep}</span>
                    <span class="pc {vc}">{ver}</span>
                    <span class="pipe-detail">{detail}</span>
                </div>""", unsafe_allow_html=True)

            for p_idx, patch_file in enumerate(selected_patches):
                kb = patch_map[patch_file]['kb']
                overall_ph.markdown(
                    f'<div class="info-box">⏳ Patch {p_idx+1}/{len(selected_patches)}: '
                    f'<strong>{kb}</strong> — running on {len(selected_servers)} server(s)…</div>',
                    unsafe_allow_html=True)

                for srv in selected_servers:

                    # ── Phase 1: Pre-Check ─────────────────────────────────
                    rrow(srv, "🔍…","—","—", f"Pre-checking {kb}…", pc="pc-run")
                    pre       = run_ps("PreCheck", st.session_state.username,
                                       st.session_state.password, server=srv)
                    pre_ok    = pre.get('success', False)
                    os_ver    = pre.get('OS','?')
                    needs_rb  = pre.get('RebootRequired','') == 'Yes'

                    if not pre_ok:
                        rrow(srv,"❌ Offline","⏭ Skip","⏭ Skip","Server unreachable",
                             pc="pc-fail", dc="pc-skip", vc="pc-skip")
                        st.session_state.pipeline_results.append(
                            {'server':srv,'kb':kb,'patch':patch_file,
                             'precheck':'❌ Offline','deploy':'⏭ Skip','verify':'⏭ Skip',
                             'detail':'Server unreachable'})
                        continue

                    if needs_rb and skip_reboot and full_auto:
                        rrow(srv,"⚠️ Reboot↑","⏭ Skip","⏭ Skip",f"{os_ver} — skipped (reboot pending)",
                             pc="pc-skip", dc="pc-skip", vc="pc-skip")
                        st.session_state.pipeline_results.append(
                            {'server':srv,'kb':kb,'patch':patch_file,
                             'precheck':'⚠️ Reboot Pending','deploy':'⏭ Skip','verify':'⏭ Skip',
                             'detail':'Skipped — pending reboot'})
                        continue

                    pre_lbl = "⚠️ Reboot↑" if needs_rb else "✅ Ready"
                    pre_cls = "pc-skip"    if needs_rb else "pc-ok"

                    if not full_auto:
                        rrow(srv, pre_lbl,"—","—", os_ver, pc=pre_cls)
                        st.session_state.pipeline_results.append(
                            {'server':srv,'kb':kb,'patch':patch_file,
                             'precheck':pre_lbl,'deploy':'—','verify':'—','detail':os_ver})
                        continue

                    # ── Phase 2: Deploy ────────────────────────────────────
                    rrow(srv, pre_lbl,"🚀…","—",f"Installing {kb}…",
                         pc=pre_cls, dc="pc-run")
                    dep    = run_ps("Deploy", st.session_state.username,
                                    st.session_state.password, server=srv, patch=patch_file)
                    dep_ok = dep.get('success', False)
                    dep_msg= dep.get('Message', dep.get('error','—'))[:80]

                    if not dep_ok:
                        rrow(srv, pre_lbl,"❌ Failed","⏭ Skip", dep_msg,
                             pc=pre_cls, dc="pc-fail", vc="pc-skip")
                        st.session_state.pipeline_results.append(
                            {'server':srv,'kb':kb,'patch':patch_file,
                             'precheck':pre_lbl,'deploy':'❌ Failed','verify':'⏭ Skip',
                             'detail':dep_msg})
                        st.session_state.history.append(
                            {'timestamp':datetime.datetime.now(),'patch':patch_file,
                             'server':srv,'success':False,'user':st.session_state.username})
                        continue

                    # ── Phase 3: Verify ────────────────────────────────────
                    rrow(srv, pre_lbl,"✅ Done","🔎…",f"Verifying {kb}…",
                         pc=pre_cls, dc="pc-ok", vc="pc-run")
                    ver      = run_ps("PreCheck", st.session_state.username,
                                      st.session_state.password, server=srv)
                    ver_ok   = ver.get('success', False)
                    still_rb = ver.get('RebootRequired','') == 'Yes'

                    if ver_ok:
                        ver_lbl = "⚠️ Restart↑" if still_rb else "✅ Verified"
                        ver_cls = "pc-skip"      if still_rb else "pc-ok"
                        detail  = "Restart needed to activate" if still_rb else f"{kb} active on {os_ver}"
                    else:
                        ver_lbl, ver_cls, detail = "❓ Unknown","pc-skip","Could not verify"

                    rrow(srv, pre_lbl,"✅ Done", ver_lbl, detail,
                         pc=pre_cls, dc="pc-ok", vc=ver_cls)
                    st.session_state.pipeline_results.append(
                        {'server':srv,'kb':kb,'patch':patch_file,
                         'precheck':pre_lbl,'deploy':'✅ Done','verify':ver_lbl,'detail':detail})
                    st.session_state.history.append(
                        {'timestamp':datetime.datetime.now(),'patch':patch_file,
                         'server':srv,'success':dep_ok,'user':st.session_state.username})

            # Final summary
            res = st.session_state.pipeline_results
            n_done   = sum(1 for r in res if '✅ Done'  in r.get('deploy',''))
            n_skip   = sum(1 for r in res if 'Skip'     in r.get('deploy',''))
            n_fail   = sum(1 for r in res if '❌'        in r.get('deploy',''))

            if n_fail == 0 and n_skip == 0:
                overall_ph.markdown(f'<div class="ok-box">🎉 All <strong>{n_done}</strong> deployment(s) succeeded!</div>', unsafe_allow_html=True)
            elif n_done > 0:
                overall_ph.markdown(f'<div class="warn-box">⚠️ {n_done} done · {n_skip} skipped · {n_fail} failed</div>', unsafe_allow_html=True)
            else:
                overall_ph.markdown(f'<div class="err-box">❌ No deployments succeeded — check details above.</div>', unsafe_allow_html=True)

        # ── Pipeline results table ────────────────────────────────────────────
        if st.session_state.pipeline_results:
            st.divider()
            st.markdown('<div class="section-title">📋 Pipeline Results</div>', unsafe_allow_html=True)
            df_pipe = pd.DataFrame(st.session_state.pipeline_results)
            st.dataframe(df_pipe.drop(columns=['patch'], errors='ignore'),
                         use_container_width=True, hide_index=True)
            st.download_button("📥 Export Log", df_pipe.to_csv(index=False),
                               f"pipeline_{datetime.datetime.now():%Y%m%d_%H%M}.csv",
                               "text/csv", key="dl_pipeline")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — HISTORY & LOGS
# ══════════════════════════════════════════════════════════════════════════════
with tab_history:
    st.markdown("**Deployments This Session**")
    if not st.session_state.history:
        st.caption("No deployments yet — go to **Deploy a Patch** to get started.")
    else:
        df = pd.DataFrame([{
            "Time":    h['timestamp'].strftime("%Y-%m-%d %H:%M:%S"),
            "Patch":   h['patch'],
            "Server":  h['server'],
            "Result":  "✅ Success" if h['success'] else "❌ Failed",
            "Done by": h['user']
        } for h in reversed(st.session_state.history)])
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "📥 Download Deployment History",
            df.to_csv(index=False),
            f"deployments_{datetime.datetime.now():%Y%m%d}.csv",
            "text/csv"
        )

    st.divider()
    st.markdown("#### Audit Log")
    st.caption("A permanent record of who signed in and what actions were taken.")
    audit_file = LOG_DIR / "audit.log"
    if audit_file.exists() and audit_file.stat().st_size > 0:
        lines = audit_file.read_text().splitlines()[-100:]
        st.code("\n".join(lines), language=None)
    else:
        st.caption("The audit log is empty — actions will appear here after sign-in and deployments.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — KB DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_kb:
    st.caption("Search the Microsoft Update Catalog and auto-download patches directly to your patch share folder.")

    # ── OS Type Toggle ────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🔍 Search Microsoft Update Catalog</div>', unsafe_allow_html=True)

    # Category toggle — determines search scope and result filtering
    os_type = st.radio(
        "Platform",
        ["🖥️  Windows Server", "💻  Windows Client", "🌐  All"],
        horizontal=True,
        key="kb_os_type",
        label_visibility="collapsed"
    )

    # Sub-version picker (shown only for Server or Client, not All)
    SERVER_VERSIONS = ["Any Server Version", "Windows Server 2022", "Windows Server 2019",
                       "Windows Server 2016", "Windows Server 2012 R2"]
    CLIENT_VERSIONS = ["Any Client Version", "Windows 11", "Windows 10"]

    os_version = None
    if "Server" in os_type:
        os_version = st.selectbox("Server version", SERVER_VERSIONS,
                                  label_visibility="collapsed", key="kb_srv_ver")
    elif "Client" in os_type:
        os_version = st.selectbox("Client version", CLIENT_VERSIONS,
                                  label_visibility="collapsed", key="kb_cli_ver")

    # ── Multi-KB Queue ───────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📋 KB Queue — Download Multiple KBs At Once</div>', unsafe_allow_html=True)
    st.caption("Enter one or more KB numbers (comma or space separated), add to queue, then click Search & Download All.")

    qc1, qc2, qc3 = st.columns([4, 1, 1])
    with qc1:
        queue_input = st.text_input("Add KBs", placeholder="KB5034441, KB5005112, KB5075902",
                                    label_visibility="collapsed", key="kb_queue_input")
    with qc2:
        if st.button("➕ Add", use_container_width=True, key="kb_queue_add"):
            for k in queue_input.replace(',', ' ').split():
                k = k.strip().upper()
                if k and not k.startswith('KB'):
                    k = 'KB' + k
                if k and k not in st.session_state.kb_queue:
                    st.session_state.kb_queue.append(k)
            st.rerun()
    with qc3:
        if st.button("🗑️ Clear", use_container_width=True, key="kb_queue_clear"):
            st.session_state.kb_queue = []
            st.rerun()

    if st.session_state.kb_queue:
        st.markdown(f'<div class="info-box">📋 Queue: <strong>{", ".join(st.session_state.kb_queue)}</strong></div>', unsafe_allow_html=True)
        queue_dest = st.text_input("Download to", value=CFG.PATCH_SHARE,
                                   label_visibility="collapsed", key="kb_queue_dest",
                                   placeholder=r"\\server\share\Patches")
        if st.button(f"⬇️ Search & Download All {len(st.session_state.kb_queue)} KB(s)",
                     type="primary", use_container_width=True, key="kb_queue_run"):
            q_results = []
            q_prog = st.progress(0)
            q_status = st.empty()
            for qi, kb in enumerate(list(st.session_state.kb_queue)):
                q_status.markdown(f"🔍 Searching **{kb}** ({qi+1}/{len(st.session_state.kb_queue)})…")
                raw = catalog_search(kb)
                hits = [r for r in raw if not r.get('error') and r.get('uid') and kb in r.get('kb','').upper()]
                if not hits:
                    hits = [r for r in raw if not r.get('error') and r.get('uid')]
                if hits:
                    r = hits[0]
                    prod = r.get('products', '')
                    ext = '.cab' if 'cab' in r.get('title','').lower() else '.msu'
                    os_tag = ('_2022' if '2022' in prod else '_2019' if '2019' in prod else
                              '_2016' if '2016' in prod else '_2012R2' if '2012' in prod else
                              '_Win11' if 'windows 11' in prod.lower() else
                              '_Win10' if 'windows 10' in prod.lower() else '')
                    fname = f"{kb}{os_tag}{ext}"
                    q_status.markdown(f"⬇️ Downloading **{fname}** ({qi+1}/{len(st.session_state.kb_queue)})…")
                    res = download_kb_to_share(uid=r['uid'], filename=fname, dest_folder=queue_dest.strip())
                    ok = res.get('success', False)
                    q_results.append({'KB': kb, 'File': fname, 'Status': '✅ Done' if ok else f"❌ {res.get('error','Failed')}"})
                    st.session_state.kb_downloads.append({
                        'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'kb': kb, 'title': r.get('title','')[:60], 'filename': fname,
                        'size_mb': res.get('size_mb', 0), 'dest': queue_dest, 'status': '✅ Success' if ok else '❌ Failed'
                    })
                else:
                    q_results.append({'KB': kb, 'File': '—', 'Status': '❌ Not found in catalog'})
                q_prog.progress((qi + 1) / len(st.session_state.kb_queue))
            q_status.empty()
            ok_n = sum(1 for x in q_results if '✅' in x['Status'])
            if ok_n == len(q_results):
                st.markdown(f'<div class="ok-box">✅ All <strong>{ok_n}</strong> KBs downloaded!</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="warn-box">⚠️ {ok_n}/{len(q_results)} succeeded.</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(q_results), use_container_width=True, hide_index=True)
            st.session_state.kb_queue = []
            st.session_state.patches = []

    st.divider()

    # Search bar + button
    col_q, col_srch = st.columns([5, 1])
    with col_q:
        kb_query = st.text_input(
            "KB or keyword",
            placeholder="e.g. KB5034441  or  Cumulative Update  or  Servicing Stack",
            label_visibility="collapsed",
            key="kb_query_input"
        )
    with col_srch:
        do_search = st.button("🔍 Search", type="primary", use_container_width=True, key="kb_search_btn")

    if do_search:
        if not kb_query.strip():
            st.warning("Enter a KB number or keyword to search.")
        else:
            # Build query — append version hint when specific version chosen
            full_query = kb_query.strip()
            if os_version and "Any" not in str(os_version):
                full_query += f" {os_version}"
            elif "Server" in os_type and (not os_version or "Any" in str(os_version)):
                full_query += " Windows Server"
            elif "Client" in os_type and (not os_version or "Any" in str(os_version)):
                full_query += " Windows"

            with st.spinner(f"Searching Microsoft Update Catalog…"):
                raw = catalog_search(full_query)

            # Client-side filter — remove server results when Client selected and vice-versa
            if "Server" in os_type:
                filtered = [r for r in raw if "error" in r or
                            "server" in r.get("products","").lower() or
                            "server" in r.get("title","").lower()]
            elif "Client" in os_type:
                filtered = [r for r in raw if "error" in r or (
                            "server" not in r.get("products","").lower() and
                            ("windows 10" in r.get("products","").lower() or
                             "windows 11" in r.get("products","").lower() or
                             "windows 10" in r.get("title","").lower() or
                             "windows 11" in r.get("title","").lower()))]
            else:
                filtered = raw

            st.session_state.kb_results = filtered
            st.session_state.kb_raw_results = raw   # keep full set for "All" toggle switch

    # ── Results ───────────────────────────────────────────────────────────────
    if st.session_state.kb_results:
        results = st.session_state.kb_results

        if len(results) == 1 and "error" in results[0]:
            st.markdown(f'<div class="err-box">❌ Search failed: {results[0]["error"]}</div>', unsafe_allow_html=True)
        elif not results:
            st.markdown('<div class="warn-box">⚠️ No results found for that filter. Try switching to <strong>All</strong> or a broader keyword.</div>', unsafe_allow_html=True)
        else:
            # Show which filter is active
            filter_label = os_version if (os_version and "Any" not in str(os_version)) else os_type.split("  ")[-1]
            st.markdown(
                f'<div class="ok-box">✅ <strong>{len(results)}</strong> result(s) — '
                f'filtered to: <strong>{filter_label}</strong></div>',
                unsafe_allow_html=True
            )

            # Get filenames already in patch share for duplicate detection
            existing_kbs = set()
            try:
                for root, _, files in os.walk(CFG.PATCH_SHARE):
                    for f in files:
                        m = re.search(r'KB(\d+)', f, re.IGNORECASE)
                        if m:
                            existing_kbs.add(f"KB{m.group(1).upper()}")
            except Exception:
                pass

            df_res = pd.DataFrame([{
                "#":        i + 1,
                "In Share": "✅" if r.get("kb","") in existing_kbs else "",
                "KB":       r.get("kb", "—"),
                "Title":    r.get("title", ""),
                "Products": r.get("products", ""),
                "Type":     r.get("classif", ""),
                "Date":     r.get("date", ""),
                "Size":     r.get("size", ""),
                "_uid":     r.get("uid", ""),
            } for i, r in enumerate(results)])

            st.dataframe(
                df_res.drop(columns=["_uid"]),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "#":         st.column_config.NumberColumn("#",         width=40),
                    "In Share":  st.column_config.TextColumn("✅ In Share", width=80),
                    "KB":        st.column_config.TextColumn("KB",          width="small"),
                    "Title":     st.column_config.TextColumn("Title",       width="large"),
                    "Products":  st.column_config.TextColumn("Products",    width="medium"),
                    "Type":      st.column_config.TextColumn("Type",        width="small"),
                    "Date":      st.column_config.TextColumn("Date",        width="small"),
                    "Size":      st.column_config.TextColumn("Size",        width="small"),
                }
            )

            valid_results = [r for r in results if r.get('uid')]
            st.divider()

            # ── BULK DOWNLOAD SECTION ─────────────────────────────────────────
            st.markdown('<div class="section-title">⬇️ Bulk Download to Patch Share</div>', unsafe_allow_html=True)

            col_dest_bulk, col_dl_bulk = st.columns([4, 1])
            with col_dest_bulk:
                dest_folder = st.text_input(
                    "Destination folder",
                    value=CFG.PATCH_SHARE,
                    label_visibility="collapsed",
                    key="kb_dest_folder",
                    placeholder=r"\\server\share\Patches"
                )
            with col_dl_bulk:
                st.markdown("<div style='margin-top:1px'></div>", unsafe_allow_html=True)

            # Checkbox grid for selecting which results to download
            st.caption("Tick the patches you want to download, then click **Download Selected**.")

            if 'kb_bulk_selected' not in st.session_state:
                st.session_state.kb_bulk_selected = {}

            # Select All / None controls
            ca, cn, _ = st.columns([1, 1, 5])
            with ca:
                if st.button("☑ All", key="bulk_sel_all", use_container_width=True):
                    for r in valid_results:
                        st.session_state.kb_bulk_selected[r["uid"]] = True
                    st.rerun()
            with cn:
                if st.button("✕ None", key="bulk_sel_none", use_container_width=True):
                    st.session_state.kb_bulk_selected = {}
                    st.rerun()

            # Per-row checkboxes
            for i, r in enumerate(valid_results):
                uid       = r.get("uid","")
                kb_num    = r.get("kb","KB_?")
                title_s   = r.get("title","")[:65]
                size_s    = r.get("size","")
                in_share  = "✅" if r.get("kb","") in existing_kbs else ""
                checked   = st.session_state.kb_bulk_selected.get(uid, False)
                new_check = st.checkbox(
                    f"{in_share} **{kb_num}** — {title_s}  `{size_s}`",
                    value=checked,
                    key=f"bulk_chk_{i}"
                )
                st.session_state.kb_bulk_selected[uid] = new_check

            # Collect selected
            selected_uids = [r for r in valid_results
                             if st.session_state.kb_bulk_selected.get(r.get("uid",""), False)]
            already_have  = [r for r in selected_uids if r.get("kb","") in existing_kbs]
            to_download   = [r for r in selected_uids if r.get("kb","") not in existing_kbs]

            sel_label = f"{len(selected_uids)} selected"
            if already_have:
                sel_label += f"  ·  {len(already_have)} already in share (will skip)"

            st.caption(sel_label if selected_uids else "Nothing selected.")

            start_bulk = st.button(
                f"⬇️ Download {len(to_download)} patch(es)" if to_download else "⬇️ Download Selected",
                type="primary",
                use_container_width=True,
                disabled=not to_download,
                key="kb_bulk_dl"
            )

            if start_bulk and to_download:
                overall_bar  = st.progress(0, text="Starting bulk download…")
                item_bar     = st.progress(0)
                item_status  = st.empty()
                bulk_results = []

                for idx, r in enumerate(to_download):
                    uid    = r.get("uid","")
                    kb_num = r.get("kb","KB_unknown")
                    prod   = r.get("products","")
                    title_l = r.get("title","").lower()
                    ext    = ".cab" if "cab" in title_l else ".msu"
                    os_tag = ("_2022" if "2022" in prod else
                              "_2019" if "2019" in prod else
                              "_2016" if "2016" in prod else
                              "_2012R2" if "2012" in prod else
                              "_Win11" if "windows 11" in prod.lower() else
                              "_Win10" if "windows 10" in prod.lower() else "")
                    fname  = f"{kb_num}{os_tag}{ext}"

                    item_status.markdown(f"⏳ Downloading **{fname}**  ({idx+1}/{len(to_download)})")
                    item_bar.progress(0, text="Resolving URL…")

                    def _item_prog(pct, _bar=item_bar, _fname=fname):
                        _bar.progress(min(pct, 1.0), text=f"{_fname}: {pct*100:.0f}%")

                    result = download_kb_to_share(
                        uid=uid, filename=fname,
                        dest_folder=dest_folder.strip(),
                        progress_callback=_item_prog
                    )
                    ok = result.get("success", False)
                    bulk_results.append({
                        "KB":       kb_num,
                        "Filename": fname,
                        "Size MB":  result.get("size_mb", 0) if ok else 0,
                        "Status":   "✅ Done" if ok else f"❌ {result.get('error','Failed')}",
                    })
                    if ok:
                        st.session_state.kb_downloads.append({
                            "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "kb": kb_num, "title": r.get("title","")[:60],
                            "filename": fname, "size_mb": result.get("size_mb", 0),
                            "dest": dest_folder, "status": "✅ Success",
                        })
                    else:
                        st.session_state.kb_downloads.append({
                            "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "kb": kb_num, "title": r.get("title","")[:60],
                            "filename": fname, "size_mb": 0,
                            "dest": dest_folder, "status": "❌ Failed",
                        })
                    overall_bar.progress((idx + 1) / len(to_download),
                                         text=f"Overall: {idx+1}/{len(to_download)}")

                item_bar.empty(); item_status.empty(); overall_bar.empty()
                st.session_state.patches = []   # refresh deploy tab
                st.session_state.kb_bulk_selected = {}

                ok_cnt = sum(1 for r in bulk_results if "Done" in r["Status"])
                if ok_cnt == len(bulk_results):
                    st.markdown(f'<div class="ok-box">✅ All <strong>{ok_cnt}</strong> patches downloaded successfully to <code>{dest_folder}</code></div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="warn-box">⚠️ {ok_cnt}/{len(bulk_results)} succeeded — check Status below.</div>', unsafe_allow_html=True)

                st.dataframe(pd.DataFrame(bulk_results), use_container_width=True, hide_index=True)

    st.divider()

    # ── Download history ──────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📋 Download History (This Session)</div>', unsafe_allow_html=True)
    if not st.session_state.kb_downloads:
        st.caption("No downloads yet this session.")
    else:
        df_dl = pd.DataFrame(st.session_state.kb_downloads)
        st.dataframe(df_dl, use_container_width=True, hide_index=True,
                     column_config={
                         "time":     st.column_config.TextColumn("Time",     width="small"),
                         "kb":       st.column_config.TextColumn("KB",       width="small"),
                         "title":    st.column_config.TextColumn("Title",    width="large"),
                         "filename": st.column_config.TextColumn("Filename", width="medium"),
                         "size_mb":  st.column_config.NumberColumn("MB",     width="small", format="%.1f"),
                         "dest":     st.column_config.TextColumn("Saved To", width="medium"),
                         "status":   st.column_config.TextColumn("Status",   width="small"),
                     })
        st.download_button(
            "📥 Export Download Log",
            df_dl.to_csv(index=False),
            f"kb_downloads_{datetime.datetime.now():%Y%m%d}.csv",
            "text/csv", key="dl_kb_log"
        )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — ROLLBACK
# ══════════════════════════════════════════════════════════════════════════════
with tab_rollback:
    st.caption("Uninstall a previously applied Windows KB patch from one or more servers.")

    st.markdown('<div class="warn-box">⚠️ <strong>Caution:</strong> Rollback removes a patch permanently. The server may need a restart to complete the uninstall. Only use this if a patch is causing issues.</div>', unsafe_allow_html=True)

    # ── Step 1: Choose servers ────────────────────────────────────────────────
    st.markdown('<span class="step-num">1</span> **Choose target servers**', unsafe_allow_html=True)
    rb_servers = st.multiselect(
        "Servers to rollback on",
        CFG.SERVERS,
        default=[],
        label_visibility="collapsed",
        key="rb_servers",
        placeholder="Select one or more servers…"
    )

    st.divider()

    # ── Step 2: KB number ─────────────────────────────────────────────────────
    st.markdown('<span class="step-num">2</span> **Enter the KB number to remove**', unsafe_allow_html=True)
    st.caption("You can find installed KBs by using the **Installed Patches** lookup on the Server Status tab.")

    col_kb, col_hint = st.columns([2, 3])
    with col_kb:
        rb_kb = st.text_input(
            "KB Number",
            placeholder="e.g. KB5034441",
            label_visibility="collapsed",
            key="rb_kb_input"
        )
    with col_hint:
        # Quick lookup — if scan results exist, show KBs from installed cache
        if st.session_state.installed_cache:
            cached_servers = list(st.session_state.installed_cache.keys())
            hint_srv = st.selectbox("Quick-pick from cached patches on:", cached_servers,
                                     label_visibility="collapsed", key="rb_hint_srv")
            cached = st.session_state.installed_cache.get(hint_srv, [])
            if isinstance(cached, list) and cached:
                kb_options = sorted(set(r.get("HotFixID","") or r.get("KB Number","")
                                        for r in cached if r.get("HotFixID") or r.get("KB Number")))
                if kb_options:
                    picked = st.selectbox("Pick a KB", kb_options,
                                           label_visibility="collapsed", key="rb_kb_pick")
                    if picked:
                        # Mirror into the text input via session state
                        st.session_state["rb_kb_prefill"] = picked
                        st.caption(f"Selected: **{picked}** — paste it into the KB field above.")

    st.divider()

    # ── Step 3: Reboot after rollback ─────────────────────────────────────────
    st.markdown('<span class="step-num">3</span> **Schedule restart after rollback** *(recommended)*', unsafe_allow_html=True)
    rb_reboot = st.toggle("Schedule restart after uninstall", key="rb_reboot_toggle")
    rb_reboot_time = None
    if rb_reboot:
        c1, c2 = st.columns(2)
        with c1:
            rb_d = st.date_input("Restart date", datetime.date.today() + datetime.timedelta(days=1),
                                  min_value=datetime.date.today(), key="rb_date")
        with c2:
            rb_t = st.time_input("Restart time", datetime.time(2, 0), key="rb_time")
        rb_reboot_time = f"{rb_d} {rb_t:%H:%M}"
        st.markdown(f'<div class="warn-box">⏰ Restart scheduled for <strong>{rb_reboot_time}</strong></div>', unsafe_allow_html=True)

    st.divider()

    # ── Step 4: Execute rollback ──────────────────────────────────────────────
    st.markdown('<span class="step-num">4</span> **Review and run rollback**', unsafe_allow_html=True)

    kb_clean = rb_kb.strip().upper()
    if not kb_clean.startswith("KB") and kb_clean.isdigit():
        kb_clean = "KB" + kb_clean

    if rb_servers and kb_clean.startswith("KB"):
        reboot_note = f"restart at {rb_reboot_time}" if rb_reboot_time else "no restart scheduled"
        st.markdown(
            f'<div class="warn-box">About to remove <strong>{kb_clean}</strong> from '
            f'<strong>{len(rb_servers)} server(s)</strong> — {reboot_note}.</div>',
            unsafe_allow_html=True
        )

        if st.button(f"↩️ Run Rollback on {len(rb_servers)} server(s)",
                     type="primary", use_container_width=True, key="rb_execute"):
            rb_results, rb_progress, rb_status = [], st.progress(0), st.empty()

            for i, srv in enumerate(rb_servers):
                rb_status.markdown(f"Rolling back **{kb_clean}** on **{srv}**… ({i+1}/{len(rb_servers)})")
                res = run_ps("Rollback", st.session_state.username, st.session_state.password,
                             server=srv, kb_number=kb_clean)
                ok = res.get('success', False)
                msg = res.get('Message', res.get('error', 'No details'))

                # Schedule reboot if requested and rollback succeeded
                if ok and rb_reboot_time:
                    run_ps("Deploy", st.session_state.username, st.session_state.password,
                           server=srv, reboot_time=rb_reboot_time)

                rb_results.append({
                    "Server":  srv,
                    "KB":      kb_clean,
                    "Result":  "✅ Removed" if ok else "❌ Failed",
                    "Details": msg,
                    "Time":    datetime.datetime.now().strftime("%H:%M:%S"),
                })
                audit.info(f"Rollback | KB={kb_clean} | Server={srv} | User={st.session_state.username} | OK={ok}")
                rb_progress.progress((i+1) / len(rb_servers))

            rb_status.empty(); rb_progress.empty()

            ok_count   = sum(1 for r in rb_results if "Removed" in r["Result"])
            fail_count = len(rb_results) - ok_count

            if ok_count == len(rb_results):
                st.markdown(f'<div class="ok-box">✅ <strong>{kb_clean}</strong> successfully removed from all {ok_count} server(s).</div>', unsafe_allow_html=True)
            elif ok_count > 0:
                st.markdown(f'<div class="warn-box">⚠️ Removed from {ok_count}, failed on {fail_count} — check Details below.</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="err-box">❌ Rollback failed on all servers — {kb_clean} may not be installed, or check server connectivity.</div>', unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Total",    len(rb_results))
            with c2: st.metric("✅ Removed", ok_count)
            with c3: st.metric("❌ Failed",  fail_count)
            st.dataframe(pd.DataFrame(rb_results), use_container_width=True, hide_index=True)

            # Invalidate installed cache for affected servers
            for srv in rb_servers:
                st.session_state.installed_cache.pop(srv, None)
    else:
        if not rb_servers:
            st.caption("Select at least one server above to continue.")
        elif not kb_clean.startswith("KB"):
            st.caption("Enter a valid KB number above (e.g. KB5034441).")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
with tab_settings:
    st.caption("Changes are saved immediately to `opcenter_config.json` next to this app and take effect on the next page action.")

    live_cfg = load_config()

    # ── Section 1: Server Hostnames ───────────────────────────────────────────
    st.markdown('<div class="section-title">🖥️ Server Hostnames</div>', unsafe_allow_html=True)

    servers = list(live_cfg["servers"])  # copy

    # Add new server
    col_add, col_btn = st.columns([4, 1])
    with col_add:
        new_host = st.text_input("Add a server hostname", placeholder="e.g. L11SGRIFP007",
                                 label_visibility="collapsed", key="new_hostname")
    with col_btn:
        if st.button("➕ Add", use_container_width=True, key="btn_add_host"):
            host = new_host.strip().upper()
            if not host:
                st.warning("Enter a hostname first.")
            elif host in [s.upper() for s in servers]:
                st.warning(f"**{host}** is already in the list.")
            else:
                servers.append(host)
                live_cfg["servers"] = servers
                save_config(live_cfg)
                st.success(f"✅ Added **{host}**")
                st.rerun()

    # Current server list with remove buttons
    if servers:
        st.markdown('<div style="margin-top:6px;"></div>', unsafe_allow_html=True)
        for i, srv in enumerate(servers):
            c1, c2 = st.columns([5, 1])
            with c1:
                st.markdown(f"""
                <div style="background:#111827;border:1px solid #1e2a45;border-radius:6px;
                            padding:8px 14px;font-size:13px;color:#e2e8f0;font-family:monospace;
                            margin-bottom:4px;">
                    🖥️ &nbsp;{srv}
                </div>""", unsafe_allow_html=True)
            with c2:
                if st.button("🗑️ Remove", key=f"rm_{i}", use_container_width=True):
                    servers.pop(i)
                    live_cfg["servers"] = servers
                    save_config(live_cfg)
                    st.rerun()
    else:
        st.markdown('<div class="warn-box">⚠️ No servers configured — add at least one hostname above.</div>', unsafe_allow_html=True)

    st.divider()

    # ── Section 2: Patch Share Folder ─────────────────────────────────────────
    st.markdown('<div class="section-title">📁 Patch Share Folder</div>', unsafe_allow_html=True)
    st.caption("The UNC path to the shared folder where patch files (.msu / .exe / .cab) are stored.")

    col_share, col_save = st.columns([5, 1])
    with col_share:
        new_share = st.text_input(
            "Patch share path",
            value=live_cfg["patch_share"],
            placeholder=r"\\server\share\Patches",
            label_visibility="collapsed",
            key="share_input"
        )
    with col_save:
        if st.button("💾 Save", use_container_width=True, key="btn_save_share"):
            val = new_share.strip()
            if not val:
                st.warning("Path cannot be empty.")
            else:
                live_cfg["patch_share"] = val
                save_config(live_cfg)
                st.success("✅ Share path saved")
                st.rerun()

    st.divider()

    # ── Section 3: Domain ─────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🔑 Domain</div>', unsafe_allow_html=True)
    st.caption("The Windows domain prefix used when authenticating (e.g. ZL → ZL\\username).")

    col_dom, col_domsave = st.columns([5, 1])
    with col_dom:
        new_domain = st.text_input("Domain", value=live_cfg["domain"],
                                   placeholder="ZL", label_visibility="collapsed", key="domain_input")
    with col_domsave:
        if st.button("💾 Save", use_container_width=True, key="btn_save_domain"):
            val = new_domain.strip()
            if not val:
                st.warning("Domain cannot be empty.")
            else:
                live_cfg["domain"] = val.upper()
                save_config(live_cfg)
                st.success(f"✅ Domain set to **{val.upper()}**")
                st.rerun()

    st.divider()

    # ── Section 4: Recurring Scan Schedule ────────────────────────────────────
    st.markdown('<div class="section-title">🔄 Recurring Scan Schedule</div>', unsafe_allow_html=True)
    st.caption("OpsCenter will automatically re-scan all servers at this interval while the app is open. Set to 0 to disable.")

    interval_opts = {
        "Disabled (manual only)": 0,
        "Every 15 minutes":       15,
        "Every 30 minutes":       30,
        "Every 1 hour":           60,
        "Every 2 hours":          120,
        "Every 4 hours":          240,
        "Every 8 hours":          480,
    }
    current_interval = live_cfg.get("scan_interval_mins", 0)
    # Find label matching current value, default to Disabled
    current_label = next((k for k, v in interval_opts.items() if v == current_interval),
                         "Disabled (manual only)")

    col_int, col_intsave = st.columns([4, 1])
    with col_int:
        chosen_interval_label = st.selectbox(
            "Scan interval",
            list(interval_opts.keys()),
            index=list(interval_opts.keys()).index(current_label),
            label_visibility="collapsed",
            key="scan_interval_select"
        )
    with col_intsave:
        if st.button("💾 Save", use_container_width=True, key="btn_save_interval"):
            new_interval = interval_opts[chosen_interval_label]
            live_cfg["scan_interval_mins"] = new_interval
            save_config(live_cfg)
            # Reset the schedule timer immediately
            if new_interval > 0:
                st.session_state.next_auto_scan = (
                    datetime.datetime.now() + datetime.timedelta(minutes=new_interval)
                )
                st.success(f"✅ Auto-scan set — next scan in {new_interval} min")
            else:
                st.session_state.next_auto_scan = None
                st.success("✅ Auto-scan disabled")
            st.rerun()

    # Show current status
    if current_interval > 0 and st.session_state.next_auto_scan:
        delta = st.session_state.next_auto_scan - datetime.datetime.now()
        mins_left = max(0, int(delta.total_seconds() // 60))
        st.caption(f"⏱️ Next scheduled scan in **{mins_left} minute(s)**")
    elif current_interval == 0:
        st.caption("Auto-scan is currently **disabled** — scans only run when you click Check All Servers or on login.")

    st.divider()

    # ── Section 5: Current config preview ─────────────────────────────────────
    with st.expander("📄 View raw config file (opcenter_config.json)"):
        st.code(json.dumps(live_cfg, indent=2), language="json")
        st.caption(f"Saved at: `{CONFIG_FILE}`")

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.divider()
c1, c2, c3 = st.columns(3)
with c1: st.caption(f"OpsCenter v2.3  ·  {datetime.datetime.now():%Y-%m-%d %H:%M}")
with c2: st.caption(f"Signed in as: {CFG.DOMAIN}\\{st.session_state.username}")
with c3: st.caption(f"Patch share: {CFG.PATCH_SHARE}")
