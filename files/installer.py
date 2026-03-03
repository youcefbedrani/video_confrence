"""
Nextcloud Enterprise Installer v3.0
=====================================
- Uses Docker Engine ONLY (no Docker Desktop, no login, no account)
- Docker Engine installs silently as a Windows service
- Client never sees Docker at all
- Generates SSL certificate with Python (no OpenSSL needed)
- Nextcloud runs with HTTPS on LAN
- Auto-starts on Windows boot
- Works on Windows 10 / 11

Build:
  pip install pyinstaller
  python -m PyInstaller --onefile --windowed --uac-admin --name NextcloudInstaller installer.py
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import subprocess
import sys
import os
import socket
import time
import platform
import webbrowser
import ctypes
import tempfile
import shutil
import urllib.request
import ssl as ssl_module
import ipaddress
import datetime
import zipfile
import json
from pathlib import Path
import hashlib
import base64

IS_WINDOWS = platform.system() == "Windows"

# Move cryptography imports to top for better PyInstaller detection
try:
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.backends import default_backend
    HAS_CRYPTO = True
except ImportError:
    HAS_CRYPTO = False

# Remove: import winshell (causes bundling errors)
# Remove: from win32com.client import Dispatch (unused)

# ─────────────────────────────────────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

APP_VERSION = "3.0"
INSTALL_DIR = Path("C:/Nextcloud-LAN") if IS_WINDOWS else Path.home() / "Nextcloud-LAN"

# Docker Engine for Windows (no Desktop, no login required)
# This is the official Docker Engine MSI — silent install
DOCKER_ENGINE_URL = (
    "https://download.docker.com/win/static/stable/x86_64/"
    "docker-27.3.1.zip"
)

# Docker Compose plugin binary
DOCKER_COMPOSE_URL = (
    "https://github.com/docker/compose/releases/download/"
    "v2.29.7/docker-compose-windows-x86_64.exe"
)

COLORS = {
    "bg":      "#0d1117",
    "panel":   "#161b22",
    "card":    "#21262d",
    "border":  "#30363d",
    "accent":  "#2f81f7",
    "accent2": "#388bfd",
    "success": "#3fb950",
    "warning": "#d29922",
    "error":   "#f85149",
    "text":    "#e6edf3",
    "muted":   "#8b949e",
    "dim":     "#484f58",
}

STEPS = [
    "System Check",
    "Docker Engine",
    "Docker Start",
    "SSL Certificate",
    "Create Config",
    "Pull Images",
    "Start Services",
    "Auto Setup",
    "Complete",
]

# ─────────────────────────────────────────────────────────────────────────────
#  UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base_path = os.path.abspath(".")

    # Prioritize bundled path
    target = Path(base_path) / relative_path
    if target.exists():
        return target
        
    # fallback to alongside the exe/script
    side_path = Path(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)) / relative_path
    if side_path.exists():
        return side_path
        
    # Last resort fallback to current directory
    alt_target = Path(os.getcwd()) / relative_path
    return alt_target

def to_wsl_path(win_path):
    """Convert a Windows path like C:\\Foo to /mnt/c/Foo."""
    p = str(Path(win_path)).replace('\\', '/')
    if ':' in p:
        drive, path = p.split(':', 1)
        return f"/mnt/{drive.lower()}{path}"
    return p

def get_desktop_path():
    """Get the desktop path for the current OS."""
    if IS_WINDOWS:
        try:
            from ctypes import wintypes
            CSIDL_DESKTOP = 0
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOP, None, 0, buf)
            return buf.value
        except Exception:
            # Fallback for Windows
            return os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
    else:
        # Linux / macOS
        return os.path.join(os.path.expanduser("~"), "Desktop")

def is_admin():
    """Check if the current user has administrative privileges."""
    if IS_WINDOWS:
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False
    else:
        return os.getuid() == 0

def win_version():
    if IS_WINDOWS:
        v = sys.getwindowsversion()
        return v.major, v.minor
    return (0, 0)

def run_cmd(cmd, timeout=300):
    try:
        r = subprocess.run(
            cmd, shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        stdout = r.stdout
        stderr = r.stderr
        
        # WSL commands on Windows often output in UTF-16
        def decode_safe(data):
            if not data: return ""
            # Try UTF-8 first
            try:
                return data.decode("utf-8").replace('\x00', '')
            except UnicodeDecodeError:
                # Try UTF-16
                try:
                    return data.decode("utf-16").replace('\x00', '')
                except:
                    return data.decode("utf-8", errors="ignore").replace('\x00', '')

        return (r.returncode, decode_safe(stdout), decode_safe(stderr))
    except subprocess.TimeoutExpired:
        return -1, "", "Timeout"
    except Exception as e:
        return -1, "", str(e)

# ─────────────────────────────────────────────────────────────────────────────
#  UNIVERSAL EXECUTION LAYER
# ─────────────────────────────────────────────────────────────────────────────

def run_universal_cmd(cmd, timeout=300):
    """
    Runs a command NATIVELY on Linux, 
    but AUTOMATICALLY wraps it in WSL if on Windows.
    """
    if IS_WINDOWS:
        distro = get_wsl_distro()
        # Escape quotes for the bash -c wrapper
        safe_cmd = cmd.replace('"', '\\"')
        full_cmd = f'wsl -d {distro} -u root bash -c "export DOCKER_DEFAULT_PLATFORM=linux/amd64 && {safe_cmd}"'
    else:
        # On Linux, run directly (assume sudo/root for docker tasks)
        full_cmd = f'bash -c "{cmd}"'

    return run_cmd(full_cmd, timeout=timeout)

def get_wsl_distro():
    """Detects the installed Ubuntu distro name on Windows."""
    if not IS_WINDOWS: return "linux"
    code, out, _ = run_cmd("wsl -l -v")
    if code != 0 or not out: return "Ubuntu"
    
    # Process output line by line, looking for Ubuntu
    # Handle cases where multiple Ubuntu distros exist
    for line in out.splitlines():
        if not line.strip(): continue
        if "Ubuntu" in line:
            parts = line.split()
            # If the first part is '*', the name is the second part
            if parts[0] == '*':
                return parts[1].strip()
            return parts[0].strip()
    return "Ubuntu"

def is_docker_engine_installed():
    """Checks if Docker is installed inside the Linux environment."""
    code, out, _ = run_universal_cmd("docker --version")
    return code == 0

def is_docker_running():
    """Checks if Docker daemon is active inside the Linux environment."""
    code, out, _ = run_universal_cmd("docker ps")
    return code == 0

def download_file(url, dest, progress_cb=None):
    ctx = ssl_module.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode    = ssl_module.CERT_NONE
    req  = urllib.request.urlopen(url, context=ctx, timeout=60)
    total = int(req.headers.get("Content-Length", 0))
    downloaded = 0
    t0 = time.time()
    with open(dest, "wb") as f:
        while True:
            chunk = req.read(65536)
            if not chunk:
                break
            f.write(chunk)
            downloaded += len(chunk)
            elapsed = max(time.time() - t0, 0.001)
            speed   = downloaded / elapsed / 1024
            pct     = int(downloaded * 100 / total) if total else 0
            if progress_cb:
                progress_cb(pct, speed)

def is_docker_engine_installed():
    """Check if Docker is installed inside WSL."""
    distro = get_wsl_distro()
    code, out, _ = run_cmd(f"wsl -d {distro} -u root docker --version", timeout=10)
    return code == 0

def is_docker_running():
    """Check if Docker daemon is running inside WSL."""
    distro = get_wsl_distro()
    code, out, _ = run_cmd(f"wsl -d {distro} -u root docker info", timeout=15)
    return code == 0

def gen_ssl_cert(ip, cert_path, key_path):
    """Generate self-signed SSL cert using bundled cryptography."""
    if not HAS_CRYPTO:
        return False, "cryptography library not bundled"
    
    try:
        return _do_gen_cert(ip, cert_path, key_path)
    except Exception as e:
        return False, str(e)

def _do_gen_cert(ip, cert_path, key_path):
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.backends import default_backend

    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
        backend=default_backend()
    )
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, ip),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Nextcloud LAN"),
        x509.NameAttribute(NameOID.COUNTRY_NAME, "DZ"),
    ])
    san = [x509.DNSName("localhost"), x509.DNSName(ip)]
    try:
        san.append(x509.IPAddress(ipaddress.IPv4Address(ip)))
        san.append(x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")))
    except Exception:
        pass

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(datetime.datetime.utcnow())
        .not_valid_after(
            datetime.datetime.utcnow() + datetime.timedelta(days=3650))
        .add_extension(
            x509.SubjectAlternativeName(san), critical=False)
        .sign(key, hashes.SHA256(), default_backend())
    )

    Path(key_path).write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.TraditionalOpenSSL,
            serialization.NoEncryption(),
        )
    )
    Path(cert_path).write_bytes(
        cert.public_bytes(serialization.Encoding.PEM)
    )
    return True, "ok"

def trust_ssl_cert(cert_path):
    """Adds the certificate to Windows Trusted Root Store."""
    if not cert_path.exists():
        return False, "Cert file missing"
    
    # Use certutil to add to Root store
    # -addstore: Add certificate to store
    # -f: Force overwrite
    # "Root": Trusted Root Certification Authorities
    cmd = f'certutil -addstore -f "Root" "{cert_path}"'
    code, out, err = run_cmd(cmd)
    if code == 0:
        return True, "Success"
    return False, err

def create_desktop_shortcut(url, name="Nextcloud", icon_path=None):
    """Creates a professional .lnk shortcut to the Nextcloud URL via PowerShell."""
    try:
        desktop = get_desktop_path()
        lnk_path = os.path.join(desktop, f"{name}.lnk")
        
        # Use PowerShell to create a proper Shell Link (.lnk)
        # This is more robust than a .url file and allows custom icons
        ps_cmd = (
            f'$WshShell = New-Object -ComObject WScript.Shell; '
            f'$Shortcut = $WshShell.CreateShortcut("{lnk_path}"); '
            f'$Shortcut.TargetPath = "{url}"; '
        )
        if icon_path and os.path.exists(icon_path):
            ps_cmd += f'$Shortcut.IconLocation = "{icon_path}"; '
        
        ps_cmd += '$Shortcut.Save()'
        
        subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True, check=True)
        return True
    except Exception as e:
        # Fallback to simple .url if powershell fails
        try:
            path = os.path.join(get_desktop_path(), f"{name}.url")
            with open(path, "w") as f:
                f.write("[InternetShortcut]\n")
                f.write(f"URL={url}\n")
            return True
        except:
            return False

def check_and_enable_wsl(log_cb=None):
    """Checks if WSL is enabled; attempts to enable it on Windows 10/11."""
    if not IS_WINDOWS:
        return True, "Linux native (no WSL needed)"

    def _log(m, tag=""):
        if log_cb: log_cb(m, tag)

    _log("    Checking 'wsl --status'...", "dim")
    # 1. Is it 100% ready? (Short timeout to prevent hang)
    code, out, _ = run_cmd("wsl --status", timeout=15)
    if code == 0:
        return True, "WSL ready"
    
    # 2. Force WSL kernel update (Essential for many Win10 machines)
    _log("    Checking for WSL kernel update...", "dim")
    run_cmd("wsl --update", timeout=60)
    
    # 3. Set WSL2 as default
    _log("    Setting WSL2 as default...", "dim")
    run_cmd("wsl --set-default-version 2", timeout=15)
    
    # 4. Is the command available? (Lenient fallback)
    _log("    Locating 'wsl' command...", "dim")
    code, out, _ = run_cmd("where wsl", timeout=10)
    if code == 0:
        return True, "WSL available"

    # 5. If missing entirely, enable it
    # Note: older Win10 builds ( < 19041 ) don't support --install
    v = sys.getwindowsversion()
    _log(f"    Windows Build: {v.build}", "dim")
    
    if v.build < 18362:
        return False, "Windows version too old for WSL (1903+ required)"

    _log("    Enabling WSL features (wsl --install)...", "warn")
    run_cmd("wsl --install --no-launch", timeout=300)
    
    # Return False to indicate a restart is likely needed
    return False, "WSL features enabled - Restart Required"

def is_windows_server():
    """Detects if we are running on Windows Server."""
    return "server" in platform.release().lower()

def enable_server_containers():
    """Enables Container support for Windows Server."""
    if is_windows_server():
        run_cmd("Install-WindowsFeature -Name Containers", timeout=120)
        return True
    return False

# ─────────────────────────────────────────────────────────────────────────────
#  LICENSE MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────

# Test Keys (Example):
# Trial (Expired): NC-TR-TESTX-6A3CF48575
# Trial:    NC-TR-12345-E93D190580
# Annual:   NC-AN-12345-420E9043C0
# Lifetime: NC-LT-12345-6677C57E6B

LICENSE_SALT = "Nextcloud_DZ_Secret_2026"
# Configuration path (Universal)
LICENSE_PATH = Path.home() / ".NextcloudInstaller" / "license.dat"

class LicenseManager:
    @staticmethod
    def generate_checksum(key_data):
        combined = key_data + LICENSE_SALT
        return hashlib.sha256(combined.encode()).hexdigest()[:10].upper()

    @staticmethod
    def validate_key(key):
        # Format: NC-[TR|AN|LT]-random-checksum
        key = key.strip().upper()
        
        # Strip common prefixes like [TR], [AN], [LT] if user copied them
        if key.startswith("[") and "]" in key:
            key = key.split("]", 1)[1].strip()

        if not key.startswith("NC-"):
            return False, "Invalid Format", None

        parts = key.split("-")
        if len(parts) != 4:
            return False, "Invalid Format", None

        type_code = parts[1]
        random_part = parts[2]
        provided_checksum = parts[3]

        expected_checksum = LicenseManager.generate_checksum(f"NC-{type_code}-{random_part}")
        
        if provided_checksum != expected_checksum:
            return False, "Invalid Checksum", None

        type_map = {"TR": "Trial (15 Days)", "AN": "Annual (1 Year)", "LT": "Lifetime"}
        return True, type_map.get(type_code, "Unknown"), type_code

    @staticmethod
    def save_license(key, type_code):
        LICENSE_PATH.parent.mkdir(parents=True, exist_ok=True)
        data = {
            "key": key,
            "type": type_code,
            "activated_at": datetime.datetime.now().isoformat()
        }
        # Simple obfuscation via base64
        encoded = base64.b64encode(json.dumps(data).encode()).decode()
        LICENSE_PATH.write_text(encoded)

    @staticmethod
    def load_license():
        if not LICENSE_PATH.exists():
            return None
        try:
            raw = LICENSE_PATH.read_text()
            data = json.loads(base64.b64decode(raw).decode())
            return data
        except Exception:
            return None

    @staticmethod
    def check_status():
        license_data = LicenseManager.load_license()
        if not license_data:
            return "MISSING", "No license found"

        ok, type_name, type_code = LicenseManager.validate_key(license_data["key"])
        if not ok:
            return "INVALID", "License key is invalid"

        activated_at = datetime.datetime.fromisoformat(license_data["activated_at"])
        now = datetime.datetime.now()
        days_passed = (now - activated_at).days

        if type_code == "TR":
            if days_passed > 15:
                return "EXPIRED", f"Trial expired ({days_passed} days passed)"
            return "OK", f"Trial: {15 - days_passed} days remaining"
        
        if type_code == "AN":
            if days_passed > 365:
                return "EXPIRED", f"Annual license expired"
            return "OK", f"Annual: {365 - days_passed} days remaining"

        return "OK", "Lifetime License"

def self_delete_app():
    """Cleans up and deletes the app from hardware (Nextcloud files and installer)."""
    cleanup_script = (
        "@echo off\n"
        "echo  [1/3] Stopping and removing containers...\n"
        f"cd /d \"{INSTALL_DIR}\" && docker compose down -v > nul 2>&1\n"
        "echo  [2/3] Deleting installation folder...\n"
        "timeout /t 3 /nobreak > nul\n"
        f"rd /s /q \"{INSTALL_DIR}\" > nul 2>&1\n"
        "echo  [3/3] Self-deleting...\n"
        "del /f /q \"%~f0\" > nul 2>&1\n"
    )
    script_path = Path(tempfile.gettempdir()) / "nc_cleanup.bat"
    script_path.write_text(cleanup_script)
    
    # Run cleanup script and exit
    subprocess.Popen(["cmd.exe", "/c", str(script_path)], shell=True)
    sys.exit(0)

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG FILE WRITERS
# ─────────────────────────────────────────────────────────────────────────────

def write_docker_compose(install_dir, ip, https_port, http_port):
    content = (
        "services:\n"
        "\n"
        "  db:\n"
        "    image: mariadb:10.11\n"
        "    container_name: nextcloud-db\n"
        "    restart: always\n"
        "    command: --transaction-isolation=READ-COMMITTED "
        "--binlog-format=ROW\n"
        "    environment:\n"
        "      MYSQL_ROOT_PASSWORD: rootpassword\n"
        "      MYSQL_PASSWORD: nextcloudpass\n"
        "      MYSQL_DATABASE: nextcloud\n"
        "      MYSQL_USER: nextclouduser\n"
        "    volumes:\n"
        "      - db:/var/lib/mysql\n"
        "\n"
        "  redis:\n"
        "    image: redis:alpine\n"
        "    container_name: nextcloud-redis\n"
        "    restart: always\n"
        "    volumes:\n"
        "      - redis:/data\n"
        "\n"
        "  app:\n"
        "    image: nextcloud:30\n"
        "    container_name: nextcloud-app\n"
        "    restart: always\n"
        "    expose:\n"
        '      - "80"\n'
        "    environment:\n"
        "      MYSQL_HOST: db\n"
        "      MYSQL_DATABASE: nextcloud\n"
        "      MYSQL_USER: nextclouduser\n"
        "      MYSQL_PASSWORD: nextcloudpass\n"
        "      REDIS_HOST: redis\n"
        "      NEXTCLOUD_MEMORY_LIMIT: 1G\n"
        "      PHP_MEMORY_LIMIT: 1G\n"
        "      PHP_UPLOAD_LIMIT: 10G\n"
        "      NEXTCLOUD_TRUSTED_DOMAINS: localhost " + ip + "\n"
        "      OVERWRITEPROTOCOL: https\n"
        "      OVERWRITECLIURL: https://" + ip + ":" + https_port + "\n"
        "    volumes:\n"
        "      - nextcloud:/var/www/html\n"
        "    depends_on:\n"
        "      - db\n"
        "      - redis\n"
        "\n"
        "  nginx:\n"
        "    image: nginx:alpine\n"
        "    container_name: nextcloud-nginx\n"
        "    restart: always\n"
        "    ports:\n"
        '      - "' + http_port + ':80"\n'
        '      - "' + https_port + ':443"\n'
        "    volumes:\n"
        "      - ./nginx/nginx.conf:/etc/nginx/nginx.conf:ro\n"
        "      - ./nginx/ssl:/etc/nginx/ssl:ro\n"
        "    depends_on:\n"
        "      - app\n"
        "\n"
        "volumes:\n"
        "  db:\n"
        "  redis:\n"
        "  nextcloud:\n"
    )
    (install_dir / "docker-compose.yml").write_text(content, encoding="utf-8")

def write_nginx_conf(install_dir):
    content = (
        "worker_processes auto;\n\n"
        "events {\n"
        "    worker_connections 1024;\n"
        "}\n\n"
        "http {\n"
        "    upstream nextcloud {\n"
        "        server app:80;\n"
        "    }\n\n"
        "    server {\n"
        "        listen 80;\n"
        "        server_name _;\n"
        "        return 301 https://$host$request_uri;\n"
        "    }\n\n"
        "    server {\n"
        "        listen 443 ssl;\n"
        "        server_name _;\n\n"
        "        ssl_certificate     /etc/nginx/ssl/nextcloud.crt;\n"
        "        ssl_certificate_key /etc/nginx/ssl/nextcloud.key;\n"
        "        ssl_protocols       TLSv1.2 TLSv1.3;\n"
        "        ssl_ciphers         HIGH:+ECDHE;\n\n"
        "        client_max_body_size 10G;\n"
        "        proxy_read_timeout   600s;\n"
        "        proxy_send_timeout   600s;\n"
        "        proxy_connect_timeout 600s;\n\n"
        '        add_header Strict-Transport-Security "max-age=15768000" always;\n'
        "        add_header X-Content-Type-Options nosniff always;\n"
        "        add_header X-Frame-Options SAMEORIGIN always;\n\n"
        "        location / {\n"
        "            proxy_pass         http://nextcloud;\n"
        "            proxy_set_header   Host $host;\n"
        "            proxy_set_header   X-Real-IP $remote_addr;\n"
        "            proxy_set_header   X-Forwarded-For "
        "$proxy_add_x_forwarded_for;\n"
        "            proxy_set_header   X-Forwarded-Proto https;\n"
        "            proxy_set_header   X-Forwarded-Host $host;\n"
        "            proxy_buffering    off;\n"
        "            proxy_request_buffering off;\n"
        "        }\n\n"
        "        location /push/ {\n"
        "            proxy_pass         http://nextcloud;\n"
        "            proxy_http_version 1.1;\n"
        "            proxy_set_header   Upgrade $http_upgrade;\n"
        "            proxy_set_header   Connection upgrade;\n"
        "            proxy_set_header   Host $host;\n"
        "            proxy_set_header   X-Forwarded-Proto https;\n"
        "        }\n"
        "    }\n"
        "}\n"
    )
    nginx_dir = install_dir / "nginx"
    nginx_dir.mkdir(parents=True, exist_ok=True)
    (nginx_dir / "nginx.conf").write_text(content, encoding="utf-8")

def write_daemon_json():
    """Configure Docker Engine daemon."""
    docker_cfg = Path("C:/ProgramData/docker/config")
    docker_cfg.mkdir(parents=True, exist_ok=True)
    cfg = {
        "hosts": ["npipe:////./pipe/docker_engine"],
        "log-driver": "json-file",
        "log-opts": {"max-size": "10m", "max-file": "3"}
    }
    (docker_cfg / "daemon.json").write_text(
        json.dumps(cfg, indent=2), encoding="utf-8"
    )

# ─────────────────────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Nextcloud Installer v" + APP_VERSION)
        self.geometry("860x600")
        self.minsize(800, 560)
        self.configure(bg=COLORS["bg"])
        self.resizable(True, True)

        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(
            "860x600+" + str((sw - 860) // 2) + "+" + str((sh - 600) // 2)
        )

        self.local_ip   = get_local_ip()
        self.https_port = tk.StringVar(value="8443")
        self.http_port  = tk.StringVar(value="8080")
        self.license_key = tk.StringVar()
        self.abort_flag = False

        self._build_ui()
        
        # Initial license check
        status, msg = LicenseManager.check_status()
        if status == "OK":
            self._show("welcome")
        elif status == "EXPIRED":
            self._show_expired(msg)
        else:
            self._show("license")

    # ── Build UI ──────────────────────────────────────────────────────────

    def _build_ui(self):
        # Sidebar
        self.sidebar = tk.Frame(self, bg=COLORS["panel"], width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Logo Section
        logo_frame = tk.Frame(self.sidebar, bg=COLORS["panel"])
        logo_frame.pack(pady=(26, 12))

        try:
            # Import Pillow inside function to handle bundling better
            from PIL import Image, ImageTk
            
            # Load and resize logo
            logo_path = resource_path("logo.png")
            
            # Additional fallback checks
            if not logo_path.exists():
                search_paths = [
                    Path(__file__).parent / "logo.png",
                    Path(__file__).parent.parent / "logo.png",
                    Path(os.path.dirname(sys.executable)) / "logo.png" if getattr(sys, 'frozen', False) else None
                ]
                for p in search_paths:
                    if p and p.exists():
                        logo_path = p
                        break

            if logo_path.exists():
                img = Image.open(logo_path)
                img.thumbnail((120, 120)) # Max size in sidebar
                self.logo_img = ImageTk.PhotoImage(img)
                
                tk.Label(logo_frame, image=self.logo_img,
                         bg=COLORS["panel"]).pack()
                self._log("  ✅  Logo loaded", "dim")
            else:
                raise FileNotFoundError(f"Logo not found at {logo_path}")
        except Exception as e:
            self._log(f"  ⚠  Logo load failure: {e}", "dim")
            self._log(f"  (Searched: {logo_path})", "dim")
            # Fallback to emoji if logo fails
            tk.Label(logo_frame, text="☁",
                     font=("Segoe UI Emoji", 36),
                     bg=COLORS["panel"],
                     fg=COLORS["accent"]).pack()

        tk.Label(self.sidebar, text="Nextcloud",
                 font=("Segoe UI", 13, "bold"),
                 bg=COLORS["panel"],
                 fg=COLORS["text"]).pack()
        tk.Label(self.sidebar,
                 text="Enterprise Installer  v" + APP_VERSION,
                 font=("Segoe UI", 8),
                 bg=COLORS["panel"],
                 fg=COLORS["muted"]).pack(pady=(0, 10))

        tk.Frame(self.sidebar, bg=COLORS["border"],
                 height=1).pack(fill="x", padx=16, pady=4)

        # Step indicators
        self.step_rows = []
        for i, name in enumerate(STEPS):
            row = tk.Frame(self.sidebar, bg=COLORS["panel"])
            row.pack(fill="x", padx=16, pady=2)
            dot = tk.Label(row, text=str(i + 1), width=2,
                           font=("Segoe UI", 8, "bold"),
                           bg=COLORS["dim"], fg=COLORS["muted"])
            dot.pack(side="left", padx=(0, 8))
            lbl = tk.Label(row, text=name,
                           font=("Segoe UI", 9),
                           bg=COLORS["panel"],
                           fg=COLORS["muted"], anchor="w")
            lbl.pack(side="left")
            self.step_rows.append((dot, lbl))

        tk.Frame(self.sidebar, bg=COLORS["border"],
                 height=1).pack(fill="x", padx=16, pady=(14, 4))
        tk.Label(self.sidebar, text="🌐  " + self.local_ip,
                 font=("Segoe UI", 8),
                 bg=COLORS["panel"],
                 fg=COLORS["muted"]).pack(pady=2)
        os_label = ("Windows " + platform.release()) if IS_WINDOWS else ("Linux " + platform.release())
        tk.Label(self.sidebar,
                 text=os_label,
                 font=("Segoe UI", 8),
                 bg=COLORS["panel"],
                 fg=COLORS["muted"]).pack()

        # Content
        self.content = tk.Frame(self, bg=COLORS["bg"])
        self.content.pack(side="left", fill="both", expand=True)

        self.pages = {}
        self._page_license()
        self._page_welcome()
        self._page_config()
        self._page_install()
        self._page_done()

    # ── Pages ─────────────────────────────────────────────────────────────

    def _page_license(self):
        p = tk.Frame(self.content, bg=COLORS["bg"])
        self.pages["license"] = p
        inner = tk.Frame(p, bg=COLORS["bg"])
        inner.pack(fill="both", expand=True, padx=38, pady=26)

        tk.Label(inner, text="Activation Required",
                 font=("Segoe UI", 17, "bold"),
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w")
        tk.Label(inner,
                 text="Please enter your license key to continue using Nextcloud.",
                 font=("Segoe UI", 10),
                 bg=COLORS["bg"], fg=COLORS["muted"]).pack(anchor="w", pady=(3, 18))

        self._field(inner, "License Key", self.license_key, 
                    "Format: NC-XXXX-XXXXX-XXXXX")

        info_lbl = tk.Label(inner, text="Contact support to purchase a key.",
                           font=("Segoe UI", 9),
                           bg=COLORS["bg"], fg=COLORS["muted"])
        info_lbl.pack(anchor="w", pady=10)

        btns = tk.Frame(p, bg=COLORS["bg"])
        btns.pack(fill="x", padx=38, pady=(0, 24))
        
        def _activate():
            key = self.license_key.get()
            ok, type_name, type_code = LicenseManager.validate_key(key)
            if ok:
                LicenseManager.save_license(key, type_code)
                self._show("welcome")
            else:
                self._alert("Invalid Key", "The license key provided is invalid or malformed.")

        self._btn(btns, "Activate License →", _activate, accent=True).pack(side="right")

    def _show_expired(self, reason):
        # Create a full-screen block for expired license
        for pg in self.pages.values(): pg.pack_forget()
        
        p = tk.Frame(self.content, bg=COLORS["bg"])
        p.pack(fill="both", expand=True)
        inner = tk.Frame(p, bg=COLORS["bg"])
        inner.pack(expand=True)

        tk.Label(inner, text="⚠️", font=("Segoe UI Emoji", 50), bg=COLORS["bg"]).pack()
        tk.Label(inner, text="License Expired",
                 font=("Segoe UI", 20, "bold"),
                 bg=COLORS["bg"], fg=COLORS["error"]).pack(pady=10)
        tk.Label(inner, text=reason,
                 font=("Segoe UI", 11),
                 bg=COLORS["bg"], fg=COLORS["text"]).pack()

        tk.Label(inner, text="\nApplication will now cleanup and exit.",
                 font=("Segoe UI", 9),
                 bg=COLORS["bg"], fg=COLORS["muted"]).pack()

        self._btn(inner, "Cleanup & Exit", self_delete_app, danger=True).pack(pady=20)

    def _page_welcome(self):
        p = tk.Frame(self.content, bg=COLORS["bg"])
        self.pages["welcome"] = p
        inner = tk.Frame(p, bg=COLORS["bg"])
        inner.pack(fill="both", expand=True, padx=38, pady=26)

        tk.Label(inner, text="Nextcloud LAN Video Conference",
                 font=("Segoe UI", 17, "bold"),
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w")
        tk.Label(inner,
                 text="Private  •  Secure  •  No account required",
                 font=("Segoe UI", 10),
                 bg=COLORS["bg"], fg=COLORS["accent"]).pack(anchor="w",
                                                             pady=(3, 18))

        grid = tk.Frame(inner, bg=COLORS["bg"])
        grid.pack(fill="x")
        feats = [
            ("🎥", "HD Video Calls",     "Up to 50 simultaneous LAN users"),
            ("🔐", "HTTPS Encrypted",    "Auto SSL — no setup needed"),
            ("🐳", "Docker Engine",      "No login — no Docker account"),
            ("⚡", "Auto Start",         "Starts on Windows boot"),
        ]
        for i, (ico, title, sub) in enumerate(feats):
            c = tk.Frame(grid, bg=COLORS["card"], padx=14, pady=11)
            c.grid(row=i // 2, column=i % 2,
                   padx=6, pady=5, sticky="nsew")
            grid.columnconfigure(i % 2, weight=1)
            tk.Label(c, text=ico,
                     font=("Segoe UI Emoji", 20),
                     bg=COLORS["card"]).grid(row=0, column=0,
                                             rowspan=2, padx=(0, 10))
            tk.Label(c, text=title,
                     font=("Segoe UI", 10, "bold"),
                     bg=COLORS["card"], fg=COLORS["text"],
                     anchor="w").grid(row=0, column=1, sticky="w")
            tk.Label(c, text=sub,
                     font=("Segoe UI", 9),
                     bg=COLORS["card"], fg=COLORS["muted"],
                     anchor="w").grid(row=1, column=1, sticky="w")

        info = tk.Frame(inner, bg=COLORS["card"], padx=16, pady=12)
        info.pack(fill="x", pady=(18, 0))
        tk.Label(info,
                 text="ℹ  What happens automatically in background:",
                 font=("Segoe UI", 9, "bold"),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
        for item in [
            "  • Installs Docker Engine silently (no login, no account)",
            "  • Generates HTTPS SSL certificate",
            "  • Downloads and starts Nextcloud + Nginx + Database",
            "  • Registers auto-start as Windows service",
        ]:
            tk.Label(info, text=item,
                     font=("Segoe UI", 9),
                     bg=COLORS["card"],
                     fg=COLORS["muted"]).pack(anchor="w")

        req = tk.Frame(inner, bg=COLORS["card"], padx=16, pady=8)
        req.pack(fill="x", pady=(8, 0))
        os_name = "Windows 10/11" if IS_WINDOWS else "Linux"
        tk.Label(req,
                 text=f"⚠  Requires: {os_name} (64-bit)  •  "
                      "8 GB RAM  •  20 GB free disk  •  Internet (first run)",
                 font=("Segoe UI", 8),
                 bg=COLORS["card"], fg=COLORS["warning"],
                 wraplength=560).pack(anchor="w")

        btns = tk.Frame(p, bg=COLORS["bg"])
        btns.pack(fill="x", padx=38, pady=(0, 24))
        self._btn(btns, "Next →",
                  lambda: self._show("config"),
                  accent=True).pack(side="right")

    def _page_config(self):
        p = tk.Frame(self.content, bg=COLORS["bg"])
        self.pages["config"] = p
        inner = tk.Frame(p, bg=COLORS["bg"])
        inner.pack(fill="both", expand=True, padx=38, pady=26)

        tk.Label(inner, text="Configuration",
                 font=("Segoe UI", 17, "bold"),
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w")
        tk.Label(inner,
                 text="Choose ports for Nextcloud (defaults work for most cases)",
                 font=("Segoe UI", 10),
                 bg=COLORS["bg"], fg=COLORS["muted"]).pack(anchor="w",
                                                            pady=(3, 18))

        self._field(inner, "HTTPS Port", self.https_port,
                    "Secure access port — default 8443")
        self._field(inner, "HTTP Port", self.http_port,
                    "Redirects to HTTPS — default 8080")

        net = tk.Frame(inner, bg=COLORS["card"], padx=16, pady=14)
        net.pack(fill="x", pady=(20, 0))
        tk.Label(net,
                 text="📡  Server IP: " + self.local_ip,
                 font=("Segoe UI", 11, "bold"),
                 bg=COLORS["card"], fg=COLORS["accent"]).pack(anchor="w")
        tk.Label(net,
                 text="After install, LAN users connect to:\n"
                      "https://" + self.local_ip + ":"
                      + self.https_port.get(),
                 font=("Segoe UI", 10),
                 bg=COLORS["card"], fg=COLORS["text"],
                 justify="left").pack(anchor="w", pady=(6, 0))

        note = tk.Frame(inner, bg=COLORS["card"], padx=16, pady=10)
        note.pack(fill="x", pady=(10, 0))
        tk.Label(note,
                 text="ℹ  Admin account is created on first browser visit.\n"
                      "   Install location: C:\\Nextcloud-LAN\\",
                 font=("Segoe UI", 9),
                 bg=COLORS["card"], fg=COLORS["muted"],
                 justify="left").pack(anchor="w")

        btns = tk.Frame(p, bg=COLORS["bg"])
        btns.pack(fill="x", padx=38, pady=(0, 24))
        self._btn(btns, "← Back",
                  lambda: self._show("welcome")).pack(side="left")
        self._btn(btns, "Install →",
                  self._start_install, accent=True).pack(side="right")

    def _page_install(self):
        p = tk.Frame(self.content, bg=COLORS["bg"])
        self.pages["install"] = p
        inner = tk.Frame(p, bg=COLORS["bg"])
        inner.pack(fill="both", expand=True, padx=38, pady=20)

        tk.Label(inner, text="Installing Nextcloud",
                 font=("Segoe UI", 17, "bold"),
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w")

        self.status_var = tk.StringVar(value="Starting...")
        tk.Label(inner, textvariable=self.status_var,
                 font=("Segoe UI", 10),
                 bg=COLORS["bg"], fg=COLORS["accent"]).pack(anchor="w",
                                                             pady=(3, 6))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Main.Horizontal.TProgressbar",
                        troughcolor=COLORS["card"],
                        background=COLORS["accent"], thickness=12)
        style.configure("Sub.Horizontal.TProgressbar",
                        troughcolor=COLORS["card"],
                        background=COLORS["success"], thickness=7)

        self.main_bar = ttk.Progressbar(
            inner, style="Main.Horizontal.TProgressbar",
            mode="determinate", maximum=100)
        self.main_bar.pack(fill="x", pady=(0, 4))

        self.sub_bar = ttk.Progressbar(
            inner, style="Sub.Horizontal.TProgressbar",
            mode="determinate", maximum=100)
        self.sub_bar.pack(fill="x", pady=(0, 4))

        self.sub_lbl = tk.Label(inner, text="",
                                font=("Consolas", 8),
                                bg=COLORS["bg"], fg=COLORS["muted"])
        self.sub_lbl.pack(anchor="w", pady=(0, 6))

        tk.Label(inner, text="Installation Log",
                 font=("Segoe UI", 9, "bold"),
                 bg=COLORS["bg"], fg=COLORS["muted"]).pack(anchor="w")

        log_frame = tk.Frame(inner, bg=COLORS["card"])
        log_frame.pack(fill="both", expand=True, pady=(3, 0))

        self.log_txt = tk.Text(
            log_frame, font=("Consolas", 9),
            bg=COLORS["card"], fg=COLORS["text"],
            insertbackground=COLORS["text"],
            relief="flat", bd=0, state="disabled", wrap="word")
        sb = ttk.Scrollbar(log_frame, command=self.log_txt.yview)
        self.log_txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_txt.pack(side="left", fill="both", expand=True,
                          padx=8, pady=6)

        self.log_txt.tag_config("ok",    foreground=COLORS["success"])
        self.log_txt.tag_config("warn",  foreground=COLORS["warning"])
        self.log_txt.tag_config("error", foreground=COLORS["error"])
        self.log_txt.tag_config("info",  foreground=COLORS["accent"])
        self.log_txt.tag_config("dim",   foreground=COLORS["muted"])

        btns = tk.Frame(p, bg=COLORS["bg"])
        btns.pack(fill="x", padx=38, pady=(0, 18))
        self._btn(btns, "Abort", self._do_abort,
                  danger=True).pack(side="right")

    def _page_done(self):
        p = tk.Frame(self.content, bg=COLORS["bg"])
        self.pages["done"] = p
        inner = tk.Frame(p, bg=COLORS["bg"])
        inner.pack(fill="both", expand=True, padx=38, pady=24)

        tk.Label(inner, text="✅",
                 font=("Segoe UI Emoji", 50),
                 bg=COLORS["bg"]).pack(pady=(8, 0))
        tk.Label(inner, text="Nextcloud Gold Ready!",
                 font=("Segoe UI", 20, "bold"),
                 bg=COLORS["bg"], fg=COLORS["success"]).pack(pady=(6, 2))
        tk.Label(inner,
                 text="Your private enterprise server is running securely on your LAN",
                 font=("Segoe UI", 10),
                 bg=COLORS["bg"], fg=COLORS["muted"]).pack()

        # ─── URL Card ─────────────────────────────────────────────────────
        url_card = tk.Frame(inner, bg=COLORS["card"], padx=20, pady=16)
        url_card.pack(fill="x", pady=20)
        tk.Label(url_card, text="🌐  Access URL",
                 font=("Segoe UI", 9, "bold"),
                 bg=COLORS["card"], fg=COLORS["muted"]).pack(anchor="w")
        self.done_url = tk.StringVar()
        url_lbl = tk.Label(url_card, textvariable=self.done_url,
                           font=("Consolas", 15, "bold"),
                           bg=COLORS["card"], fg=COLORS["accent"],
                           cursor="hand2")
        url_lbl.pack(anchor="w", pady=(4, 0))
        url_lbl.bind("<Button-1>",
                     lambda e: webbrowser.open(self.done_url.get()))
        tk.Label(url_card,
                 text="Click to open  •  Share URL with LAN users  •  SSL is auto-trusted",
                 font=("Segoe UI", 8),
                 bg=COLORS["card"], fg=COLORS["muted"]).pack(anchor="w", pady=(6, 0))

        # ─── Credentials Card ─────────────────────────────────────────────
        cred_card = tk.Frame(inner, bg=COLORS["card"], padx=20, pady=14)
        cred_card.pack(fill="x", pady=(0, 20))
        tk.Label(cred_card, text="🔑  Admin Credentials (Auto-Generated)",
                 font=("Segoe UI", 10, "bold"),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
        
        cgrid = tk.Frame(cred_card, bg=COLORS["card"])
        cgrid.pack(fill="x", pady=(8, 0))
        
        tk.Label(cgrid, text="Username:", font=("Segoe UI", 9),
                 bg=COLORS["card"], fg=COLORS["muted"]).grid(row=0, column=0, sticky="w")
        tk.Label(cgrid, text="admin", font=("Consolas", 10, "bold"),
                 bg=COLORS["card"], fg=COLORS["text"]).grid(row=0, column=1, sticky="w", padx=10)
        
        tk.Label(cgrid, text="Password:", font=("Segoe UI", 9),
                 bg=COLORS["card"], fg=COLORS["muted"]).grid(row=1, column=0, sticky="w")
        tk.Label(cgrid, text="nextcloud-admin", font=("Consolas", 10, "bold"),
                 bg=COLORS["card"], fg=COLORS["text"]).grid(row=1, column=1, sticky="w", padx=10)
        
        def _copy_p():
            self.clipboard_clear()
            self.clipboard_append("nextcloud-admin")
            self._status("Password copied to clipboard!")

        copy_btn = self._btn(cred_card, "📋 Copy Password", _copy_p)
        copy_btn.pack(side="right")
        copy_btn.config(font=("Segoe UI", 9, "bold"), pady=6)

        # ─── Next Steps ───────────────────────────────────────────────────
        steps = tk.Frame(inner, bg=COLORS["card"], padx=20, pady=14)
        steps.pack(fill="x")
        tk.Label(steps, text="Setup Video Calling (Nextcloud Talk)",
                 font=("Segoe UI", 10, "bold"),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
        for s in [
            "1.  Login with the credentials above",
            "2.  Top-right menu (User Icon) → Apps",
            "3.  Search for 'Talk' and click Enable",
            "4.  A camera icon will appear in the top navigation bar",
        ]:
            tk.Label(steps, text=s, font=("Segoe UI", 9),
                     bg=COLORS["card"], fg=COLORS["muted"],
                     anchor="w").pack(anchor="w", pady=1)

        btns = tk.Frame(p, bg=COLORS["bg"])
        btns.pack(fill="x", padx=38, pady=(0, 28))
        self._btn(btns, "📋  Copy URL",
                  self._copy_url).pack(side="right", padx=(8, 0))
        self._btn(btns, "🌐  Open Nextcloud",
                  lambda: webbrowser.open(self.done_url.get()),
                  accent=True).pack(side="right")
        self._btn(btns, "📥  Download Clients",
                  lambda: webbrowser.open("https://nextcloud.com/install/#install-clients"),
                  ).pack(side="left")

    # ── Widgets ───────────────────────────────────────────────────────────

    def _btn(self, parent, text, cmd, accent=False, danger=False):
        if danger:
            bg, hv = COLORS["error"],  "#ff6b6b"
        elif accent:
            bg, hv = COLORS["accent"], COLORS["accent2"]
        else:
            bg, hv = COLORS["card"],   COLORS["border"]
        b = tk.Button(parent, text=text, command=cmd,
                      font=("Segoe UI", 10, "bold"),
                      bg=bg, fg=COLORS["text"],
                      activebackground=hv,
                      activeforeground=COLORS["text"],
                      relief="flat", bd=0,
                      padx=20, pady=9, cursor="hand2")
        b.bind("<Enter>", lambda e: b.config(bg=hv))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _field(self, parent, label, var, tip=""):
        f = tk.Frame(parent, bg=COLORS["bg"])
        f.pack(fill="x", pady=6)
        tk.Label(f, text=label,
                 font=("Segoe UI", 10, "bold"),
                 bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w")
        if tip:
            tk.Label(f, text=tip,
                     font=("Segoe UI", 8),
                     bg=COLORS["bg"], fg=COLORS["muted"]).pack(anchor="w")
        e = tk.Entry(f, textvariable=var,
                     font=("Consolas", 11),
                     bg=COLORS["card"], fg=COLORS["text"],
                     insertbackground=COLORS["text"],
                     relief="flat", bd=0,
                     highlightthickness=1,
                     highlightcolor=COLORS["accent"],
                     highlightbackground=COLORS["border"])
        e.pack(fill="x", ipady=8, pady=(3, 0))

    def _show(self, name):
        for pg in self.pages.values():
            pg.pack_forget()
        self.pages[name].pack(fill="both", expand=True)

    def _set_step(self, idx):
        self.after(0, lambda i=idx: self._do_set_step(i))

    def _do_set_step(self, idx):
        for i, (dot, lbl) in enumerate(self.step_rows):
            if i < idx:
                dot.config(bg=COLORS["success"], fg="#fff")
                lbl.config(fg=COLORS["success"], font=("Segoe UI", 9))
            elif i == idx:
                dot.config(bg=COLORS["accent"], fg="#fff")
                lbl.config(fg=COLORS["text"],
                           font=("Segoe UI", 9, "bold"))
            else:
                dot.config(bg=COLORS["dim"], fg=COLORS["muted"])
                lbl.config(fg=COLORS["muted"], font=("Segoe UI", 9))

    def _log(self, msg, tag=""):
        def _do():
            self.log_txt.config(state="normal")
            self.log_txt.insert("end", msg + "\n", tag or "")
            self.log_txt.see("end")
            self.log_txt.config(state="disabled")
        self.after(0, _do)

    def _status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    def _mprog(self, v):
        self.after(0, lambda: self.main_bar.config(value=v))

    def _sprog(self, v, lbl=""):
        def _do():
            self.sub_bar.config(value=v)
            if lbl:
                self.sub_lbl.config(text=lbl)
        self.after(0, _do)

    def _copy_url(self):
        self.clipboard_clear()
        self.clipboard_append(self.done_url.get())

    # ── Validation ────────────────────────────────────────────────────────

    def _validate(self):
        for var, name in [(self.https_port, "HTTPS Port"),
                          (self.http_port,  "HTTP Port")]:
            val = var.get().strip()
            if not val.isdigit() or not (80 <= int(val) <= 65535):
                self._alert(name + " invalid",
                            name + " must be a number between 80–65535.")
                return False
        if self.https_port.get() == self.http_port.get():
            self._alert("Port conflict",
                        "HTTPS and HTTP ports must be different.")
            return False
        return True

    def _alert(self, title, msg):
        w = tk.Toplevel(self)
        w.title(title)
        w.configure(bg=COLORS["bg"])
        w.geometry("340x130")
        w.grab_set()
        tk.Label(w, text=msg,
                 font=("Segoe UI", 10),
                 bg=COLORS["bg"], fg=COLORS["text"],
                 wraplength=300).pack(pady=20)
        self._btn(w, "OK", w.destroy, accent=True).pack()

    # ── Install ───────────────────────────────────────────────────────────

    def _start_install(self):
        if not self._validate():
            return
        self.abort_flag = False
        self._show("install")
        threading.Thread(target=self._run_install, daemon=True).start()

    def _do_abort(self):
        self.abort_flag = True
        self._log("⚠  Aborting...", "warn")

    # ─────────────────────────────────────────────────────────────────────
    #  INSTALL SEQUENCE
    # ─────────────────────────────────────────────────────────────────────

    def _run_install(self):
        try:
            https_port = self.https_port.get().strip()
            http_port  = self.http_port.get().strip()
            ip         = self.local_ip

            # Step 0 — Prepare Branding
            try:
                INSTALL_DIR.mkdir(parents=True, exist_ok=True)
                from PIL import Image
                logo_path = resource_path("logo.png")
                if not logo_path.exists():
                    logo_path = Path(__file__).parent.parent / "logo.png"
                
                if logo_path.exists():
                    img = Image.open(logo_path)
                    img.save(INSTALL_DIR / "logo.ico", format="ICO", sizes=[(256, 256)])
                    self._log("  ✅  Branding assets ready", "dim")
            except Exception as e:
                self._log(f"  ⚠  Icon generation skipped: {e}", "dim")

            # Step 0 — System check
            self._set_step(0)
            if not self._check_system():
                return
            if self.abort_flag: return

            # Step 1 — Docker Engine
            self._set_step(1)
            if not self._install_docker_engine():
                self._alert("Docker Install Error", "Failed to install Docker in WSL. See log for details.")
                return
            if self.abort_flag: return

            # Step 2 — Start Docker
            self._set_step(2)
            if not self._start_docker():
                self._alert("Docker Start Error", "Docker daemon failed to start. Please restart your PC or check WSL status.")
                return
            if self.abort_flag: return

            # Step 3 — SSL Certificate
            self._set_step(3)
            self._mprog(38)
            self._do_ssl(ip)
            if self.abort_flag: return

            # Step 4 — Config files
            self._set_step(4)
            self._mprog(48)
            self._do_config(ip, https_port, http_port)
            if self.abort_flag: return

            # Step 5 — Pull images
            self._set_step(5)
            self._mprog(55)
            if not self._do_pull():
                self._alert("Download Error", "Failed to download Docker images. Check your internet connection and try again.")
                return
            if self.abort_flag: return

            # Step 6 — Start services
            self._set_step(6)
            self._mprog(80)
            if not self._do_start(ip, https_port):
                return

            # Step 7 — Zero-Config Setup
            self._set_step(7)
            self._mprog(88)
            self._do_nextcloud_setup(ip, https_port)
            if self.abort_flag: return

            # Step 8 — Done
            self._set_step(8)
            self._mprog(100)
            self._mprog(100)
            self._sprog(100)

            url = "https://" + ip + ":" + https_port
            
            # Trust SSL
            self._status("Trusting SSL certificate...")
            ssl_dir = INSTALL_DIR / "nginx" / "ssl"
            cert_path = ssl_dir / "nextcloud.crt"
            ok, msg = trust_ssl_cert(cert_path)
            if ok:
                self._log("  ✅  SSL Certificate trusted globally", "ok")
            else:
                self._log("  ⚠  Failed to trust SSL: " + msg, "warn")

            # Create Shortcut
            if create_desktop_shortcut(url, icon_path=str(INSTALL_DIR / "logo.ico")):
                self._log("  ✅  Desktop shortcut created", "ok")

            self._log("\n══════════════════════════════════════════", "ok")
            self._log("  ✅  DONE!   " + url, "ok")
            self._log("══════════════════════════════════════════\n", "ok")
            self._status("✅  Installation complete!")

            self.done_url.set(url)
            self.after(1500, lambda: self._show("done"))
            self.after(2500, lambda: webbrowser.open(url))

        except Exception as e:
            import traceback
            self._log("\n❌  Unexpected error: " + str(e), "error")
            self._log(traceback.format_exc(), "error")
            self._status("Installation failed — see log")

    # ── Steps ─────────────────────────────────────────────────────────────

    def _check_system(self):
        self._log("─── System Check ─────────────────────────", "info")
        self._status("Checking system...")
        self._mprog(2)

        # Windows Server vs Desktop logic
        if IS_WINDOWS:
            if is_windows_server():
                self._log("  🖥️  Windows Server detected", "info")
                self._status("Enabling Server Features...")
                enable_server_containers()
            else:
                self._log("  💻  Windows Desktop detected", "info")
                self._status("Checking WSL2...")
                # Pass our logger to the function so user sees progress
                ok, msg = check_and_enable_wsl(log_cb=self._log)
                if not ok:
                    self._log(f"  ⚠  {msg}", "warn")

        if IS_WINDOWS:
            major, _ = win_version()
            if major < 10:
                self._log("  ❌  Windows 10 or higher required.", "error")
                self._status("Windows 10+ required — cannot install")
                return False
            self._log("  ✅  Windows " + platform.release(), "ok")
        else:
            self._log("  ✅  Linux Native (" + platform.release() + ")", "ok")

        # Check disk space
        drive = "C:/" if IS_WINDOWS else "/"
        try:
            free_gb = shutil.disk_usage(drive).free / 1024 ** 3
            self._log("  Free disk: " + str(round(free_gb, 1)) + " GB", "dim")
            if free_gb < 5:
                self._log("  ⚠  Low disk — need at least 5 GB", "warn")
        except:
            pass

        # Free up ports (Windows only for now)
        if IS_WINDOWS:
            for port in [self.https_port.get(), self.http_port.get()]:
                code, out, _ = run_cmd("netstat -ano | findstr :" + port)
                if out.strip():
                    self._log("  Freeing port " + port + "...", "dim")
                    for line in out.strip().splitlines():
                        parts = line.split()
                        if parts:
                            pid = parts[-1]
                            if pid.isdigit() and pid != "0":
                                run_cmd("taskkill /F /PID " + pid + " >nul 2>&1")
        self._log("  ✅  System check passed\n", "ok")
        self._mprog(8)
        return True

    def _install_docker_engine(self):
        self._log("\n─── Installing Docker (WSL) ──────────────", "info")
        self._status("Checking Linux subsystem...")
        
        # 1. Use the refined check
        ok, msg = check_and_enable_wsl(log_cb=self._log)
        if not ok:
            # Only prompt if it's NOT installed at all
            self._log("  ! WSL not ready: " + msg, "warn")
            self._status("Enabling WSL features...")
            
            # Show a clear popup for restart
            self._log("  ! A system restart is REQUIRED to finish WSL setup.", "error")
            if IS_WINDOWS:
                res = ctypes.windll.user32.MessageBoxW(
                    0, 
                    "Windows WSL features have been enabled, but a RESTART is required to finish.\n\n"
                    "Would you like to restart your computer now?\n"
                    "(Please save your work first!)",
                    "Restart Required",
                    0x24 # YESNO + ICONQUESTION
                )
            else:
                from tkinter import messagebox
                res = 6 if messagebox.askyesno("Restart Required", "A restart is recommended. Restart now?") else 7
            
            if res == 6: # IDYES
                self._log("  Restarting system in 5 seconds...", "warn")
                run_cmd("shutdown /r /t 5")
            else:
                self._log("  Please restart manually and run this installer again.", "info")
            
            return False

        self._log("  ✅  WSL kernel and features ready", "ok")

        # 2. Check for Distro
        self._log("  Checking for Linux distro...", "dim")
        distro = get_wsl_distro()
        # Double check with wsl -l -v
        _, out, _ = run_cmd("wsl -l -v")
        if "Ubuntu" not in out:
            self._log("  ! Ubuntu not found, installing (takes 5m+)...", "warn")
            self._status("Installing Ubuntu distro...")
            
            # Using Popen to allow background check and status updates
            cmd = ["wsl", "--install", "-d", "Ubuntu", "--no-launch"]
            try:
                self._log("  Starting wsl --install...", "dim")
                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=0x08000000 if IS_WINDOWS else 0)
                
                # Wait loop with feedback
                for i in range(150): # Up to 12.5 minutes polling
                    if self.abort_flag:
                        process.terminate()
                        return False
                    
                    # Check if process finished
                    ret = process.poll()
                    if ret is not None:
                        stdout, stderr = process.communicate()
                        if ret == 0:
                            self._log("  ✅  Distro installation command finished", "ok")
                        elif "0x4294967295" in (stderr + stdout) or "already exists" in (stderr + stdout).lower():
                            self._log("  ✅  Ubuntu already exists, skipping install", "ok")
                        else:
                            # Detect specific network error 0x80072efe
                            if "0x80072efe" in (stderr + stdout):
                                self._log("  ❌  Network Error (0x80072efe): Connection interrupted.", "error")
                                self._log("      Please check your internet and disable any VPN/Firewall.", "warn")
                                return False
                            self._log(f"  ⚠  Installation finished with code {ret}: {stderr or stdout}", "warn")
                        break
                    
                    # Periodically check if "Ubuntu" appeared in wsl -l -v even if command is still running
                    c_check, o_check, _ = run_cmd("wsl -l -v")
                    if "Ubuntu" in o_check:
                        self._log("  ✅  Ubuntu detected in system list", "ok")
                        break
                        
                    self._status(f"Installing Ubuntu... {i*5}s (Please don't close)")
                    time.sleep(5)
                else:
                    self._log("  ⚠  Installation taking too long. Check 'wsl --status' manually.", "warn")
            except Exception as e:
                self._log(f"  ❌  Failed to start installation: {e}", "error")
                return False
            
            # Wait for distro to be registered and ready
            self._log("  Waiting for distro registration...", "dim")
            for i in range(40):
                distro = get_wsl_distro() # Refresh distro name
                code, out, _ = run_cmd("wsl -l -v")
                if "Ubuntu" in out:
                    # Distro exists, check readiness
                    c_check, _, _ = run_cmd(f"wsl -d {distro} -u root bash -c 'id'", timeout=15)
                    if c_check == 0:
                        self._log(f"  ✅  Ubuntu ({distro}) ready", "ok")
                        break
                time.sleep(10)
                self._sprog(int(i*100/40), f"Waiting for Linux... {i*10}s")
            else:
                self._log("  ❌  Ubuntu installation timed out or failed.", "error")
                return False

        # 3. Install Docker in WSL
        self._log(f"  Updating {distro} and installing Docker...", "dim")
        self._status(f"Installing Docker in {distro}...")
        scripts = [
            'apt-get update',
            'apt-get install -y docker.io docker-compose-v2',
            'service docker start'
        ]
        total = len(scripts)
        for i, s in enumerate(scripts):
            self._sprog(int((i+1)*100/total), f"Linux Setup: {s}")
            code, out, err = run_universal_cmd(s, timeout=900)
            if code != 0:
                self._log(f"  ❌  Failed step: {s}", "error")
                self._log(f"  Command Error: {err[:500] if err else out[:500]}", "error")
                # Fallback for network issues
                if "Resolve" in (err or ""):
                    self._log("  ⚠  Network issue detected, trying DNS fix...", "warn")
                    run_universal_cmd('echo "nameserver 8.8.8.8" > /etc/resolv.conf', timeout=10)
                    code, out, err = run_universal_cmd(s, timeout=600)
                    if code == 0: continue
                return False

        self._log("  ✅  Docker Engine ready in WSL", "ok")
        self._mprog(30)
        return True

    def _start_docker(self):
        self._log("\n─── Starting Docker Service ─────────────", "info")
        self._status("Starting Docker daemon in WSL...")
        
        # Try primary start method
        run_universal_cmd("service docker start", timeout=60)
        
        for i in range(15): # Increased wait time to 45s
            self._sprog(int(i*100/15), f"Waiting for Docker... {i*3}s")
            if is_docker_running():
                self._log("  ✅  Docker is alive!", "ok")
                self._mprog(40)
                return True
            
            # Try alternative start method if hanging
            if i == 5:
                self._log("  Trying alternative start method...", "dim")
                run_universal_cmd("dockerd > /dev/null 2>&1 &", timeout=5)
                
            time.sleep(3)
            
        self._log("  ❌  Docker failed to start within 45 seconds.", "error")
        return False

    def _do_ssl(self, ip):
        self._log("─── SSL Certificate ──────────────────────", "info")
        self._status("Generating SSL certificate...")
        self._sprog(0, "Generating certificate...")

        ssl_dir = INSTALL_DIR / "nginx" / "ssl"
        ssl_dir.mkdir(parents=True, exist_ok=True)
        cert_path = ssl_dir / "nextcloud.crt"
        key_path  = ssl_dir / "nextcloud.key"

        ok, msg = gen_ssl_cert(ip, cert_path, key_path)
        if ok:
            self._log("  ✅  SSL certificate created (" + msg + ")", "ok")
        else:
            self._log("  ❌  SSL failed: " + msg, "error")
            self._log("  Continuing — HTTPS may not work", "warn")

        self._sprog(100)
        self._mprog(45)

    def _do_config(self, ip, https_port, http_port):
        self._log("\n─── Configuration Files ──────────────────", "info")
        self._status("Writing configuration...")

        INSTALL_DIR.mkdir(parents=True, exist_ok=True)
        write_docker_compose(INSTALL_DIR, ip, https_port, http_port)
        self._log("  ✅  docker-compose.yml", "ok")
        write_nginx_conf(INSTALL_DIR)
        self._log("  ✅  nginx/nginx.conf", "ok")
        self._log("  📁  " + str(INSTALL_DIR), "dim")
        self._mprog(52)

    def _do_pull(self) -> bool:
        self._log("\n─── Pulling Docker Images ────────────────", "info")
        self._status("Downloading Linux images (~500 MB)...")
        self._log("  This may take 5–15 minutes...", "warn")
        
        wsl_dir = to_wsl_path(INSTALL_DIR)
        cmd = f'cd {wsl_dir} && docker compose pull'
        code, out, err = run_universal_cmd(cmd, timeout=1800)
        
        if code == 0:
            self._log("  ✅  All images downloaded", "ok")
            self._sprog(100)
            self._mprog(78)
            return True
        else:
            self._log("  ❌  Pull failed!", "error")
            self._log(f"  Error: {err[:300] if err else out[:300]}", "error")
            self._log("  Check internet and try: wsl --shutdown", "warn")
            return False

    def _do_start(self, ip, https_port):
        self._log("\n─── Starting Nextcloud ───────────────────", "info")
        self._status("Starting containers in WSL...")

        wsl_dir = to_wsl_path(INSTALL_DIR)
        
        # Cleanup any old ones
        run_universal_cmd("docker rm -f nextcloud-nginx nextcloud-app nextcloud-db nextcloud-redis > /dev/null 2>&1")

        # Start
        cmd = f"cd {wsl_dir} && docker compose up -d"
        code, out, err = run_universal_cmd(cmd, timeout=300)
        
        if code != 0:
            self._log("  ❌  Failed to start: " + err[:300], "error")
            return False

        self._log("  ✅  Containers are running", "ok")

        # LAN Connectivity: WSL2 Port Proxy
        if IS_WINDOWS:
            self._status("Configuring LAN access...")
            self._log("  Setting up network bridge/proxy...", "dim")
            
            # Find the internal WSL IP (eth0)
            code, wsl_ip, _ = run_cmd("wsl -u root hostname -I")
            wsl_ip = wsl_ip.split()[0] if wsl_ip.strip() else ""
            
            if wsl_ip:
                self._log(f"  Internal bridge: {wsl_ip}", "dim")
                h_port = self.http_port.get()
                s_port = https_port # Already a string
                
                # NETSH Port Proxy: Forward Host IP -> WSL IP
                # Clear old ones first to prevent conflicts
                run_cmd(f'netsh interface portproxy reset')
                
                # Add proxy for HTTPS
                run_cmd(f'netsh interface portproxy add v4tov4 listenport={s_port} connectport=443 connectaddress={wsl_ip}')
                # Add proxy for HTTP (redirects)
                run_cmd(f'netsh interface portproxy add v4tov4 listenport={h_port} connectport=80 connectaddress={wsl_ip}')
                
                # Firewall rules (Netsh)
                run_cmd(f'netsh advfirewall firewall delete rule name="Nextcloud-HTTPS" >nul 2>&1')
                run_cmd(f'netsh advfirewall firewall add rule name="Nextcloud-HTTPS" protocol=TCP dir=in localport={s_port} action=allow')
                run_cmd(f'netsh advfirewall firewall delete rule name="Nextcloud-HTTP" >nul 2>&1')
                run_cmd(f'netsh advfirewall firewall add rule name="Nextcloud-HTTP" protocol=TCP dir=in localport={h_port} action=allow')
                
                self._log("  ✅  LAN Port Proxy active", "ok")
            else:
                self._log("  ⚠  Could not detect WSL IP — LAN access may fail", "warn")

            # Register auto-start on boot
            self._log("  Registering auto-start on boot...", "dim")
            distro = get_wsl_distro()
            # More robust relaunch with portproxy persistence would need a scheduled task 
            # that runs a script, but for now we keep it simple.
            wsl_launch = f"service docker start && cd {wsl_dir} && docker compose up -d"
            safe_launch = wsl_launch.replace('"', '\\"')
            wsl_up_cmd = f'wsl -d {distro} -u root bash -c "{safe_launch}"'
            
            run_cmd(
                f'schtasks /create /tn "NextcloudAutoStart" /tr "{wsl_up_cmd}" /sc onstart /ru SYSTEM /f'
            )
            self._log("  ✅  Auto-start registered", "ok")

        else:
            self._log("  (Skipping Windows auto-start/firewall on Linux)", "dim")

        # Wait for HTTP response - HARDENED HEALTH CHECK
        self._log(f"\n  Checking site health (https://{ip}:{https_port})...", "dim")
        url = f"https://{ip}:{https_port}"
        ctx = ssl_module.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode    = ssl_module.CERT_NONE

        responding = False
        for i in range(50): # 250 seconds max
            if self.abort_flag: return False
            self._sprog(int((i+1)*100/50), f"Waiting for Web... { (i+1)*5 }s")
            try:
                # Use a very short timeout for the attempt
                with urllib.request.urlopen(url, context=ctx, timeout=3) as r:
                    if r.status in (200, 302):
                        self._log("  ✅  Nextcloud is responding!", "ok")
                        responding = True
                        break
            except Exception:
                pass
            time.sleep(5)
        
        if not responding:
            self._log("  ❌  Nextcloud failed to respond. Check Docker logs.", "error")
            self._log("      HINT: Ensure no other web server is on port 8443.", "warn")
            self._log("      HINT: Check if Docker is running in WSL (wsl --list -v).", "warn")
            self._status("Health check failed — site unreachable")
            return False

        self._mprog(90)
        return True

    def _do_nextcloud_setup(self, ip, https_port):
        self._log("\n─── Nextcloud Auto-Config ────────────────", "info")
        self._status("Automating setup inside WSL...")

        admin_user = "admin"
        admin_pass = "nextcloud-admin"

        # Check if already installed
        code, out, _ = run_universal_cmd("docker exec --user www-data nextcloud-app php occ status")
        if "installed: true" in out:
            self._log("  ✅  Nextcloud already initialized", "ok")
            return

        # 1. Wait for DB
        self._log("  Waiting for database...", "dim")
        for i in range(20):
            code, out, _ = run_universal_cmd("docker exec nextcloud-db mariadb-admin ping -prootpassword")
            if "mysqld is alive" in out:
                break
            time.sleep(5)
        
        # 2. Install
        self._log("  Running occ install...", "dim")
        run_universal_cmd(
            f'docker exec --user www-data nextcloud-app php occ maintenance:install '
            f'--database "mysql" --database-name "nextcloud" '
            f'--database-user "nextclouduser" --database-pass "nextcloudpass" '
            f'--database-host "db" --admin-user "{admin_user}" --admin-pass "{admin_pass}"'
        )

        # 3. Config
        run_universal_cmd(f'docker exec --user www-data nextcloud-app php occ config:system:set trusted_domains 1 --value="{ip}"')
        run_universal_cmd(f'docker exec --user www-data nextcloud-app php occ config:system:set overwrite.cli.url --value="https://{ip}:{https_port}"')
        run_universal_cmd('docker exec --user www-data nextcloud-app php occ config:system:set overwriteprotocol --value="https"')

        self._log("\n  ==========================================", "ok")
        self._log("   SUCCESS! Nextcloud Enterprise is Ready", "ok")
        self._log("  ==========================================\n", "ok")

        # 4. Talk
        self._log("  Enabling Talk...", "dim")
        run_universal_cmd("docker exec --user www-data nextcloud-app php occ app:install talk")
        run_universal_cmd("docker exec --user www-data nextcloud-app php occ config:app:set talk stun_servers --value='[\"stun.nextcloud.com:443\"]'")

        self._log("  ✅  Setup complete!", "ok")
        (INSTALL_DIR / "CREDENTIALS.txt").write_text(f"URL: https://{ip}:{https_port}\nUser: {admin_user}\nPass: {admin_pass}\n")


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Prevent infinite relaunch loop
    has_admin_flag = "--admin" in sys.argv
    
    if not is_admin():
        if IS_WINDOWS:
            if has_admin_flag:
                # We already tried to relaunch as admin but failed
                ctypes.windll.user32.MessageBoxW(0, "Failed to get Administrator privileges. Please right-click and 'Run as Administrator'.", "Error", 0x10)
                sys.exit(1)
                
            try:
                # Add a flag to arguments to detect if we've already tried relaunching
                args = sys.argv + ["--admin"]
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable,
                    " ".join('"' + a + '"' for a in args),
                    None, 1
                )
            except Exception:
                pass
            sys.exit(0)
        else:
            # Linux: We don't auto-relaunch with sudo for now to avoid complexity,
            # but we show a warning.
            print("! Running as non-root user. Some installation steps may require 'sudo'.")

    app = App()
    app.mainloop()
