"""
ensure_venv_windows.py — Zero-fuss venv bootstrapper for Windows

Drop this file next to your entry script and add AT THE VERY TOP of your entry script:

    from ensure_venv_windows import venv_wizard
    venv_wizard()

Only edit the CONFIG BLOCK below (usually just REQUIRED_PACKAGES). The rest is drag‑and‑drop.

What this does on Windows:
  • Picks a Python interpreter (defaults to 3.12) — installs it if missing
  • Creates/reuses .venv and makes sure it matches the selected Python version
  • Upgrades pip/setuptools/wheel
  • Installs your REQUIRED_PACKAGES (with special handling for GeoPandas to avoid GDAL builds)
  • Relaunches your script inside the venv so the rest of your code just works

No py launcher required. Works with/without conda on PATH.
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Iterable, List, Dict, Optional, Tuple

# =============================
# CONFIG — EDIT PER REPOSITORY
# =============================
REQUIRED_PACKAGES: List[str] = [
    "customtkinter",
    "xlwings",
    "python-dateutil",
    "pywin32"
    # Example packages — replace with what your script needs:
    # "requests>=2.32.0",
    # "pandas==2.2.3",
]

# Where to put the virtualenv
VENV_DIR = Path(".venv")

# Upgrade pip/setuptools/wheel inside the venv
UPGRADE_PIP = True

# Prefer this Python minor for the venv (None = use current interpreter)
TARGET_PYTHON_VERSION: Optional[str] = "3.12"  # e.g., "3.12" or None

# If TARGET_PYTHON_VERSION not found, try to install it automatically
AUTO_INSTALL_PYTHON: bool = True

# How to auto-install Python if missing: "winget" | "python_org" | "choco"
PYTHON_INSTALL_METHOD: str = "winget"

# If using python.org, you must give an exact patch version (with installer available)
# Python 3.12 is now source-only for later patches, so use 3.12.10 which still ships an .exe
PYTHON_ORG_EXACT_VERSION: Optional[str] = "3.12.10"

# GeoPandas smart stack: these constraints prefer pre-built wheels on Windows
DEFAULT_GEO_STACK: Dict[str, str] = {
    "shapely": ">=2.0,<3",
    "pyproj": ">=3.6,<4",
    "pyogrio": ">=0.10.0",
}

# Fail fast if any install step fails (recommended). If False, will attempt best effort.
STRICT_INSTALL = True

# =============================
# INTERNALS — NO EDITS BELOW
# =============================

class EnsureVenvError(RuntimeError):
    pass


def _print(msg: str) -> None:
    print(f"[ensure_venv] {msg}")


def _assert_windows() -> None:
    if os.name != "nt":
        raise EnsureVenvError("Windows-only (expects Scripts\python.exe)")


def _venv_python_path(venv_dir: Path) -> Path:
    return (venv_dir / "Scripts" / "python.exe").resolve()


def _pkg_name(spec: str) -> str:
    s = spec.strip()
    for sep in ("==", ">=", "<=", "~=", "!=", ">", "<", "[", " "):
        i = s.find(sep)
        if i > 0:
            return s[:i].strip().lower()
    return s.lower()


def _normalize_minor(ver: Optional[str | float | int]) -> Optional[str]:
    if ver is None:
        return None
    s = str(ver).strip()
    if not s:
        return None
    parts = s.split(".")
    if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
        return f"{parts[0]}.{parts[1]}"
    if s.isdigit() and len(s) >= 2:
        return f"{s[0]}.{s[1:]}"
    if s.isdigit():
        return f"{s}.0"
    return s


def _signature(packages: Iterable[str], py_minor: str) -> str:
    h = hashlib.sha256()
    norm = tuple(sorted((packages or []), key=str.lower))
    h.update(json.dumps(norm).encode())
    h.update(py_minor.encode())
    for k in ("PIP_INDEX_URL", "PIP_EXTRA_INDEX_URL", "PIP_TRUSTED_HOST"):
        h.update((os.environ.get(k, "") or "").encode())
    return h.hexdigest()


def _load_state(state_file: Path) -> dict:
    try:
        return json.loads(state_file.read_text())
    except Exception:
        return {}


def _save_state(state_file: Path, data: dict) -> None:
    state_file.parent.mkdir(parents=True, exist_ok=True)
    state_file.write_text(json.dumps(data, indent=2))


def _run(cmd: List[str], env: Optional[dict] = None) -> None:
    _print("$ " + " ".join(cmd))
    try:
        subprocess.check_call(cmd, env=env)
    except subprocess.CalledProcessError as e:
        if STRICT_INSTALL:
            raise
        _print(f"Command failed with exit code {e.returncode}: {' '.join(cmd)}")


def _run_capture(cmd: List[str]) -> Tuple[int, str]:
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True)
        return 0, out
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        return 1, getattr(e, "output", "") or ""


def _detect_minor(python_path: Path) -> Optional[str]:
    if not python_path.exists():
        return None
    rc, out = _run_capture([str(python_path), "-c", "import sys;print(f'{sys.version_info[0]}.{sys.version_info[1]}')"])
    return out.strip() if rc == 0 else None


def _py_launcher_locate(minor: str) -> Optional[Path]:
    minor = _normalize_minor(minor) or ""
    if shutil.which("py") is None:
        return None
    rc, out = _run_capture(["py", f"-{minor}", "-c", "import sys;print(sys.executable)"])
    if rc == 0:
        p = Path(out.strip().splitlines()[-1]).resolve()
        if p.exists():
            return p
    return None


def _common_install_paths(minor: str) -> List[Path]:
    minor = _normalize_minor(minor) or ""
    mm = minor.replace(".", "")  # e.g., 3.12 -> 312
    local = Path.home() / "AppData" / "Local" / "Programs" / "Python" / f"Python{mm}"
    program_files = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / f"Python{mm}"
    return [p for p in (local / "python.exe", program_files / "python.exe") if p.exists()]


def _locate_python_minor(minor: str) -> Optional[Path]:
    minor = _normalize_minor(minor) or ""
    # If current interpreter already matches, use it
    if minor == f"{sys.version_info.major}.{sys.version_info.minor}":
        return Path(sys.executable).resolve()
    for finder in (_py_launcher_locate, _common_install_paths):
        if finder is _common_install_paths:
            paths = finder(minor)
            if paths:
                return paths[0]
        else:
            p = finder(minor)
            if p:
                return p
    return None


def _install_python_minor(minor: str) -> Optional[Path]:
    minor = _normalize_minor(minor) or ""
    _print(f"Attempting to install Python {minor} (method={PYTHON_INSTALL_METHOD})…")

    if PYTHON_INSTALL_METHOD == "winget":
        rc, _ = _run_capture(["winget", "install", "-e", "--id", f"Python.Python.{minor}", "--silent"])
        if rc == 0:
            return _locate_python_minor(minor)
        _print("winget install failed or not available.")

    if PYTHON_INSTALL_METHOD == "choco":
        pkg = f"python{minor.replace('.', '')}"  # e.g., python312
        rc, _ = _run_capture(["choco", "install", pkg, "-y"])
        if rc != 0:
            rc, _ = _run_capture(["choco", "install", "python", "-y"])  # generic fallback
        if rc == 0:
            return _locate_python_minor(minor)
        _print("chocolatey install failed or not available.")

    if PYTHON_INSTALL_METHOD == "python_org":
        # Need an exact patch version that ships an .exe
        exact = PYTHON_ORG_EXACT_VERSION
        if not exact or not exact.startswith(minor + "."):
            _print("python_org requires PYTHON_ORG_EXACT_VERSION like '3.12.10'. Skipping.")
        else:
            arch = "amd64"
            url = f"https://www.python.org/ftp/python/{exact}/python-{exact}-{arch}.exe"
            with tempfile.TemporaryDirectory() as td:
                inst = Path(td) / f"python-{exact}-{arch}.exe"
                try:
                    import urllib.request
                    _print(f"Downloading {url} …")
                    urllib.request.urlretrieve(url, inst)
                    _run([str(inst), "/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_launcher=1"])
                    return _locate_python_minor(minor)
                except Exception as e:
                    _print(f"python.org install failed: {e}")

    return None


def _select_interpreter() -> Path:
    desired = _normalize_minor(TARGET_PYTHON_VERSION)
    if not desired:
        return Path(sys.executable).resolve()
    found = _locate_python_minor(desired)
    if found:
        return found
    if AUTO_INSTALL_PYTHON:
        maybe = _install_python_minor(desired)
        if maybe:
            return maybe
        _print("Unable to auto-install requested Python; falling back to current interpreter.")
    return Path(sys.executable).resolve()


def _pip_install(venv_python: Path, pkgs: List[str], only_binary: Optional[List[str]] = None) -> None:
    if not pkgs:
        return
    env = dict(os.environ)
    install_cmd = [str(venv_python), "-m", "pip", "install"]
    # Index config
    if env.get("PIP_INDEX_URL"):
        install_cmd += ["--index-url", env["PIP_INDEX_URL"]]
    extra = (env.get("PIP_EXTRA_INDEX_URL", "").strip() or "")
    if extra:
        for url in extra.split():
            install_cmd += ["--extra-index-url", url]
    trusted = (env.get("PIP_TRUSTED_HOST", "").strip() or "")
    if trusted:
        for host in trusted.split():
            install_cmd += ["--trusted-host", host]
    # Force wheels for heavy geo stack
    if only_binary:
        current = env.get("PIP_ONLY_BINARY", "").strip(", ")
        merged = ",".join(sorted(set([*(current.split(",") if current else []), *only_binary])))
        env["PIP_ONLY_BINARY"] = merged
        _print(f"Using wheels only for: {merged}")
    install_cmd += pkgs
    _run(install_cmd, env=env)


def _with_geopandas_strategy(packages: List[str]) -> List[str]:
    names = {_pkg_name(p): p for p in packages}
    if "geopandas" not in names:
        return packages
    _print("GeoPandas detected — installing geospatial stack first (wheels).")
    prereqs: List[str] = []
    for base, default_spec in DEFAULT_GEO_STACK.items():
        prereqs.append(names.get(base, f"{base}{default_spec}"))
    gp_spec = names["geopandas"]
    excluded = {"shapely", "pyproj", "pyogrio", "geopandas"}
    others = [spec for spec in packages if _pkg_name(spec) not in excluded]
    return prereqs + [gp_spec] + others


def _install_packages(venv_python: Path, packages: List[str]) -> None:
    if not packages:
        return
    ordered = _with_geopandas_strategy(list(packages))
    names = [_pkg_name(p) for p in ordered]
    geo_targets = [n for n in ("pyogrio", "shapely", "pyproj", "geopandas") if n in names]
    geo_segment = [p for p in ordered if _pkg_name(p) in set(geo_targets)]
    rest = [p for p in ordered if p not in geo_segment]
    if geo_segment:
        _pip_install(venv_python, geo_segment, only_binary=geo_targets)
    if rest:
        _pip_install(venv_python, rest)


def venv_wizard() -> None:
    """Create venv on Windows, install REQUIRED_PACKAGES, and re-exec under that venv.

    Place `from ensure_venv_windows import venv_wizard; venv_wizard()` at the very top of your entry script.
    """
    _assert_windows()

    # Choose interpreter and ensure venv matches it
    base_python = _select_interpreter()
    desired_minor = _detect_minor(base_python) or f"{sys.version_info.major}.{sys.version_info.minor}"

    venv_python = _venv_python_path(VENV_DIR)
    existing_minor = _detect_minor(venv_python)

    if existing_minor and existing_minor != desired_minor:
        _print(f"Existing venv is Python {existing_minor}, desired is {desired_minor}. Rebuilding…")
        shutil.rmtree(VENV_DIR, ignore_errors=True)

    # Create venv if needed
    if not venv_python.exists():
        _print(f"Creating venv at {VENV_DIR.resolve()} using {base_python}")
        _run([str(base_python), "-m", "venv", str(VENV_DIR)])
        venv_python = _venv_python_path(VENV_DIR)

    # Upgrade pip et al.
    if UPGRADE_PIP:
        _run([str(venv_python), "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel"])

    # Install deps when signature changes
    state_file = VENV_DIR / ".ensure_venv" / "state.json"
    sig = _signature(REQUIRED_PACKAGES, py_minor=desired_minor)
    state = _load_state(state_file)
    if state.get("signature") != sig:
        _print("Installing Python dependencies…")
        _install_packages(venv_python, REQUIRED_PACKAGES)
        _save_state(state_file, {"signature": sig})
    else:
        _print("Dependencies up to date.")

    # Relaunch under the venv's Python if we're not already inside it
    if Path(sys.executable).resolve() != venv_python:
        _print("Relaunching under venv Python…")
        os.execv(str(venv_python), [str(venv_python)] + sys.argv)


if __name__ == "__main__":
    venv_wizard()
