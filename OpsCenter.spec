# OpsCenter.spec
# NOTE: patch_config.ps1 and patch_rollback.ps1 have been merged into
# Invoke-ModernPatch.ps1 — only one PowerShell file is needed now.

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_dynamic_libs
import os, glob

block_cipher = None

numpy_binaries  = collect_dynamic_libs('numpy')
streamlit_datas = collect_data_files("streamlit", include_py_files=True)

def collect_dist_info(package_name):
    import site
    results = []
    for site_dir in site.getsitepackages():
        for pattern in [
            os.path.join(site_dir, f"{package_name}-*.dist-info"),
            os.path.join(site_dir, f"{package_name.replace('-','_')}-*.dist-info"),
        ]:
            for m in glob.glob(pattern):
                if os.path.isdir(m):
                    results.append((m, os.path.basename(m)))
    return results

extra_datas = []
for pkg in ["streamlit", "click", "altair", "pandas", "pyarrow",
            "requests", "urllib3", "tornado", "packaging",
            "importlib_metadata", "attrs", "jsonschema", "narwhals"]:
    extra_datas += collect_dist_info(pkg)

hidden = (
    collect_submodules("streamlit")
    + collect_submodules("streamlit.web")
    + collect_submodules("streamlit.runtime")
    + collect_submodules("streamlit.components")
    + [
        "streamlit.web.cli",
        "streamlit.__main__",
        "streamlit.runtime.scriptrunner.magic_funcs",
        "streamlit.web.server.websocket_headers",
        "streamlit.components.v1.components",
        "importlib.metadata", "importlib_metadata",
        "altair", "pyarrow",
        "pandas", "pandas.io.formats.style",
        "requests", "urllib3", "charset_normalizer",
        "tkinter", "tkinter.messagebox", "tkinter.ttk",
        "click", "click.core",
        "tornado", "tornado.web", "tornado.websocket",
        "tornado.ioloop", "tornado.httpserver",
        "validators", "packaging", "blinker", "watchdog",
        "gitpython", "pydeck",
    ]
)

a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=numpy_binaries,
    datas=[
        # ── Only two files needed now (config + rollback merged into the PS1) ──
        ("patcher_app.py",         "."),
        ("Invoke-ModernPatch.ps1", "."),
    ] + streamlit_datas + extra_datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=["rthook_numpy.py"],
    excludes=[
        "matplotlib", "scipy", "PIL", "cv2",
        "IPython", "notebook", "jupyter",
        "pywebview", "win32com",
        "numpy.typing",
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="OpsCenter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False, upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=False, upx_exclude=[],
    name="OpsCenter",
)
