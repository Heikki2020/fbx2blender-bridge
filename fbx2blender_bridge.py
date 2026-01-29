import sys
import os
import subprocess
import json
import ctypes
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

CONFIG_PATH = Path.home() / ".fbx2blender_bridge_config.json"


def show_error(message, title="FBX 2 Blender Bridge"):
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x10 | 0x0)


def show_info(message, title="FBX 2 Blender Bridge"):
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x40 | 0x0)


def load_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text())
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def save_config(config):
    CONFIG_PATH.write_text(json.dumps(config, indent=2))


def prompt_blender_path():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Select Blender Executable",
        filetypes=[("Blender", "blender.exe")],
        initialdir=str(
            Path(os.environ.get("ProgramFiles", "C:\\Program Files"))
            / "Blender Foundation"
        ),
    )
    root.destroy()
    return path if path else None


def find_blender():
    config = load_config()
    if "blender_path" in config:
        path = Path(config["blender_path"])
        if path.exists():
            return str(path)
    env_path = os.environ.get("BLENDER_EXECUTABLE")
    if env_path and Path(env_path).exists():
        save_config({"blender_path": env_path})
        return env_path
    try:
        import winreg

        try:
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, r"blender\shell\open\command"
            ) as key:
                cmd = winreg.QueryValue(key, None)
                if "blender.exe" in cmd.lower():
                    path = cmd.split('"')[1] if '"' in cmd else cmd.split()[0]
                    if Path(path).exists():
                        return path
        except OSError:
            pass
    except ImportError:
        pass
    pf = os.environ.get("ProgramFiles", "C:\\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
    for base in [pf, pf86]:
        for ver in [
            "5.0",
            "4.5",
            "4.4",
            "4.3",
            "4.2",
            "4.1",
            "4.0",
            "3.6",
            "3.5",
            "3.4",
            "3.3",
            "3.2",
            "3.1",
            "3.0",
            "",
        ]:
            path = (
                Path(base)
                / "Blender Foundation"
                / (f"Blender {ver}" if ver else "Blender")
                / "blender.exe"
            )
            if path.exists():
                return str(path)
    return None


def launch_blender_with_fbx(fbx_path):
    blender_path = find_blender()
    if not blender_path:
        show_error(
            "Blender not found!\n\n"
            "Please locate blender.exe:\n"
            "• Usually in C:\\Program Files\\Blender Foundation\\Blender X.X\\blender.exe\n"
            "• Or right-click Blender shortcut → Properties → 'Target'",
            "Setup Required",
        )
        new_path = prompt_blender_path()
        if new_path and Path(new_path).exists():
            save_config({"blender_path": new_path})
            blender_path = new_path
        else:
            show_error(
                "Blender path not set.\n\n"
                "Please install Blender from https://blender.org\n"
                "or set BLENDER_EXECUTABLE environment variable.",
                "Setup Failed",
            )
            sys.exit(1)
    script = f"""
import bpy
import sys
fbx_path = {json.dumps(fbx_path)}
try:
    bpy.ops.import_scene.fbx(filepath=fbx_path)
except Exception as e:
    print(f"Import failed: {{e}}", file=sys.stderr)
    bpy.ops.wm.quit_blender()
    sys.exit(1)
for area in bpy.context.screen.areas:
    if area.type == 'VIEW_3D':
        override = {{'area': area, 'region': area.regions[0]}}
        with bpy.context.temp_override(**override):
            bpy.ops.view3d.view_all(center=True)
        break
"""
    import tempfile

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".py", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(script)
        script_path = tmp.name
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.Popen(
            [blender_path, "--python", script_path],
            startupinfo=startupinfo,
            close_fds=True,
        )
    except (subprocess.SubprocessError, OSError) as e:
        show_error(
            f"Failed to launch Blender:\n{e}\n\nPath: {blender_path}", "Launch Error"
        )
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        show_error(
            "No FBX file specified. Please double-click an .fbx file.", "Usage Error"
        )
        sys.exit(1)
    fbx_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(fbx_path):
        show_error(f"File not found:\n{fbx_path}", "File Error")
        sys.exit(1)
    if not fbx_path.lower().endswith(".fbx"):
        show_error(f"Not an FBX file:\n{fbx_path}", "File Type Error")
        sys.exit(1)
    launch_blender_with_fbx(fbx_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        show_error(f"Unexpected error:\n{e}", "Critical Error")
        sys.exit(1)
