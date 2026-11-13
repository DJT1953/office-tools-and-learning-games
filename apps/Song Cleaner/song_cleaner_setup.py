"""
song_cleaner_setup.py - one-time setup for the Song Cleaner.

Run by Install_SongCleaner.bat. Installs everything the Song Cleaner needs:

  1. the Python add-ons: PySide6 (the window), numpy (sound arithmetic),
     mutagen (song tags) and imageio-ffmpeg (which brings FFmpeg, for reading
     and cutting MP3s)
  2. a check that FFmpeg works
  3. if wanted, the small free program that recognises songs (Chromaprint
     fpcalc, about 4 MB), used only by "Look up doubtful names online"
  4. a "Song Cleaner" shortcut on the Desktop

Running it again is safe: whatever is already there is left alone.
"""

import importlib
import importlib.util
import os
import site
import subprocess
import sys

PACKAGES = {"PySide6": "PySide6", "numpy": "numpy>=2", "mutagen": "mutagen",
            "imageio_ffmpeg": "imageio-ffmpeg"}
SHORTCUT = "Song Cleaner.lnk"
LAUNCHER = "Run_SongCleaner.bat"


def missing_packages(packages=PACKAGES):
    """pip names of the add-ons this Python does not have (numpy older than 2 counts as missing)."""
    missing = [pip for module, pip in packages.items() if importlib.util.find_spec(module) is None]
    if "numpy>=2" not in missing and "numpy" in packages:
        try:
            import numpy
            if int(numpy.__version__.split(".")[0]) < 2:
                missing.append(packages["numpy"])
        except (ImportError, ValueError):
            missing.append(packages["numpy"])
    return missing


def install_packages(names, run=subprocess.run):
    """pip-install the named add-ons. Returns an error text, or "" when fine."""
    if not names:
        return ""
    result = run([sys.executable, "-m", "pip", "install", "--upgrade", "--disable-pip-version-check"] + names)
    site.addsitedir(site.getusersitepackages())       # a first install may have just made it
    importlib.invalidate_caches()
    return "" if result.returncode == 0 else "pip could not install: " + ", ".join(names)


def console_python(executable=None):
    """python.exe beside pythonw.exe, so pip's progress can be seen in a window."""
    executable = executable or sys.executable
    folder, name = os.path.split(executable)
    if name.lower() == "pythonw.exe" and os.path.exists(os.path.join(folder, "python.exe")):
        return os.path.join(folder, "python.exe")
    return executable


def offer_missing(ask=None, tell=None, run=subprocess.run):
    """When the app starts without its add-ons: offer to install them, or explain and stop.

    Uses Python's own small window kit (tkinter), because the app's window kit may be
    one of the missing add-ons.
    """
    needed = missing_packages()
    if not needed:
        return True
    if ask is None or tell is None:
        import tkinter
        import tkinter.messagebox as boxes
        root = tkinter.Tk()
        root.withdraw()
        ask = ask or (lambda text: boxes.askyesno("Song Cleaner", text, parent=root))
        tell = tell or (lambda text: boxes.showerror("Song Cleaner", text, parent=root))
    if not ask("The Song Cleaner needs some free add-ons that this copy of Python "
               f"({sys.version.split()[0]}) does not have yet:\n\n    " + ", ".join(needed) +
               "\n\nDownload and install them now? It is about 300 MB to download (850 MB once installed) and takes a few minutes; "
               "a window shows the progress. (Install_SongCleaner.bat does the same.)"):
        sys.exit(0)
    command = [console_python(), "-m", "pip", "install", "--upgrade", "--disable-pip-version-check"] + needed
    flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
    if run(command, creationflags=flags).returncode != 0:
        tell("The add-ons could not be installed. Check the internet connection, then run "
             "Install_SongCleaner.bat in the Song Cleaner folder.")
        sys.exit(1)
    site.addsitedir(site.getusersitepackages())
    importlib.invalidate_caches()
    return True


def ffmpeg_works():
    """The path of the FFmpeg that imageio-ffmpeg brought, or "" if it cannot be run."""
    try:
        import imageio_ffmpeg
        path = imageio_ffmpeg.get_ffmpeg_exe()
        done = subprocess.run([path, "-version"], capture_output=True,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return path if done.returncode == 0 else ""
    except Exception:  # noqa: BLE001 - any trouble means "not working", said to the person
        return ""


def shortcut_command(folder):
    """The PowerShell lines that put a shortcut to the launcher on the Desktop."""
    target = os.path.join(folder, LAUNCHER).replace("'", "''")
    start_in = folder.replace("'", "''")
    return ("$desk = [Environment]::GetFolderPath('Desktop'); "
            "$s = (New-Object -ComObject WScript.Shell).CreateShortcut("
            f"(Join-Path $desk '{SHORTCUT}')); "
            f"$s.TargetPath = '{target}'; $s.WorkingDirectory = '{start_in}'; "
            "$s.WindowStyle = 7; $s.Description = 'Song Cleaner'; $s.Save(); "
            f"Join-Path $desk '{SHORTCUT}'")


def make_shortcut(folder, run=subprocess.run):
    """Returns the shortcut's path, or "" if it could not be made."""
    result = run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                  shortcut_command(folder)], capture_output=True, text=True)
    return result.stdout.strip() if result.returncode == 0 else ""


def ask_yes(question, read=input):
    """Yes unless the person types N. No one at the keyboard counts as no."""
    try:
        return not read(question + " [Y/N] ").strip().lower().startswith("n")
    except (EOFError, OSError):
        return False


def main(folder=None, out=print, ask=ask_yes, run=subprocess.run):
    folder = folder or os.path.dirname(os.path.abspath(__file__))
    problems = []

    out("\n[1 of 4] Python add-ons (PySide6, numpy, mutagen, imageio-ffmpeg)")
    needed = missing_packages()
    if needed:
        out("      Installing: " + ", ".join(needed) + "  (about 300 MB to download; this takes a few minutes)")
        error = install_packages(needed, run)
        if error:
            problems.append(error)
        else:
            out("      Installed.")
    else:
        out("      Already installed.")

    out("\n[2 of 4] FFmpeg, for reading and cutting MP3s")
    found = ffmpeg_works() if not problems else ""
    if found:
        out("      Working: " + found)
    elif not problems:
        problems.append("FFmpeg (from imageio-ffmpeg) could not be started.")

    out("\n[3 of 4] The song recogniser, for looking up names online (optional, about 4 MB)")
    sys.path.insert(0, folder)
    try:
        import song_cleaner_online as online
        if online.fpcalc_path():
            out("      Already installed.")
        elif ask("      It is only needed with a free AcoustID key (see Settings in the app).\n"
                 "      Download it now from the Chromaprint project's page on GitHub?"):
            out("      Downloaded to " + online.fetch_fpcalc())
        else:
            out("      Skipped. The app offers it again the first time you look names up.")
    except Exception as error:  # noqa: BLE001 - optional: reported, not fatal
        out("      Not installed (" + str(error) + "). The app can fetch it later.")

    out("\n[4 of 4] Desktop shortcut")
    made = make_shortcut(folder, run)
    out("      " + (made or "Could not be made. Use Run_SongCleaner.bat in this folder instead."))

    if problems:
        out("\nSETUP DID NOT FINISH:")
        for problem in problems:
            out("  - " + problem)
        out("\nCheck the internet connection and run Install_SongCleaner.bat again.")
        return 1
    out("\nAll set. Start the Song Cleaner from the Desktop shortcut or Run_SongCleaner.bat.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
