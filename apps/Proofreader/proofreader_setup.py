"""
proofreader_setup.py - one-time setup for the Proofreader.

Run by Install_Proofreader.bat. Installs everything the Proofreader needs
except Microsoft Word (which is optional):

  1. the Python add-ons for the window and for Word documents
  2. the free checker, LanguageTool, with its own copy of Java (about 300 MB)
  3. a "Proofreader" shortcut on the Desktop

Running it again is safe: whatever is already there is left alone.
"""

import importlib
import importlib.util
import os
import subprocess
import sys

PACKAGES = {"PySide6": "PySide6", "comtypes": "comtypes", "docx": "python-docx"}
SHORTCUT = "Proofreader.lnk"
LAUNCHER = "Run_Proofreader.bat"


def missing_packages(packages=PACKAGES):
    return [pip for module, pip in packages.items() if importlib.util.find_spec(module) is None]


def install_packages(names, run=subprocess.run):
    """pip-install the named add-ons. Returns an error text, or "" when fine."""
    if not names:
        return ""
    result = run([sys.executable, "-m", "pip", "install", "--disable-pip-version-check"] + names)
    importlib.invalidate_caches()
    return "" if result.returncode == 0 else "pip could not install: " + ", ".join(names)


def shortcut_command(folder):
    """The PowerShell lines that put a shortcut to the launcher on the Desktop."""
    target = os.path.join(folder, LAUNCHER).replace("'", "''")
    start_in = folder.replace("'", "''")
    return ("$desk = [Environment]::GetFolderPath('Desktop'); "
            "$s = (New-Object -ComObject WScript.Shell).CreateShortcut("
            f"(Join-Path $desk '{SHORTCUT}')); "
            f"$s.TargetPath = '{target}'; $s.WorkingDirectory = '{start_in}'; "
            "$s.WindowStyle = 7; $s.Description = 'Proofreader'; $s.Save(); "
            f"Join-Path $desk '{SHORTCUT}'")


def make_shortcut(folder, run=subprocess.run):
    """Returns the shortcut's path, or "" if it could not be made."""
    result = run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                  shortcut_command(folder)], capture_output=True, text=True)
    return result.stdout.strip() if result.returncode == 0 else ""


class Progress:
    """Prints a download's progress every 5 percent instead of every block."""

    def __init__(self, out=print):
        self.out = out
        self.last = None

    def __call__(self, label, part):
        step = (label, int(part * 20))
        if step != self.last:
            self.last = step
            self.out(f"      {label}: {int(part * 100)}%")


def main(folder=None, out=print):
    folder = folder or os.path.dirname(os.path.abspath(__file__))
    problems = []

    out("\n[1 of 4] Python add-ons (PySide6, comtypes, python-docx)")
    needed = missing_packages()
    if needed:
        out("      Installing: " + ", ".join(needed))
        error = install_packages(needed)
        if error:
            problems.append(error)
    else:
        out("      Already installed.")

    out("\n[2 of 4] The free checker (LanguageTool) and its Java")
    sys.path.insert(0, folder)
    import proofreader_tools as tools
    if tools.installed():
        out("      Already installed in " + tools.tools_folder())
    else:
        out(f"      Downloading about {tools.DOWNLOAD_MB} MB. This takes a few minutes.")
        try:
            tools.install(Progress(out))
            out("      Installed in " + tools.tools_folder())
        except Exception as error:              # shown to the person, not hidden
            problems.append(str(error))

    out("\n[3 of 4] Desktop shortcut")
    made = make_shortcut(folder)
    out("      " + (made or "Could not be made. Use Run_Proofreader.bat in this folder instead."))

    out("\n[4 of 4] Microsoft Word (optional)")
    import proofreader_job
    if proofreader_job.word_installed():
        out("      Found. Word's checker will be used after the free checker.")
    else:
        out("      Not found. That is fine: the Proofreader works without it on .docx and\n"
            "      .txt documents. Only .doc and .rtf documents need Word.")

    if problems:
        out("\nSETUP DID NOT FINISH:")
        for problem in problems:
            out("  - " + problem)
        out("\nCheck the internet connection and run Install_Proofreader.bat again.")
        return 1
    out("\nAll set. Start the Proofreader from the Desktop shortcut or Run_Proofreader.bat.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
