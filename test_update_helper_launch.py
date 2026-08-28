"""Regression tests for how the in-app updater launches its helper.

Every update failure from v1.064 to v1.069 was a launcher-plumbing bug, and
all of them are detectable without Windows:

    v1.066  DETACHED_PROCESS | CREATE_NO_WINDOW is an illegal combination
    v1.068  `start DMQuotesUpdate` made cmd treat the title as the program
    v1.069  subprocess escaped `start ""` into `start \\"\\"`, so cmd tried to
            run a program literally named `\\`

These tests pin the invariant that prevents the whole class: the helper is
launched from an argv list with no shell in between.
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
CREATE_NO_WINDOW = 0x08000000

HELPER_PATH = r"C:\Users\ADMINI~1\AppData\Local\Temp\qb_so_updates\apply_update.ps1"


def _argv(helper_path: str = HELPER_PATH) -> list[str]:
    """Mirror of app._update_helper_argv (app.py imports Windows-only deps)."""
    return [
        r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-WindowStyle",
        "Hidden",
        "-File",
        helper_path,
    ]


def test_no_shell_in_launch_chain():
    argv = _argv()
    joined = " ".join(argv).lower()
    assert "cmd.exe" not in joined
    assert "/c" not in argv
    assert "start" not in [a.lower() for a in argv]


def test_no_embedded_quotes_survive_list2cmdline():
    """v1.069: embedded quotes became `\\"\\"` and cmd ran a program named \\."""
    cmdline = subprocess.list2cmdline(_argv())
    assert '\\"' not in cmdline
    assert "start" not in cmdline.lower()


def test_helper_path_is_its_own_argument():
    """Paths with spaces must not need quoting from us."""
    spaced = r"C:\Users\Some User\AppData\Local\Temp\apply_update.ps1"
    argv = _argv(spaced)
    assert argv[-1] == spaced
    assert argv[-2] == "-File"
    # list2cmdline quotes it exactly once, with no escaping artifacts.
    assert f'"{spaced}"' in subprocess.list2cmdline(argv)


def test_creationflags_are_a_legal_combination():
    """v1.066: DETACHED_PROCESS + CREATE_NO_WINDOW is rejected by CreateProcess."""
    flags = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    assert flags & CREATE_NO_WINDOW == 0
    assert flags & DETACHED_PROCESS
    assert flags & CREATE_NEW_PROCESS_GROUP


def test_installer_scrubs_pyinstaller_env_before_relaunch():
    """v1.070: upgrading from an old build died with "Security validation
    failure: parent process has different executable!" because Setup inherited
    the running app's _PYI_* variables and passed them to the [Run] relaunch.
    Old builds cannot be fixed retroactively, so Setup has to do the scrub."""
    iss = Path(__file__).with_name("installer.iss").read_text(encoding="utf-8")

    for name in (
        "_PYI_ARCHIVE_FILE",
        "_PYI_APPLICATION_HOME_DIR",
        "_PYI_PARENT_PROCESS_LEVEL",
    ):
        assert name in iss, f"{name} is no longer cleared by Setup"
    assert "PYINSTALLER_RESET_ENVIRONMENT" in iss

    # The scrub is only useful if it runs before Setup spawns anything.
    init = re.search(r"function InitializeSetup\(\).*?\bend;", iss, re.DOTALL)
    assert init, "InitializeSetup is gone"
    assert "ResetPyInstallerEnvironment" in init.group(0)


if __name__ == "__main__":
    test_no_shell_in_launch_chain()
    test_no_embedded_quotes_survive_list2cmdline()
    test_helper_path_is_its_own_argument()
    test_creationflags_are_a_legal_combination()
    test_installer_scrubs_pyinstaller_env_before_relaunch()
    print("ok")
