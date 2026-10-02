#!/usr/bin/env python3
"""
Windows Bloat Remover v1.2 (by George Tsakalos, 10/2026)
--------------------------------
Windows 10/11 Tkinter utility for inspecting & removing AppX/MSIX bloat / garbage packages!

Features
- Reads installed packages using PowerShell Get-AppxPackage.
- Shows Name, PackageFullName, description, risk, installation/provisioning status.
- Checkbox-style selection in a sortable/filterable GUI.
- Loads default-checked package names from:
      default_checked_packages.txt
  beside this script.
- Default packages are shown even when not found in the initial Get-AppxPackage query.
- "Remove Now" removes selected packages:
    1) from existing users with Remove-AppxPackage -AllUsers
    2) from the online Windows image with Remove-AppxProvisionedPackage -Online -AllUsers
       when a provisioned package exists.
- "REINSTALL / RESTORE" attempts to restore selected packages by:
    1) re-registering an existing/staged AppxManifest.xml when package files still exist
    2) installing known Microsoft Store apps through WinGet when a Store ID is known.
- Automatically requests Administrator elevation.
- Writes a detailed log beside the script.

IMPORTANT
This program intentionally allows advanced users to attempt removal of Red/Dangerous
packages after an explicit confirmation. Windows may refuse removal of protected
NonRemovable packages.
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import subprocess
import sys
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox


APP_TITLE = "Windows AppX Package Remover v1.2 (by George Tsakalos)"
DEFAULTS_FILENAME = "default_checked_packages.txt"
LOG_FILENAME = "AppxPackageRemover.log"

RISK_GREEN = "Green"
RISK_YELLOW = "Yellow"
RISK_RED = "Red"

# Optional built-in defaults.
# The preferred method is default_checked_packages.txt beside the .py file.
# One package Name per line is enough; lines copied from Get-AppxPackage output
# are also accepted.
BUILTIN_DEFAULT_CHECKED: set[str] = set()




# Package knowledge base
# ----------------------

PACKAGE_INFO: dict[str, tuple[str, str]] = {
    "Microsoft.BioEnrollment": (
        "Windows Hello biometric-enrollment interface (fingerprint/face/PIN setup support).",
        RISK_RED,
    ),
    "Microsoft.Windows.CloudExperienceHost": (
        "Core Windows first-run, account, network and cloud setup experience.",
        RISK_RED,
    ),
    "Microsoft.Windows.OOBENetworkConnectionFlow": (
        "Windows Out-of-Box Experience network connection workflow.",
        RISK_RED,
    ),
    "Microsoft.AAD.BrokerPlugin": (
        "Microsoft account / Entra ID (Azure AD) authentication broker used by Windows and apps.",
        RISK_RED,
    ),
    "Microsoft.Windows.OOBENetworkCaptivePortal": (
        "Out-of-Box Experience support for captive-portal Wi-Fi/network sign-in.",
        RISK_RED,
    ),
    "MicrosoftWindows.Client.CBS": (
        "Core Windows shell/servicing component used by modern Windows experiences.",
        RISK_RED,
    ),
    "MicrosoftWindows.UndockedDevKit": (
        "Windows shell component supporting modern/undocked shell experiences.",
        RISK_RED,
    ),
    "Microsoft.Windows.StartMenuExperienceHost": (
        "Windows Start menu host.",
        RISK_RED,
    ),
    "Microsoft.Windows.ShellExperienceHost": (
        "Core Windows shell visual experience host.",
        RISK_RED,
    ),
    "windows.immersivecontrolpanel": (
        "Windows Settings application.",
        RISK_RED,
    ),
    "Microsoft.Windows.ContentDeliveryManager": (
        "Windows suggested content, Spotlight/promotional content and related content delivery.",
        RISK_YELLOW,
    ),
    "Microsoft.DesktopAppInstaller": (
        "Microsoft App Installer; handles MSIX/AppX installation and is associated with WinGet functionality.",
        RISK_YELLOW,
    ),
    "Microsoft.WindowsCamera": (
        "Built-in Windows Camera application.",
        RISK_GREEN,
    ),
    "Microsoft.Windows.XGpuEjectDialog": (
        "Windows external/discrete GPU safe-eject user interface.",
        RISK_YELLOW,
    ),
    "Microsoft.XboxGameCallableUI": (
        "Xbox gaming sign-in/callable user interface used by some games.",
        RISK_GREEN,
    ),
    "NcsiUwpApp": (
        "Windows Network Connectivity Status Indicator interface.",
        RISK_RED,
    ),
    "Windows.CBSPreview": (
        "Windows shell / Component-Based Servicing related system package.",
        RISK_RED,
    ),
    "Microsoft.AsyncTextService": (
        "Windows text input, typing and text-services component.",
        RISK_RED,
    ),
    "Microsoft.CredDialogHost": (
        "Windows credential/password/PIN dialog host.",
        RISK_RED,
    ),
    "Microsoft.ECApp": (
        "Windows system experience component used by built-in shell functionality.",
        RISK_YELLOW,
    ),
    "Microsoft.AccountsControl": (
        "Windows account-management user interface.",
        RISK_RED,
    ),
    "Microsoft.LockApp": (
        "Windows lock-screen application.",
        RISK_YELLOW,
    ),
    "Microsoft.MicrosoftEdgeDevToolsClient": (
        "System-hosted Microsoft Edge/WebView developer-tools client.",
        RISK_YELLOW,
    ),
    "Microsoft.Win32WebViewHost": (
        "Web-content host used by Win32 applications and Windows integration.",
        RISK_RED,
    ),
    "Microsoft.Windows.Apprep.ChxApp": (
        "Windows application-preparation / compatibility system component.",
        RISK_RED,
    ),
    "Windows.PrintDialog": (
        "Modern Windows print dialog.",
        RISK_RED,
    ),
    "Microsoft.Windows.SecureAssessmentBrowser": (
        "Windows secure assessment/exam browser used in education/test environments.",
        RISK_GREEN,
    ),
    "Microsoft.Windows.SecHealthUI": (
        "Windows Security user interface for Defender, firewall and security status.",
        RISK_RED,
    ),
    "Microsoft.Windows.PinningConfirmationDialog": (
        "Windows shell confirmation dialog for application pinning operations.",
        RISK_YELLOW,
    ),
    "Microsoft.Windows.PeopleExperienceHost": (
        "Windows People/contact integration host used by shell experiences.",
        RISK_YELLOW,
    ),
    "Microsoft.Windows.ParentalControls": (
        "Microsoft Family / Windows parental-controls component.",
        RISK_GREEN,
    ),
    "Microsoft.Windows.NarratorQuickStart": (
        "Narrator accessibility quick-start/tutorial application.",
        RISK_GREEN,
    ),
    "Microsoft.Windows.CapturePicker": (
        "Windows screen/window capture picker used by applications.",
        RISK_YELLOW,
    ),
    "Microsoft.Windows.CallingShellApp": (
        "Windows calling/phone integration shell component.",
        RISK_GREEN,
    ),
    "Microsoft.Windows.AssignedAccessLockApp": (
        "Assigned Access / kiosk-mode lock application.",
        RISK_GREEN,
    ),
    "Microsoft.StorePurchaseApp": (
        "Microsoft Store purchase/licensing support component.",
        RISK_YELLOW,
    ),
    "Microsoft.Windows.Search": (
        "Windows Search shell application and search interface.",
        RISK_RED,
    ),
    "Microsoft.VCLibs.140.00": (
        "Microsoft Visual C++ UWP runtime library required by many Store/UWP applications.",
        RISK_RED,
    ),
    "Microsoft.NET.Native.Framework.2.2": (
        ".NET Native framework runtime dependency for UWP applications.",
        RISK_RED,
    ),
    "Microsoft.NET.Native.Runtime.2.2": (
        ".NET Native runtime dependency for UWP applications.",
        RISK_RED,
    ),
    "Microsoft.NET.Native.Framework.1.7": (
        "Older .NET Native framework runtime dependency for UWP applications.",
        RISK_RED,
    ),
    "Microsoft.NET.Native.Runtime.1.7": (
        "Older .NET Native runtime dependency for UWP applications.",
        RISK_RED,
    ),
    "Microsoft.VCLibs.140.00.UWPDesktop": (
        "Visual C++ runtime bridge used by packaged desktop/UWP applications.",
        RISK_RED,
    ),
    "Microsoft.UI.Xaml.2.8": (
        "Microsoft WinUI/XAML runtime framework used by modern applications.",
        RISK_RED,
    ),
    "Microsoft.UI.Xaml.2.7": (
        "Microsoft WinUI/XAML runtime framework used by modern applications.",
        RISK_RED,
    ),
    "Microsoft.WindowsAppRuntime.1.8": (
        "Windows App SDK runtime used by modern desktop applications.",
        RISK_RED,
    ),
    "Microsoft.Windows.DevHome": (
        "Microsoft Dev Home dashboard for developer environments, repositories and development tools.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsCalculator": (
        "Built-in Windows Calculator application.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsAlarms": (
        "Built-in Windows Clock/Alarms/Timer application.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsStore": (
        "Microsoft Store application and Store management interface.",
        RISK_YELLOW,
    ),
    "AppUp.IntelGraphicsExperience": (
        "Intel Graphics software/control experience supplied through Microsoft Store packaging.",
        RISK_YELLOW,
    ),
    "AppUp.IntelOptaneMemoryandStorageManagement": (
        "Intel Optane Memory and Storage Management / Intel RST management interface.",
        RISK_YELLOW,
    ),
    "WavesAudio.MaxxAudioProforDell2020": (
        "Waves MaxxAudio Pro audio-control/effects application commonly supplied with Dell systems.",
        RISK_YELLOW,
    ),
}

# Additional friendly descriptions for common Windows 10 AppX packages.
# These names/purposes were cross-checked against Microsoft documentation,
# Microsoft Store listings and Microsoft's Windows optimization package lists.
COMMON_PACKAGE_INFO: dict[str, tuple[str, str]] = {
    "Microsoft.MixedReality.Portal": (
        "Mixed Reality Portal; sets up and manages Windows Mixed Reality headsets and experiences.",
        RISK_GREEN,
    ),
    "Microsoft.3DBuilder": (
        "3D Builder; Microsoft app for viewing, creating, editing and printing 3D models.",
        RISK_GREEN,
    ),
    "Microsoft.Microsoft3DViewer": (
        "3D Viewer; opens and displays common 3D model formats.",
        RISK_GREEN,
    ),
    "Microsoft.AV1VideoExtension": (
        "AV1 Video Extension; codec support that lets Windows apps decode AV1 video.",
        RISK_GREEN,
    ),
    "microsoft.windowscommunicationsapps": (
        "Windows Mail and Calendar applications.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsFeedbackHub": (
        "Feedback Hub; sends Windows problem reports and feature suggestions to Microsoft.",
        RISK_GREEN,
    ),
    "Microsoft.MicrosoftOfficeHub": (
        "Microsoft Office / Microsoft 365 hub app and Office shortcuts.",
        RISK_GREEN,
    ),
    "Microsoft.Getstarted": (
        "Microsoft Tips / Get Started app with Windows feature tips and guidance.",
        RISK_GREEN,
    ),
    "Microsoft.ZuneMusic": (
        "Groove Music / Media Player music application.",
        RISK_GREEN,
    ),
    "Microsoft.HEIFImageExtension": (
        "HEIF Image Extensions; adds HEIF/HEIC image decoding support to Windows applications.",
        RISK_GREEN,
    ),
    "Microsoft.GetHelp": (
        "Get Help; Microsoft's Windows support and troubleshooting application.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsMaps": (
        "Windows Maps; Microsoft's built-in mapping application.",
        RISK_GREEN,
    ),
    "Microsoft.MicrosoftSolitaireCollection": (
        "Microsoft Solitaire Collection game package.",
        RISK_GREEN,
    ),
    "Microsoft.Todos": (
        "Microsoft To Do task-list and reminder application.",
        RISK_GREEN,
    ),
    "MicrosoftTeams": (
        "Microsoft Teams communication and collaboration application.",
        RISK_GREEN,
    ),
    "MSTeams": (
        "Microsoft Teams communication and collaboration application.",
        RISK_GREEN,
    ),
    "Microsoft.BingFinance": (
        "MSN Money / Finance app for markets, finance news and portfolio information.",
        RISK_GREEN,
    ),
    "Microsoft.ZuneVideo": (
        "Movies & TV application for playing and managing video content.",
        RISK_GREEN,
    ),
    "Microsoft.BingNews": (
        "Microsoft News / MSN News application.",
        RISK_GREEN,
    ),
    "Microsoft.Office.OneNote": (
        "Microsoft OneNote note-taking application.",
        RISK_GREEN,
    ),
    "Microsoft.OneDriveSync": (
        "Microsoft OneDrive synchronization application/integration.",
        RISK_YELLOW,
    ),
    "Microsoft.MSPaint": (
        "Microsoft Paint / Paint 3D packaged application, depending on Windows build.",
        RISK_GREEN,
    ),
    "Microsoft.People": (
        "Microsoft People / Contacts application.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsPhone": (
        "Legacy Windows phone companion/integration application.",
        RISK_GREEN,
    ),
    "Microsoft.PowerAutomateDesktop": (
        "Power Automate for desktop; Microsoft's desktop workflow/RPA automation application.",
        RISK_GREEN,
    ),
    "Microsoft.ScreenSketch": (
        "Snip & Sketch / Snipping Tool screen-capture and annotation application.",
        RISK_GREEN,
    ),
    "Microsoft.SkypeApp": (
        "Microsoft Skype communication application.",
        RISK_GREEN,
    ),
    "Microsoft.MicrosoftStickyNotes": (
        "Microsoft Sticky Notes; quick desktop notes application.",
        RISK_GREEN,
    ),
    "SpotifyAB.SpotifyMusic": (
        "Spotify desktop music-streaming application packaged from Microsoft Store.",
        RISK_GREEN,
    ),
    "Microsoft.BingSports": (
        "MSN Sports application.",
        RISK_GREEN,
    ),
    "Microsoft.WindowsSoundRecorder": (
        "Windows Voice Recorder / Sound Recorder application.",
        RISK_GREEN,
    ),
    "Microsoft.BingWeather": (
        "MSN Weather application.",
        RISK_GREEN,
    ),
    "Microsoft.WebpImageExtension": (
        "WebP Image Extension; adds WebP image decoding support to Windows applications.",
        RISK_GREEN,
    ),
    "Microsoft.WebMediaExtensions": (
        "Web Media Extensions; adds support for open web media formats such as OGG/Vorbis/Theora.",
        RISK_GREEN,
    ),
    "Microsoft.XboxApp": (
        "Xbox application / Xbox Console Companion functionality.",
        RISK_GREEN,
    ),
    "Microsoft.XboxGamingOverlay": (
        "Xbox Game Bar overlay for gaming capture, widgets and game-related controls.",
        RISK_GREEN,
    ),
    "Microsoft.XboxIdentityProvider": (
        "Xbox identity/sign-in provider used by Xbox-enabled games and services.",
        RISK_YELLOW,
    ),
    "Microsoft.XboxSpeechToTextOverlay": (
        "Xbox speech-to-text overlay/accessibility component.",
        RISK_GREEN,
    ),
    "Microsoft.YourPhone": (
        "Phone Link (formerly Your Phone); links Android/iPhone features with Windows.",
        RISK_GREEN,
    ),
    "Microsoft.Wallet": (
        "Legacy Microsoft Wallet application/component.",
        RISK_GREEN,
    ),
}

# Pattern descriptions used by wildcard entries in default_checked_packages.txt.
PATTERN_PACKAGE_INFO: list[tuple[str, str, str]] = [
    ("*mixedreality*", "Windows Mixed Reality software / Mixed Reality Portal.", RISK_GREEN),
    ("*3dbuilder*", "3D Builder; Microsoft tool for creating/editing 3D models.", RISK_GREEN),
    ("*3dviewer*", "3D Viewer; displays common 3D model formats.", RISK_GREEN),
    ("*alarms*", "Windows Clock / Alarms / Timer / Stopwatch application.", RISK_GREEN),
    ("*av1videoextension*", "AV1 codec extension for Windows video playback.", RISK_GREEN),
    ("*communications*", "Windows Mail and Calendar application package.", RISK_GREEN),
    ("*windowsfeedbackhub*", "Feedback Hub for Windows feedback and problem reports.", RISK_GREEN),
    ("*officehub*", "Microsoft Office / Microsoft 365 hub application.", RISK_GREEN),
    ("*getstarted*", "Microsoft Tips / Get Started Windows guidance application.", RISK_GREEN),
    ("*zunemusic*", "Groove Music / Media Player music application.", RISK_GREEN),
    ("*heifimageextension*", "HEIF/HEIC image codec extension for Windows.", RISK_GREEN),
    ("*gethelp*", "Microsoft Get Help support/troubleshooting application.", RISK_GREEN),
    ("*maps*", "Windows Maps application.", RISK_GREEN),
    ("*solitairecollection*", "Microsoft Solitaire Collection.", RISK_GREEN),
    ("*todos*", "Microsoft To Do task manager.", RISK_GREEN),
    ("*teams*", "Microsoft Teams communication/collaboration application.", RISK_GREEN),
    ("*bingfinance*", "MSN Money / Finance application.", RISK_GREEN),
    ("*zunevideo*", "Movies & TV video application.", RISK_GREEN),
    ("*bingnews*", "Microsoft News / MSN News application.", RISK_GREEN),
    ("*onenote*", "Microsoft OneNote note-taking application.", RISK_GREEN),
    ("*onedrivesync*", "Microsoft OneDrive synchronization application.", RISK_YELLOW),
    ("*paint*", "Microsoft Paint / Paint 3D package.", RISK_GREEN),
    ("*people*", "Microsoft People / Contacts application.", RISK_GREEN),
    ("*windowsphone*", "Legacy Windows phone companion/integration application.", RISK_GREEN),
    ("*powerautomatedesktop*", "Microsoft Power Automate for desktop automation.", RISK_GREEN),
    ("*screensketch*", "Snip & Sketch / Snipping Tool screen capture application.", RISK_GREEN),
    ("*skype*", "Microsoft Skype communication application.", RISK_GREEN),
    ("*stickynotes*", "Microsoft Sticky Notes application.", RISK_GREEN),
    ("*spotifymusic*", "Spotify music-streaming application.", RISK_GREEN),
    ("*bingsports*", "MSN Sports application.", RISK_GREEN),
    ("*soundrecorder*", "Windows Voice Recorder / Sound Recorder.", RISK_GREEN),
    ("*bingweather*", "MSN Weather application.", RISK_GREEN),
    ("*webpimageextension*", "WebP image codec extension.", RISK_GREEN),
    ("*webmediaextension*", "Windows web-media codec extensions.", RISK_GREEN),
    ("*xbox*", "Xbox gaming application/component. Exact function depends on matched package.", RISK_GREEN),
    ("*yourphone*", "Phone Link (formerly Your Phone).", RISK_GREEN),
    ("*wallet*", "Legacy Microsoft Wallet application/component.", RISK_GREEN),
    ("*onedrive*", "Microsoft OneDrive synchronization/integration package.", RISK_YELLOW),
    ("*bioenrollment*", "Windows Hello biometric-enrollment interface.", RISK_RED),
    ("*brokerplugin*", "Microsoft account / Entra ID authentication broker.", RISK_RED),
    ("*peopleexperience*", "Windows People/contact shell integration host.", RISK_YELLOW),
    ("*advertising*", "Microsoft advertising / advertising-ID related packaged component.", RISK_YELLOW),
]

# Microsoft Store IDs used only as a fallback when no local AppxManifest.xml
# remains to re-register. Not every Windows system package has a Store listing.
STORE_IDS: dict[str, str] = {
    "Microsoft.DesktopAppInstaller": "9NBLGGH4NNS1",
    "Microsoft.WindowsCalculator": "9WZDNCRFHVN5",
    "Microsoft.WindowsAlarms": "9WZDNCRFJ3PR",
    "Microsoft.WindowsCamera": "9WZDNCRFJBBG",
    "Microsoft.WindowsFeedbackHub": "9NBLGGH4R32N",
    "Microsoft.WindowsMaps": "9WZDNCRDTBVB",
    "Microsoft.ZuneMusic": "9WZDNCRFJ3PT",
    "Microsoft.ZuneVideo": "9WZDNCRFJ3P2",
    "Microsoft.WindowsSoundRecorder": "9WZDNCRFHWKN",
    "Microsoft.ScreenSketch": "9MZ95KL8MR0L",
    "Microsoft.MixedReality.Portal": "9NG1H8B3ZC7M",
    "Microsoft.Microsoft3DViewer": "9NBLGGH42THS",
}



# Known opaque Windows package identities seen on Windows.
# Descriptions are deliberately conservative because Microsoft does not expose
# friendly names for all of these internal package IDs.
OPAQUE_SYSTEM_PACKAGES = {
    "1527c705-839a-4832-9118-54d4Bd6a0c89",
    "c5e2524a-ea46-4f67-841f-6a9465d9d515",
    "E2A4F912-2574-4A75-9BB0-0D023378592B",
    "F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE",
}


@dataclass
class PackageRow:
    name: str
    package_full_name: str = ""
    installed_current_user: bool = False
    installed_any_user: bool = False
    provisioned: bool = False
    provisioned_package_name: str = ""
    non_removable: bool = False
    selected: bool = False

    @property
    def status(self) -> str:
        bits = []
        if self.installed_current_user:
            bits.append("Installed")
        elif self.installed_any_user:
            bits.append("Other user")
        if self.provisioned:
            bits.append("Provisioned")
        if not bits:
            return "Not found"
        return " + ".join(bits)


def script_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


LOG_PATH = script_dir() / LOG_FILENAME
DEFAULTS_PATH = script_dir() / DEFAULTS_FILENAME


def log(message: str) -> None:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"[{stamp}] {message}\n")
    except Exception:
        pass


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    """Relaunch this script elevated. Returns True if launch request was made."""
    if is_admin():
        return False

    if getattr(sys, "frozen", False):
        executable = sys.executable
        params = " ".join(f'"{a}"' for a in sys.argv[1:])
    else:
        executable = sys.executable
        params = " ".join([f'"{Path(__file__).resolve()}"'] + [f'"{a}"' for a in sys.argv[1:]])

    try:
        rc = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", executable, params, str(script_dir()), 1
        )
        if rc <= 32:
            raise OSError(f"ShellExecuteW returned {rc}")
        return True
    except Exception as exc:
        messagebox.showerror(
            APP_TITLE,
            "Administrator rights are required for -AllUsers removal.\n\n"
            f"Could not request elevation:\n{exc}",
        )
        return False


def run_powershell(script: str, timeout: int = 120) -> tuple[int, str, str]:
    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]
    creationflags = 0
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    p = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        creationflags=creationflags,
    )
    return p.returncode, p.stdout.strip(), p.stderr.strip()


def ps_json(script: str):
    wrapped = (
        "$ProgressPreference='SilentlyContinue'; "
        "$ErrorActionPreference='Stop'; "
        f"{script} | ConvertTo-Json -Depth 5 -Compress"
    )
    rc, out, err = run_powershell(wrapped)
    if rc != 0:
        raise RuntimeError(err or out or "PowerShell command failed.")
    if not out:
        return []
    data = json.loads(out)
    return data if isinstance(data, list) else [data]


def parse_default_line(line: str) -> str | None:
    line = line.strip().lstrip("\ufeff")
    if not line or line.startswith("#"):
        return None
    if line.lower().startswith("name ") or line.startswith("----"):
        return None

    # Accept legacy PowerShell removal commands such as:
    # Get-AppxPackage *MixedReality* -allusers | Remove-AppxPackage
    # Store the wildcard expression itself; Get-AppxPackage -Name accepts it.
    m = re.search(r"Get-AppxPackage\s+([^|\s]+)", line, flags=re.IGNORECASE)
    if m:
        candidate = m.group(1).strip().strip('"').strip("'")
        if candidate:
            return candidate

    # Accept a bare package name, PackageFullName, or a copied two-column line.
    token = re.split(r"\s{2,}|\t+", line, maxsplit=1)[0].strip()
    if not token:
        return None

    # If a full package identity was supplied, reduce it to its package Name.
    # Example:
    # Microsoft.WindowsCalculator_11.2607.0.0_x64__8wekyb3d8bbwe
    m = re.match(r"^(.+?)_\d+(?:\.\d+){1,}_.+__", token)
    if m:
        return m.group(1)

    return token


def load_default_checked() -> set[str]:
    result = set(BUILTIN_DEFAULT_CHECKED)
    if DEFAULTS_PATH.exists():
        try:
            for line in DEFAULTS_PATH.read_text(encoding="utf-8-sig", errors="replace").splitlines():
                name = parse_default_line(line)
                if name:
                    result.add(name)
        except Exception as exc:
            log(f"Could not read defaults file: {exc}")
    return result


def package_description_and_risk(name: str, non_removable: bool) -> tuple[str, str]:
    if non_removable:
        desc = PACKAGE_INFO.get(name, ("Protected Windows system package.", RISK_RED))[0]
        return desc + " Windows reports this package as NonRemovable.", RISK_RED

    if name in PACKAGE_INFO:
        return PACKAGE_INFO[name]

    if name in COMMON_PACKAGE_INFO:
        return COMMON_PACKAGE_INFO[name]

    # Friendly descriptions for wildcard defaults such as *MixedReality*.
    lname_for_pattern = name.lower()
    if "*" in name or "?" in name:
        import fnmatch
        for pattern, description, risk in PATTERN_PACKAGE_INFO:
            if fnmatch.fnmatch(lname_for_pattern, pattern.lower()):
                return description, risk

    if name in OPAQUE_SYSTEM_PACKAGES or re.fullmatch(
        r"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}",
        name,
    ):
        return (
            "Opaque/internal Windows system AppX package. Friendly purpose is not exposed by its package identity.",
            RISK_RED,
        )

    lname = name.lower()

    # Framework/library heuristics: removing shared dependencies can break apps.
    if any(x in lname for x in ("framework", "runtime", "vclibs", "xaml")):
        return (
            "Shared application framework/runtime dependency. Other packaged applications may depend on it.",
            RISK_RED,
        )

    if lname.startswith(("microsoft.windows.", "microsoftwindows.", "windows.")):
        return (
            "Windows/Microsoft AppX system or inbox component. No verified friendly description was found for this exact identity.",
            RISK_YELLOW,
        )

    if lname.startswith("microsoft."):
        return (
            "Microsoft packaged application/component. No verified friendly description was found for this exact identity.",
            RISK_YELLOW,
        )

    return (
        "Third-party or OEM packaged application. Verify its purpose before removal.",
        RISK_GREEN,
    )


def discover_packages(default_names: set[str]) -> dict[str, PackageRow]:
    rows: dict[str, PackageRow] = {}

    # User-requested initial inventory:
    # Get-AppxPackage | Select Name, PackageFullName
    current = ps_json(
        "Get-AppxPackage | "
        "Select-Object Name,PackageFullName,NonRemovable"
    )
    for item in current:
        name = str(item.get("Name") or "").strip()
        if not name:
            continue
        rows[name] = PackageRow(
            name=name,
            package_full_name=str(item.get("PackageFullName") or ""),
            installed_current_user=True,
            installed_any_user=True,
            non_removable=bool(item.get("NonRemovable", False)),
        )

    # Additional all-user lookup is needed for accurate -AllUsers removal/status.
    try:
        all_users = ps_json(
            "Get-AppxPackage -AllUsers | "
            "Select-Object Name,PackageFullName,NonRemovable"
        )
        for item in all_users:
            name = str(item.get("Name") or "").strip()
            if not name:
                continue
            row = rows.setdefault(name, PackageRow(name=name))
            row.installed_any_user = True
            if not row.package_full_name:
                row.package_full_name = str(item.get("PackageFullName") or "")
            row.non_removable = row.non_removable or bool(item.get("NonRemovable", False))
    except Exception as exc:
        log(f"All-users discovery failed: {exc}")

    # Provisioned inventory: needed if "complete removal" should prevent the
    # package from being installed for future accounts.
    try:
        prov = ps_json(
            "Get-AppxProvisionedPackage -Online | "
            "Select-Object DisplayName,PackageName"
        )
        for item in prov:
            name = str(item.get("DisplayName") or "").strip()
            if not name:
                continue
            row = rows.setdefault(name, PackageRow(name=name))
            row.provisioned = True
            row.provisioned_package_name = str(item.get("PackageName") or "")
    except Exception as exc:
        log(f"Provisioned-package discovery failed: {exc}")

    # Always include user defaults, even if not present in current inventory.
    for name in default_names:
        rows.setdefault(name, PackageRow(name=name)).selected = True

    return rows


class AppxRemoverGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1500x820")
        self.minsize(1100, 650)

        self.default_names = load_default_checked()
        self.rows: dict[str, PackageRow] = {}
        self.item_to_name: dict[str, str] = {}
        self.busy = False

        self.filter_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready.")
        self.summary_var = tk.StringVar()
        self.remove_provisioned_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.after(100, self.refresh_packages)

    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 10, 10, 6))
        top.pack(fill="x")

        ttk.Label(top, text="Filter:").pack(side="left")
        ent = ttk.Entry(top, textvariable=self.filter_var, width=38)
        ent.pack(side="left", padx=(6, 12))
        ent.bind("<KeyRelease>", lambda e: self.populate_tree())

        ttk.Button(top, text="Refresh", command=self.refresh_packages).pack(side="left", padx=3)
        ttk.Button(top, text="Select Green", command=lambda: self.select_by_risk(RISK_GREEN)).pack(side="left", padx=3)
        ttk.Button(top, text="Clear All", command=self.clear_all).pack(side="left", padx=3)
        ttk.Button(top, text="Restore Defaults", command=self.restore_defaults).pack(side="left", padx=3)

        ttk.Checkbutton(
            top,
            text="Also remove provisioned copy (recommended for complete removal)",
            variable=self.remove_provisioned_var,
        ).pack(side="right")

        note = ttk.Label(
            self,
            text=(
                f"Default selections file: {DEFAULTS_PATH}   |   "
                "Green = low system-stability risk, Yellow = medium, Red = dangerous/system dependency"
            ),
            padding=(10, 0, 10, 6),
        )
        note.pack(fill="x")

        frame = ttk.Frame(self, padding=(10, 0, 10, 0))
        frame.pack(fill="both", expand=True)

        columns = ("check", "name", "risk", "status", "description", "full")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")

        headings = {
            "check": "✓",
            "name": "Package Name",
            "risk": "Risk",
            "status": "Status",
            "description": "Description",
            "full": "PackageFullName",
        }
        widths = {
            "check": 42,
            "name": 285,
            "risk": 85,
            "status": 130,
            "description": 545,
            "full": 400,
        }

        for col in columns:
            self.tree.heading(col, text=headings[col], command=lambda c=col: self.sort_tree(c, False))
            self.tree.column(col, width=widths[col], minwidth=35, stretch=(col in ("description", "full")))

        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # Row coloring. Colors are intentionally soft enough to preserve readability.
        self.tree.tag_configure("green", background="#dff5df")
        self.tree.tag_configure("yellow", background="#fff4bf")
        self.tree.tag_configure("red", background="#ffd6d6")
        self.tree.tag_configure("missing", foreground="#777777")

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<space>", self.on_space_toggle)

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")

        ttk.Label(bottom, textvariable=self.summary_var).pack(side="left")
        self.remove_btn = ttk.Button(
            bottom,
            text="REMOVE NOW",
            command=self.remove_selected,
        )
        self.remove_btn.pack(side="right", padx=(10, 0))

        self.reinstall_btn = ttk.Button(
            bottom,
            text="REINSTALL / RESTORE",
            command=self.reinstall_selected,
        )
        self.reinstall_btn.pack(side="right", padx=(10, 0))

        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=180)
        self.progress.pack(side="right", padx=10)

        status = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=(6, 3))
        status.pack(fill="x", side="bottom")

    def set_busy(self, state: bool, text: str | None = None):
        self.busy = state
        self.remove_btn.configure(state="disabled" if state else "normal")
        self.reinstall_btn.configure(state="disabled" if state else "normal")
        if state:
            self.progress.start(10)
        else:
            self.progress.stop()
        if text:
            self.status_var.set(text)
        self.update_idletasks()

    def refresh_packages(self):
        if self.busy:
            return
        self.set_busy(True, "Reading installed AppX packages...")
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self):
        try:
            old_selected = {name for name, row in self.rows.items() if row.selected}
            defaults = load_default_checked()
            rows = discover_packages(defaults)

            # Preserve manual checks through Refresh.
            for name in old_selected:
                rows.setdefault(name, PackageRow(name=name)).selected = True

            self.after(0, lambda: self._refresh_done(rows, defaults, None))
        except Exception as exc:
            detail = f"{exc}\n\n{traceback.format_exc()}"
            log(detail)
            self.after(0, lambda: self._refresh_done({}, self.default_names, detail))

    def _refresh_done(self, rows, defaults, error):
        self.set_busy(False)
        if error:
            self.status_var.set("Package discovery failed.")
            messagebox.showerror(APP_TITLE, "Could not read AppX packages:\n\n" + error)
            return
        self.rows = rows
        self.default_names = defaults
        self.populate_tree()
        self.status_var.set(f"Loaded {len(rows)} package identities.")

    def visible_rows(self):
        q = self.filter_var.get().strip().lower()
        items = list(self.rows.values())
        if q:
            filtered = []
            for row in items:
                desc, risk = package_description_and_risk(row.name, row.non_removable)
                hay = " ".join([
                    row.name,
                    row.package_full_name,
                    desc,
                    risk,
                    row.status,
                ]).lower()
                if q in hay:
                    filtered.append(row)
            items = filtered
        return sorted(items, key=lambda r: r.name.lower())

    def populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.item_to_name.clear()

        for row in self.visible_rows():
            desc, risk = package_description_and_risk(row.name, row.non_removable)
            check = "☑" if row.selected else "☐"
            tag = risk.lower()
            extra_tags = [tag]
            if row.status == "Not found":
                extra_tags.append("missing")

            iid = self.tree.insert(
                "",
                "end",
                values=(
                    check,
                    row.name,
                    risk,
                    row.status,
                    desc,
                    row.package_full_name or row.provisioned_package_name or "",
                ),
                tags=tuple(extra_tags),
            )
            self.item_to_name[iid] = row.name

        self.update_summary()

    def update_summary(self):
        selected = [r for r in self.rows.values() if r.selected]
        counts = {RISK_GREEN: 0, RISK_YELLOW: 0, RISK_RED: 0}
        for row in selected:
            _, risk = package_description_and_risk(row.name, row.non_removable)
            counts[risk] += 1
        self.summary_var.set(
            f"Selected: {len(selected)}   "
            f"Green: {counts[RISK_GREEN]}   "
            f"Yellow: {counts[RISK_YELLOW]}   "
            f"Red: {counts[RISK_RED]}"
        )

    def toggle_item(self, iid):
        if not iid or iid not in self.item_to_name:
            return
        name = self.item_to_name[iid]
        row = self.rows[name]
        row.selected = not row.selected
        vals = list(self.tree.item(iid, "values"))
        vals[0] = "☑" if row.selected else "☐"
        self.tree.item(iid, values=vals)
        self.update_summary()

    def on_tree_click(self, event):
        if self.busy:
            return
        region = self.tree.identify("region", event.x, event.y)
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if region == "cell" and col == "#1":
            self.toggle_item(iid)
            return "break"

    def on_space_toggle(self, event):
        iid = self.tree.focus()
        self.toggle_item(iid)
        return "break"

    def clear_all(self):
        for row in self.rows.values():
            row.selected = False
        self.populate_tree()

    def restore_defaults(self):
        defaults = load_default_checked()
        self.default_names = defaults
        for row in self.rows.values():
            row.selected = row.name in defaults
        for name in defaults:
            self.rows.setdefault(name, PackageRow(name=name, selected=True))
        self.populate_tree()

    def select_by_risk(self, wanted_risk):
        for row in self.rows.values():
            _, risk = package_description_and_risk(row.name, row.non_removable)
            if risk == wanted_risk and row.status != "Not found":
                row.selected = True
        self.populate_tree()

    def sort_tree(self, col, reverse):
        # Sort only currently displayed items.
        items = [(self.tree.set(iid, col), iid) for iid in self.tree.get_children("")]
        if col == "risk":
            order = {RISK_GREEN: 0, RISK_YELLOW: 1, RISK_RED: 2}
            items.sort(key=lambda x: order.get(x[0], 9), reverse=reverse)
        else:
            items.sort(key=lambda x: x[0].lower(), reverse=reverse)
        for idx, (_, iid) in enumerate(items):
            self.tree.move(iid, "", idx)
        self.tree.heading(col, command=lambda: self.sort_tree(col, not reverse))

    def selected_rows(self) -> list[PackageRow]:
        return [r for r in self.rows.values() if r.selected]

    def remove_selected(self):
        if self.busy:
            return
        selected = self.selected_rows()
        if not selected:
            messagebox.showinfo(APP_TITLE, "No packages are selected.")
            return

        red = []
        missing = []
        for row in selected:
            _, risk = package_description_and_risk(row.name, row.non_removable)
            if risk == RISK_RED:
                red.append(row.name)
            if row.status == "Not found":
                missing.append(row.name)

        lines = [
            f"You selected {len(selected)} package(s).",
            "",
            "The operation will attempt:",
            "• Remove-AppxPackage -AllUsers for installed copies.",
        ]
        if self.remove_provisioned_var.get():
            lines.append("• Remove-AppxProvisionedPackage -Online -AllUsers for provisioned copies.")

        if red:
            lines += [
                "",
                f"WARNING: {len(red)} selected package(s) are RED / dangerous:",
                *[f"  • {x}" for x in red[:12]],
            ]
            if len(red) > 12:
                lines.append(f"  • ...and {len(red)-12} more")

        if missing:
            lines += [
                "",
                f"{len(missing)} default-selected package(s) were not found in the initial inventory.",
                "They will still be searched by exact package Name during removal.",
            ]

        lines += [
            "",
            "This can break Windows features or applications and may require repair/reinstallation.",
            "",
            "Continue?",
        ]

        if not messagebox.askyesno(APP_TITLE, "\n".join(lines), icon="warning"):
            return

        # Extra confirmation if Red packages were selected.
        if red:
            if not messagebox.askyesno(
                APP_TITLE,
                "RED packages are selected.\n\n"
                "Windows may become unstable or lose core functionality.\n"
                "Windows-protected NonRemovable packages may simply refuse removal.\n\n"
                "Attempt removal anyway?",
                icon="warning",
            ):
                return

        self.set_busy(True, f"Removing {len(selected)} selected package(s)...")
        threading.Thread(
            target=self._remove_worker,
            args=(selected, self.remove_provisioned_var.get()),
            daemon=True,
        ).start()

    def _remove_worker(self, selected: list[PackageRow], remove_provisioned: bool):
        results = []
        for index, row in enumerate(selected, 1):
            self.after(
                0,
                lambda i=index, n=row.name, total=len(selected):
                    self.status_var.set(f"[{i}/{total}] Removing {n}...")
            )
            result = self.remove_one(row.name, remove_provisioned)
            results.append((row.name, result))
            log(f"REMOVE {row.name}: {result}")

        self.after(0, lambda: self._remove_done(results))

    def remove_one(self, name: str, remove_provisioned: bool) -> dict:
        # Escape for single-quoted PowerShell literal.
        qname = name.replace("'", "''")

        ps = rf"""
$ErrorActionPreference = 'Continue'
$name = '{qname}'
$result = [ordered]@{{
    Name = $name
    InstalledFound = 0
    InstalledRemoved = 0
    InstalledErrors = @()
    ProvisionedFound = 0
    ProvisionedRemoved = 0
    ProvisionedErrors = @()
}}

$installed = @(Get-AppxPackage -AllUsers -Name $name -ErrorAction SilentlyContinue)
$result.InstalledFound = $installed.Count

foreach ($pkg in $installed) {{
    try {{
        Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
        $result.InstalledRemoved++
    }}
    catch {{
        $result.InstalledErrors += $_.Exception.Message
    }}
}}
"""

        if remove_provisioned:
            ps += r"""
$provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like $name })
$result.ProvisionedFound = $provisioned.Count

foreach ($pkg in $provisioned) {
    try {
        Remove-AppxProvisionedPackage -Online -PackageName $pkg.PackageName -AllUsers -ErrorAction Stop | Out-Null
        $result.ProvisionedRemoved++
    }
    catch {
        $result.ProvisionedErrors += $_.Exception.Message
    }
}
"""
        ps += "\n$result | ConvertTo-Json -Depth 5 -Compress"

        try:
            rc, out, err = run_powershell(ps, timeout=180)
            if out:
                # PowerShell warnings can occasionally precede JSON. Find the last JSON object.
                pos = out.rfind('{"Name"')
                if pos >= 0:
                    data = json.loads(out[pos:])
                else:
                    data = json.loads(out)
            else:
                data = {
                    "Name": name,
                    "InstalledFound": 0,
                    "InstalledRemoved": 0,
                    "InstalledErrors": [err or f"PowerShell exited with code {rc}"],
                    "ProvisionedFound": 0,
                    "ProvisionedRemoved": 0,
                    "ProvisionedErrors": [],
                }
            if err:
                data.setdefault("PowerShellStderr", err)
            return data
        except Exception as exc:
            return {
                "Name": name,
                "InstalledFound": 0,
                "InstalledRemoved": 0,
                "InstalledErrors": [str(exc)],
                "ProvisionedFound": 0,
                "ProvisionedRemoved": 0,
                "ProvisionedErrors": [],
            }

    def reinstall_selected(self):
        if self.busy:
            return

        selected = self.selected_rows()
        if not selected:
            messagebox.showinfo(APP_TITLE, "No packages are selected.")
            return

        lines = [
            f"You selected {len(selected)} package(s) for restore/reinstall.",
            "",
            "The program will try, in this order:",
            "• Re-register a package manifest if its files still exist locally.",
            "• Search C:\\Program Files\\WindowsApps for a matching package manifest.",
            "• For known Store apps, use WinGet + Microsoft Store as a fallback.",
            "",
            "Important: Windows has no universal Add-AppxPackage -AllUsers inverse.",
            "A package that was fully removed and deprovisioned may require Microsoft Store,",
            "Windows installation media, or another valid APPX/MSIX source.",
            "",
            "Continue?",
        ]
        if not messagebox.askyesno(APP_TITLE, "\n".join(lines), icon="question"):
            return

        self.set_busy(True, f"Restoring {len(selected)} selected package(s)...")
        threading.Thread(
            target=self._reinstall_worker,
            args=(selected,),
            daemon=True,
        ).start()

    def _reinstall_worker(self, selected: list[PackageRow]):
        results = []
        for index, row in enumerate(selected, 1):
            self.after(
                0,
                lambda i=index, n=row.name, total=len(selected):
                    self.status_var.set(f"[{i}/{total}] Restoring {n}...")
            )
            result = self.reinstall_one(row.name)
            results.append((row.name, result))
            log(f"REINSTALL {row.name}: {result}")

        self.after(0, lambda: self._reinstall_done(results))

    def reinstall_one(self, name: str) -> dict:
        qname = name.replace("'", "''")

        # Resolve Store ID for exact names and, where possible, wildcard defaults.
        store_id = STORE_IDS.get(name, "")
        if not store_id and ("*" in name or "?" in name):
            import fnmatch
            matches = [sid for pkg, sid in STORE_IDS.items()
                       if fnmatch.fnmatch(pkg.lower(), name.lower())]
            if len(matches) == 1:
                store_id = matches[0]

        qstore = store_id.replace("'", "''")

        ps = rf"""
$ErrorActionPreference = 'Continue'
$name = '{qname}'
$storeId = '{qstore}'
$result = [ordered]@{{
    Name = $name
    Method = ''
    Success = $false
    AlreadyInstalled = $false
    Message = ''
    Errors = @()
}}

# 1) If an installed/staged copy still exists, re-register its manifest.
$pkgs = @(Get-AppxPackage -AllUsers -Name $name -ErrorAction SilentlyContinue)
$manifestCandidates = @()

foreach ($pkg in $pkgs) {{
    if ($pkg.InstallLocation) {{
        $manifest = Join-Path $pkg.InstallLocation 'AppxManifest.xml'
        if (Test-Path -LiteralPath $manifest) {{
            $manifestCandidates += $manifest
        }}
    }}
}}

# 2) Search WindowsApps directly for matching package folders if Get-AppxPackage
# no longer exposes an installed registration.
if ($manifestCandidates.Count -eq 0) {{
    $windowsApps = Join-Path $env:ProgramFiles 'WindowsApps'
    if (Test-Path -LiteralPath $windowsApps) {{
        try {{
            $dirs = @(Get-ChildItem -LiteralPath $windowsApps -Directory -Force -ErrorAction Stop |
                Where-Object {{ $_.Name -like ($name + '_*') -or $_.Name -like $name }})
            foreach ($dir in $dirs) {{
                $manifest = Join-Path $dir.FullName 'AppxManifest.xml'
                if (Test-Path -LiteralPath $manifest) {{
                    $manifestCandidates += $manifest
                }}
            }}
        }}
        catch {{
            $result.Errors += "WindowsApps search: " + $_.Exception.Message
        }}
    }}
}}

$manifestCandidates = @($manifestCandidates | Select-Object -Unique)

if ($manifestCandidates.Count -gt 0) {{
    # Prefer the newest candidate path when several package versions remain.
    $manifest = $manifestCandidates | Sort-Object -Descending | Select-Object -First 1
    try {{
        Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
        $result.Method = 'Local manifest re-registration'
        $result.Success = $true
        $result.Message = "Registered: $manifest"
    }}
    catch {{
        $result.Errors += "Manifest registration: " + $_.Exception.Message
    }}
}}

# 3) Store fallback. This is for ordinary Store apps only.
if (-not $result.Success -and $storeId) {{
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($winget) {{
        try {{
            $args = @(
                'install',
                '--id', $storeId,
                '--exact',
                '--source', 'msstore',
                '--accept-source-agreements',
                '--accept-package-agreements',
                '--disable-interactivity'
            )
            $output = & $winget.Source @args 2>&1 | Out-String
            $exitCode = $LASTEXITCODE
            if ($exitCode -eq 0) {{
                $result.Method = 'Microsoft Store via WinGet'
                $result.Success = $true
                $result.Message = $output.Trim()
            }}
            else {{
                $result.Errors += "WinGet exit code $exitCode: " + $output.Trim()
            }}
        }}
        catch {{
            $result.Errors += "WinGet: " + $_.Exception.Message
        }}
    }}
    else {{
        $result.Errors += 'WinGet is not available.'
    }}
}}

if (-not $result.Success -and $result.Errors.Count -eq 0) {{
    $result.Message = 'No local package payload/manifest was found and no known Microsoft Store ID is available.'
}}

$result | ConvertTo-Json -Depth 5 -Compress
"""
        try:
            rc, out, err = run_powershell(ps, timeout=300)
            if out:
                pos = out.rfind('{"Name"')
                data = json.loads(out[pos:] if pos >= 0 else out)
            else:
                data = {
                    "Name": name,
                    "Method": "",
                    "Success": False,
                    "Message": err or f"PowerShell exited with code {rc}",
                    "Errors": [],
                }
            if err:
                data.setdefault("PowerShellStderr", err)
            return data
        except Exception as exc:
            return {
                "Name": name,
                "Method": "",
                "Success": False,
                "Message": str(exc),
                "Errors": [str(exc)],
            }

    def _reinstall_done(self, results):
        self.set_busy(False)

        succeeded = []
        failed = []

        for name, data in results:
            if bool(data.get("Success", False)):
                succeeded.append((name, data.get("Method", "Restored")))
            else:
                errors = list(data.get("Errors") or [])
                message = str(data.get("Message") or "").strip()
                if message:
                    errors.append(message)
                failed.append((name, errors or ["No restore source was available."]))

        summary = [
            "Restore/reinstall pass completed.",
            "",
            f"Restored successfully: {len(succeeded)}",
            f"Could not restore automatically: {len(failed)}",
            "",
            f"Detailed log: {LOG_PATH}",
        ]

        if succeeded:
            summary += ["", "Restored:"]
            for name, method in succeeded[:10]:
                summary.append(f"• {name} — {method}")
            if len(succeeded) > 10:
                summary.append(f"• ...and {len(succeeded)-10} more")

        if failed:
            summary += ["", "Could not restore automatically:"]
            for name, errors in failed[:8]:
                msg = " | ".join(str(e) for e in errors)
                if len(msg) > 220:
                    msg = msg[:217] + "..."
                summary.append(f"• {name}: {msg}")
            if len(failed) > 8:
                summary.append(f"• ...and {len(failed)-8} more (see log)")

        messagebox.showinfo(APP_TITLE, "\n".join(summary))
        self.refresh_packages()


    def _remove_done(self, results):
        self.set_busy(False)

        succeeded = 0
        failed = []
        not_found = []

        for name, data in results:
            found = int(data.get("InstalledFound", 0) or 0) + int(data.get("ProvisionedFound", 0) or 0)
            removed = int(data.get("InstalledRemoved", 0) or 0) + int(data.get("ProvisionedRemoved", 0) or 0)
            errors = list(data.get("InstalledErrors") or []) + list(data.get("ProvisionedErrors") or [])

            if removed > 0 and not errors:
                succeeded += 1
            elif found == 0:
                not_found.append(name)
            else:
                failed.append((name, errors or ["Removal did not report success."]))

        summary = [
            "Removal pass completed.",
            "",
            f"Successfully processed without reported errors: {succeeded}",
            f"Not found: {len(not_found)}",
            f"With errors / blocked: {len(failed)}",
            "",
            f"Detailed log: {LOG_PATH}",
        ]

        if failed:
            summary += ["", "Errors:"]
            for name, errors in failed[:8]:
                msg = " | ".join(str(e) for e in errors)
                if len(msg) > 220:
                    msg = msg[:217] + "..."
                summary.append(f"• {name}: {msg}")
            if len(failed) > 8:
                summary.append(f"• ...and {len(failed)-8} more (see log)")

        messagebox.showinfo(APP_TITLE, "\n".join(summary))
        self.refresh_packages()


def ensure_defaults_template():
    if DEFAULTS_PATH.exists():
        return
    try:
        DEFAULTS_PATH.write_text(
            "# AppX package names to check by default.\n"
            "# One Name per line. You can also paste rows copied from:\n"
            "# Get-AppxPackage | Select Name, PackageFullName\n"
            "#\n"
            "# Example:\n"
            "# Microsoft.Windows.DevHome\n"
            "# Microsoft.WindowsAlarms\n",
            encoding="utf-8",
        )
    except Exception:
        pass





def main():
    ensure_defaults_template()

    # All-users package removal requires an elevated (Admin) process!
    # Ask for elevation before loading the GUI, so inventory/removal share consistent permissions! 
    root_probe = tk.Tk()
    root_probe.withdraw()

    if not is_admin():
        answer = messagebox.askyesno(
            APP_TITLE,
            "This tool needs Administrator rights to manage AppX packages and remove them for ALL USERS.\n\n"
            "Restart it as Administrator now?",
            icon="question",
        )
        if answer:
            launched = relaunch_as_admin()
            root_probe.destroy()
            if launched:
                return
        else:
            root_probe.destroy()
            return
    else:
        root_probe.destroy()

    app = AppxRemoverGUI()
    app.mainloop()






if __name__ == "__main__":
    main()
