# Windows Bloat Remover (by G.Tsakalos)

A lightweight Python/Tkinter GUI for inspecting, removing, and restoring Windows AppX/MSIX packages.

The tool is designed for advanced Windows users who want more control over built-in Microsoft Store apps, optional Windows components, OEM packages, and provisioned AppX packages.

> **Current version:** v1.2  
> **Platform:** Windows 10 / Windows 11  
> **Author:** George Tsakalos

---

## Features

- Reads installed AppX/MSIX packages from Windows using PowerShell.
- Displays:
  - package name
  - full package name
  - installation/provisioning status
  - simple package description
  - removal risk level
- Checkbox-based package selection.
- Search/filter support.
- Sortable package table.
- Supports a separate `default_checked_packages.txt` file.
- Default entries can be shown even if they are not returned by the initial package scan.
- Supports both:
  - exact package names
  - wildcard package patterns
- Can parse legacy PowerShell-style entries such as:

```powershell
Get-AppxPackage *MixedReality* -AllUsers | Remove-AppxPackage
```

and automatically extract:

```text
*MixedReality*
```

- Removes selected packages for all users.
- Optionally removes provisioned copies from the Windows image.
- Attempts to reinstall/restore selected packages.
- Automatically requests Administrator privileges.
- Writes detailed results to a log file.
- Includes built-in package descriptions and risk classification.

---

## Risk Levels

Each package is classified according to the likely effect of removing it.

| Color | Risk | Meaning |
|---|---|---|
| 🟢 Green | Low | Normally safe to remove if you do not use the feature/app |
| 🟡 Yellow | Medium | Removal may disable Windows functionality or another application dependency |
| 🔴 Red | Dangerous | Core Windows component, runtime, shell component, authentication component, or shared dependency |

Windows packages reported by the operating system as:

```text
NonRemovable=True
```

are automatically treated as **Red / Dangerous**.

> The risk classification is advisory only. Windows versions, installed software, OEM configurations, and dependencies vary.

---

## Requirements

- Windows 10 or Windows 11
- Python 3.x
- Tkinter
- Windows PowerShell
- Administrator privileges

Tkinter is included with the standard Windows Python installer in most installations.

No third-party Python packages are required.

---

## Files

The repository should contain at least:

```text
Win_BloatRemover.py
default_checked_packages.txt
README.md
```

At runtime the program may also create:

```text
AppxPackageRemover.log
```

---

## Running the Program

Open PowerShell or Command Prompt in the project directory and run:

```powershell
python Win_BloatRemover.py
```

The program will request Administrator elevation because removing AppX packages for all users requires elevated permissions.

---

## Package Discovery

The initial package list is based on PowerShell commands equivalent to:

```powershell
Get-AppxPackage | Select Name, PackageFullName
```

The program also performs additional checks using:

```powershell
Get-AppxPackage -AllUsers
```

and:

```powershell
Get-AppxProvisionedPackage -Online
```

This allows it to distinguish between:

- packages installed for the current user
- packages installed for other users
- provisioned packages
- packages that are not currently installed

---

## Default Checked Packages

The file:

```text
default_checked_packages.txt
```

contains packages that should be checked automatically when the program starts.

The file must be located in the same directory as the Python script.

### Exact package names

Example:

```text
Microsoft.Windows.DevHome
Microsoft.WindowsAlarms
Microsoft.WindowsCamera
Microsoft.Windows.NarratorQuickStart
Microsoft.Windows.ParentalControls
```

### Wildcard package names

Example:

```text
*MixedReality*
*3dbuilder*
*WindowsFeedbackHub*
*GetHelp*
*bingweather*
*xbox*
```

### Legacy PowerShell syntax

The program can also parse lines copied from older debloat scripts:

```powershell
Get-AppxPackage *MixedReality* -allusers | Remove-AppxPackage
Get-AppxPackage *3dbuilder* -allusers | Remove-AppxPackage
Get-AppxPackage *GetHelp* -allusers | Remove-AppxPackage
```

The wildcard package expression is extracted automatically.

### Comments

Lines beginning with `#` are ignored.

Example:

```text
# Packages I always want selected
Microsoft.Windows.DevHome
*MixedReality*
```

---

## Removing Packages

Select the desired packages and press:

```text
REMOVE NOW
```

For installed packages, the program attempts removal using the equivalent of:

```powershell
Remove-AppxPackage -Package <PackageFullName> -AllUsers
```

If:

```text
Also remove provisioned copy
```

is enabled, the program also attempts:

```powershell
Remove-AppxProvisionedPackage -Online -PackageName <PackageName> -AllUsers
```

Removing the provisioned copy prevents Windows from automatically installing that package for newly created user profiles.

### Red package protection

If one or more Red packages are selected, the program displays an additional warning before continuing.

Windows may also refuse removal of protected system packages.

---

## Reinstall / Restore

Press:

```text
REINSTALL / RESTORE
```

to attempt to restore the selected packages.

The program tries several methods.

### 1. Re-register an existing package

If the application files still exist locally, the program attempts:

```powershell
Add-AppxPackage -DisableDevelopmentMode -Register AppxManifest.xml
```

### 2. Search WindowsApps

If the package registration is missing, the program searches:

```text
C:\Program Files\WindowsApps
```

for a matching package directory and `AppxManifest.xml`.

If found, the manifest is re-registered.

### 3. Microsoft Store / WinGet fallback

For supported Microsoft Store applications whose Store IDs are known to the program, the tool can fall back to:

```powershell
winget install --id <STORE_ID> --exact --source msstore
```

---

## Important Restore Limitation

Windows does **not** provide a universal command that is the direct opposite of:

```powershell
Remove-AppxPackage -AllUsers
```

There is no universal:

```text
Add-AppxPackage -AllUsers
```

operation that can reconstruct a package after all package files and provisioning information have been removed.

If a package has been:

1. removed from all user profiles,
2. removed from the provisioned Windows image,
3. and its package files no longer exist under `WindowsApps`,

then Windows requires another valid package source.

Depending on the package, this may be:

- Microsoft Store
- WinGet
- Windows installation media
- DISM / Windows servicing
- an APPX/MSIX package
- OEM recovery software

Therefore **REINSTALL / RESTORE is intentionally best-effort**.

Core Windows components may require Windows servicing rather than ordinary AppX reinstallation.

---

## Package Descriptions

The application contains a built-in knowledge base for common Windows packages, including components such as:

- Windows Camera
- Windows Alarms / Clock
- Feedback Hub
- Mixed Reality Portal
- 3D Viewer
- 3D Builder
- Microsoft Maps
- Microsoft To Do
- Microsoft Teams
- OneNote
- Groove Music / Media Player
- Movies & TV
- Phone Link
- Xbox components
- Snipping Tool / Snip & Sketch
- Voice Recorder
- HEIF / WebP / AV1 codec extensions
- Windows Security
- Windows Search
- App Installer
- Visual C++ UWP runtimes
- .NET Native runtimes
- Microsoft UI XAML
- Windows App Runtime

Unknown packages receive a conservative fallback description rather than being silently treated as safe.

---

## Examples of Packages That Should Usually Be Treated Carefully

Examples include:

```text
Microsoft.Windows.ShellExperienceHost
Microsoft.Windows.StartMenuExperienceHost
Microsoft.Windows.Search
Microsoft.Windows.SecHealthUI
Microsoft.AAD.BrokerPlugin
Microsoft.CredDialogHost
Microsoft.VCLibs.*
Microsoft.NET.Native.*
Microsoft.UI.Xaml.*
Microsoft.WindowsAppRuntime.*
```

These packages may provide:

- Windows shell functionality
- Start menu functionality
- Windows Security UI
- authentication
- credential dialogs
- application runtimes
- shared dependencies

Removing them can break Windows features or installed software.

---

## GUI Controls

### Refresh

Reloads package information from Windows.

### Select Green

Selects currently installed packages classified as low-risk.

### Clear All

Clears all selections.

### Restore Defaults

Reloads selections from:

```text
default_checked_packages.txt
```

### REMOVE NOW

Attempts complete removal of selected packages.

### REINSTALL / RESTORE

Attempts to restore or reinstall selected packages.

---

## Log File

Operations are recorded in:

```text
AppxPackageRemover.log
```

The log includes removal and restore results and is useful when Windows blocks a package operation.

---

## Building an EXE

The application can optionally be packaged with PyInstaller.

Install PyInstaller:

```powershell
pip install pyinstaller
```

Example build:

```powershell
pyinstaller --onefile --windowed --name Win_BloatRemover Win_BloatRemover.py
```

The resulting executable will be created under:

```text
dist\
```

Keep:

```text
default_checked_packages.txt
```

beside the generated executable.

---

## Safety Warning

This is an advanced Windows administration tool.

Removing AppX packages can:

- disable Windows features
- break Microsoft Store applications
- break shared application dependencies
- remove codecs
- affect authentication
- affect shell components
- require Windows repair or reinstallation

Always review the package description and risk level before removal.

Creating a restore point or system image before aggressively removing Windows components is strongly recommended.

The author assumes no responsibility for damage, data loss, broken Windows installations, or removed applications resulting from use of this tool.

---

## Compatibility

Designed primarily for:

- Windows 10
- Windows 11

Package names, provisioning behavior, Store availability, and Windows servicing rules may differ between Windows builds.

---

## License

Choose a license appropriate for your repository.

For a simple open-source utility, the MIT License is a common option.

Example repository files:

```text
LICENSE
README.md
Win_BloatRemover.py
default_checked_packages.txt
```

---

## Version History

### v1.2

- Added **REINSTALL / RESTORE** button.
- Added local `AppxManifest.xml` re-registration.
- Added `WindowsApps` package-manifest recovery.
- Added Microsoft Store / WinGet restore fallback for supported packages.
- Expanded package descriptions.
- Improved handling of unknown packages.
- Added support for wildcard package names.
- Added support for parsing legacy `Get-AppxPackage ... | Remove-AppxPackage` lines from the defaults file.
- Preserved all-user removal and provisioned-package removal functionality.

### v1.1

- Package discovery through PowerShell.
- GUI package list.
- Risk classification.
- Package descriptions.
- Default checked package list.
- All-user package removal.
- Provisioned package removal.
- Logging.
- Administrator elevation.

---

## Disclaimer

This project is not affiliated with, endorsed by, or supported by Microsoft.

Windows, Microsoft Store, PowerShell, WinGet, and related product names are trademarks of Microsoft Corporation.
