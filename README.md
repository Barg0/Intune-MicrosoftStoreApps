# 🏪 Intune Microsoft Store Apps

Automates the deployment of Microsoft Store apps to Microsoft Intune using the [Microsoft Graph Beta API](https://learn.microsoft.com/en-us/graph/api/intune-apps-wingetapp-create?view=graph-rest-beta). Define your apps in a CSV, fetch metadata via winget, optionally add icons, and deploy everything in one run. 🚀

## 📋 Overview

The project consists of two scripts that run in sequence:

| Script | Runs on | Purpose |
|---|---|---|
| `fetch.ps1` | 🪟 Windows (winget required) | Fetches app metadata (publisher, description, URLs) from winget and saves it as JSON |
| `deploy.ps1` | 💻 Windows or any OS with PowerShell 7+ | Authenticates to Graph API, creates the apps in Intune, and optionally creates assignment groups |

## ✅ Prerequisites

- 🪟 **Windows** with [winget](https://github.com/microsoft/winget-cli) installed (for `fetch.ps1`)
- ⚡ **PowerShell 7+** recommended
- 📦 **Microsoft Graph PowerShell modules**:
  ```powershell
  Install-Module -Name Microsoft.Graph, Microsoft.Graph.Beta -Scope CurrentUser
  ```
- 🔑 An Entra ID account with the following Graph API permissions:
  - `DeviceManagementApps.ReadWrite.All`
  - `Group.ReadWrite.All` (only if group creation is enabled)

## 📁 Project Structure

```
├── apps.csv              # App definitions (name, store ID, install context)
├── fetch.ps1             # Fetches metadata from winget
├── deploy.ps1            # Deploys apps to Intune via Graph API
├── metadata/             # Per-app JSON metadata (created by fetch.ps1)
├── icons/                # Optional PNG icons for apps
├── jsons/
│   └── language.json     # Winget localization mappings (non-English Windows support)
└── logs/                 # Script execution logs
```

## 🛠️ Usage

### 1️⃣ Define your apps in `apps.csv`

```csv
ApplicationName,StoreId,InstallContext
Microsoft Photos,9WZDNCRFJBH4,user
Mozilla Firefox,9NZVDKPMR9RD,system
Visual Studio Code,XP9KHM4BK9FZ7Q,user
```

| Column | Description |
|---|---|
| `ApplicationName` | Display name for the app in Intune |
| `StoreId` | Microsoft Store package identifier (see below) |
| `InstallContext` | `user` or `system` (see below) |

#### 🔎 Finding the StoreId

The easiest way to find the Store ID is with winget:

```powershell
winget search "Firefox"
```

This returns a table with the app name, ID, and source. Use the **Id** column value from results where the **Source** is `msstore`:

```
Name            Id            Version  Source
-----------------------------------------------
Mozilla Firefox 9NZVDKPMR9RD  Unknown  msstore
```

Alternatively, you can find the ID in the Microsoft Store URL. For example:
`https://apps.microsoft.com/detail/9NZVDKPMR9RD` -- the last segment is the StoreId.

#### 💻 InstallContext explained

The `InstallContext` column controls how the app is installed on target devices:

| Value | Intune `runAsAccount` | When to use |
|---|---|---|
| `user` | `User` | Apps that are personal or per-user (e.g. Photos, Notepad, Calculator). Installs in the logged-in user's context. |
| `system` | `System` | Apps that should be available machine-wide or require elevated permissions (e.g. Company Portal, Firefox, Adobe tools). Installs via the SYSTEM account. |

If unsure, `system` is generally the safer default for managed environments. Use `user` for lightweight apps that don't need admin privileges.

### 2️⃣ Fetch metadata

Run on a Windows machine with winget installed:

```powershell
.\fetch.ps1
```

This queries `winget show --id <StoreId>` for each app and saves metadata to `metadata/<StoreId>.json`. Already-fetched apps are skipped automatically. ⏭️

To re-fetch metadata for an app, delete its JSON file from `metadata/` and run again. 🔄

### 3️⃣ Add icons (optional)

🎨 Place PNG files in the `icons/` folder. The matching logic works as follows:

1. **Exact match** -- `Microsoft Photos.png` matches the app named "Microsoft Photos" ✅
2. **Prefix match** -- `Microsoft.png` would match "Microsoft Photos", "Microsoft Teams", etc. The longest matching prefix wins. 🏆

### 4️⃣ Deploy to Intune

```powershell
.\deploy.ps1
```

The script will:
1. 🔐 Authenticate to Microsoft Graph (interactive browser sign-in)
2. 📄 Load metadata from `metadata/` for each app
3. 🖼️ Match icons from `icons/` if available
4. 📤 Create each app in Intune as a **winGetApp** (Microsoft Store app)
5. ⏭️ Skip apps that already exist in Intune (matched by display name)
6. 👥 Optionally create Required (RQ) and Available (AV) assignment groups

## ⚙️ Configuration

Configuration variables are at the top of `deploy.ps1`:

| Variable | Default | Description |
|---|---|---|
| `$enableGroupCreation` | `$true` | 👥 Automatically create RQ/AV security groups and assign them to the app |
| `$groupNamingAppSuffix` | `$true` | 🏷️ `$true`: "Win - SW - RQ - AppName", `$false`: "Win - SW - AppName - RQ" |
| `$groupCreationBlacklist` | *(see script)* | 🚫 Array of StoreIds to skip group creation for (e.g. built-in Windows apps) |

## 🗂️ Metadata Fields

The following fields are fetched from winget and sent to the Graph API:

| Metadata Field | Graph API Property | Notes |
|---|---|---|
| `Name` | `displayName` | Falls back to CSV `ApplicationName` |
| `Description` | `description` | |
| `Publisher` | `publisher`, `developer` | |
| `PublisherUrl` | `informationUrl` (fallback) | Used when `InformationUrl` is empty |
| `InformationUrl` | `informationUrl` | From winget PackageUrl or Homepage |
| `PrivacyUrl` | `privacyInformationUrl` | |

Additional Graph API fields set during deployment: `owner`, `notes`, `isFeatured`, `roleScopeTagIds`, `repositoryType`, `installExperience`, and `largeIcon` (from icons folder).

## 📝 Logs

Both scripts write logs to `logs/fetch.log` and `logs/deploy.log`. Enable verbose logging by setting `$logDebug = $true` at the top of either script. When debug logging is enabled, `deploy.ps1` also dumps the full Graph API request body to `logs/deploy-request-body.json`. 🔍

## 🌍 Localization

`fetch.ps1` supports non-English Windows installations. The `jsons/language.json` file contains mappings to translate localized winget output labels back to English before parsing. 🌐
