# SO Upgrade Assistant

## Overview
A helper script for managing SO upgrades.
Handles stopping services/processes, downloading setup files, running installers, and restoring state.

![SO Upgrade Screenshot](https://raw.githubusercontent.com/SMControl/SO_Upgrade/main/SO_Upgrade_Screenshot3.png)

## Features
- Installs Firebird with our criteria if needed.
- Downloads latest SO setup files.
- Sets a task to check daily for new SO Setup files.
- Automatically handles the start stop state of LiveSales Service.
- Automatically handles the start stop state of both PDTWiFi's
- Ensures correct folder permissions for Client access.
- Non-Intrusive - Only assists with before and after.

**⚠️ NB - Does not support the Setup requiring a Reboot.**
> If a reboot is required, either finish the Setup manually after a reboot OR leave the reboot until the end.

## How to Run
Open **PowerShell/Terminal as Administrator**.

<details>
  <summary>Click to show "Open as Admin" image</summary>
  <img src="https://raw.githubusercontent.com/SMControl/SO_Upgrade/main/Open-as-admin2.png" alt="Open as Admin">
</details>

Copy and Paste in the following and press Enter.
  ```powershell
  irm [https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/main/soua.ps1](https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/main/soua.ps1) | iex
