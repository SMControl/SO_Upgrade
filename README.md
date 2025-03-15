# SO Upgrade Assistant

## Overview
A helper script for managing SO upgrades.
Handles stopping services/processes, downloading setup files, running installers, and restoring state.

![SO Upgrade Screenshot](https://raw.githubusercontent.com/SMControl/SO_Upgrade/main/SO_Upgrade_Screenshot3.png)

## Features
- Checks for admin privileges.
- Downloads and verifies SO setup files.
- Stops and restarts SO processes/services.
- Installs Firebird if needed.
- Ensures correct folder permissions.

**⚠️ Does not support the Setup requiring a Reboot.**
> If a reboot is required, either finish the Setup manually after a reboot OR leave the reboot until the end.

## How to Run
1. Open **PowerShell as Administrator**.
2. Run:
   ```powershell
   irm https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/main/soua.ps1 | iex
