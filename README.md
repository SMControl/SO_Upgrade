# SO Upgrade Assistant

## Overview
A helper script for managing SO upgrades.
Handles stopping services/processes, downloading setup files, running installers, and restoring state.

## Features
- Checks for admin privileges.
- Downloads and verifies SO setup files.
- Stops and restarts SO processes/services.
- Installs Firebird if needed.
- Ensures correct folder permissions.

## How to Run
1. Open **PowerShell as Administrator**.
2. Run:
   ```powershell
   irm https://raw.githubusercontent.com/SMControl/SO_Upgrade/refs/heads/main/main/soua.ps1 | iex
