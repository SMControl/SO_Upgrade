# SO Upgrade Assistant

## Overview
A helper script for managing SO upgrades. It handles stopping services, downloading setup files, running installers, and restoring configurations.

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
   irm https://raw.githubusercontent.com/YourGitHubUser/SO_Upgrade/main/soua.ps1 | iex
