# SimeonOnSecurity - KMS-Auto-PS
GLVK Auto Install Script for KMS Activation

[![VirusTotal Scan](https://github.com/simeononsecurity/KMS-Auto-PS/actions/workflows/virustotal.yml/badge.svg)](https://github.com/simeononsecurity/KMS-Auto-PS/actions/workflows/virustotal.yml)

## Recommened Reading:
- [Microsoft - KMS Client Keys (GLVK)](https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys)

## How to run the script:
### Manual Install:
If manually downloaded, the script must be launched from an administrative powershell in the directory containing all the files from the [GitHub Repository](https://github.com/simeononsecurity/KMS-Auto-PS/archive/main.zip)
```ps
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
Get-ChildItem -Recurse *.ps1 | Unblock-File
.\sos-kmsglvkactivationauto.ps1
```
