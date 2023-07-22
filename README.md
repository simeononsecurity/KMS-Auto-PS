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
<a href="https://simeononsecurity.ch" target="_blank" rel="noopener noreferrer">
  <h2>Explore the World of Cybersecurity</h2>
</a>
<a href="https://simeononsecurity.ch" target="_blank" rel="noopener noreferrer">
  <img src="https://simeononsecurity.ch/img/banner.png" alt="SimeonOnSecurity Logo" width="300" height="300">
</a>

### Links:
- #### [github.com/simeononsecurity](https://github.com/simeononsecurity)
- #### [simeononsecurity.ch](https://simeononsecurity.ch)
