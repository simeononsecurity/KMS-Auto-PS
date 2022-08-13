#SimeonOnSecurity - https://simeononsecurity.ch
#https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys
#https://docs.microsoft.com/en-us/windows-server/get-started/activation-slmgr-vbs-options
#GLVK Auto Install for KMS Activation

#Require elivation for script run
#Requires -RunAsAdministrator

#Windows Activation
#Require elivation for script run
#Requires -RunAsAdministrator
Write-Output "Elevating priviledges for this process"
do {} until (Elevate-Privileges SeTakeOwnershipPrivilege)

#Set Directory to PSScriptRoot
if ((Get-Location).Path -NE $PSScriptRoot) { Set-Location $PSScriptRoot }

#Path to slmgr.vbs
$slmgrpath = "C:\windows\system32"

#Set Activation Types
Write-Host "Set Activation Types" -ForegroundColor Green
#Change Windows Activation Type to Token
#Set activation type to 1 (for AD) or 2 (for KMS) or 3 (for Token) or 0 (for all).
cscript $slmgrpath\slmgr.vbs /act-type 0

#Activate KMS Caching
Write-Host "Activate KMS Caching" -ForegroundColor Green
cmd /ccscript "$slmgrpath\slmgr.vbs" /skhc

#Windows Activation
#Installs GLVK Keys for Enterprise or Datacenter versions of windows
#https://py-kms.readthedocs.io/en/latest/Keys.html
Write-Host "Detecting Windows Version" -ForegroundColor Green
if ($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.2%" AND ProductType="1"')) {
    $kmskey = "32JNW-9KQ84-P47T8-D8GGY-CWCK7"
    $windowsversion = "Windows 8"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.3%" AND ProductType="1"')) {
    $kmskey = "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7"
    $windowsversion = "Windows 8.1"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "%10%" AND ProductType = "1"')) {
    $kmskey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
    $windowsversion = "Windows 10"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "%11%" AND ProductType = "1"')) {
    $kmskey = "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
    $windowsversion = "Windows 11"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.3%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"
    $windowsversion = "Windows Server 2012 R2"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.0.14393" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "CB7KF-BWN84-R7R2Y-793K2-8XDDG"
    $windowsversion = "Windows Server 2016"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.0.17763" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
    $windowsversion = "Windows Server 2019"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "%11%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "WX4NM-KYWYW-QJJR4-XV3QB-6VM33"
    $windowsversion = "Windows Server 2022"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.1%" AND ProductType = "1"')) {
    $kmskey = "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH"
    $windowsversion = "Windows 7"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.0%" AND ProductType = "1"')) {
    $kmskey = "VKK3X-68KWM-X2YGT-QR4M6-4BWMV"
    $windowsversion = "Windows Vista"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE (Version like "5.1%" or Version like "5.2%") AND ProductType = "1"' )) {
    $kmskey = "CD87T-HFP4C-V7X7H-8VY68-W7D7M"
    $windowsversion = "Windows XP"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "5.2%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "KJTHV-V4BVY-6R9JK-YJM7X-X7FDY"
    $windowsversion = "Windows Server 2003"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "5.2.3%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "KJTHV-V4BVY-6R9JK-YJM7X-X7FDY"
    $windowsversion = "Windows Server 2003 R2"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.0%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "7M67G-PC374-GR742-YH8V4-TCBY3"
    $windowsversion = "Windows Server 2008"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.1%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "74YFP-3QFB3-KQT8W-PMXWJ-7M648"
    $windowsversion = "Windows Server 2008 R2"
}elseif($null -ne (Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.2%" AND ProductType="2" OR ProductType="3"')) {
    $kmskey = "48HP8-DN98B-MYWDG-T2DCC-8W83P"
    $windowsversion = "Windows Server 2012"
}
Write-Host "$windowsversion detected" -ForegroundColor yellow

#Office Activation
#Installs GLVK for Office LTSC Professional Plus or just Professional Plus where available
#https://theitbros.com/how-to-change-office-product-key-from-mak-to-kms/
#https://docs.microsoft.com/en-us/deployoffice/vlactivation/gvlks
Write-Host "Detecting Installed Office Versions"

if ((Test-Path -Path "C:\Program Files\Microsoft Office\Office15") -eq $true) {
    #Office 2013 - 64-Bit
    $officepath = '"C:\Program Files\Microsoft Office\Office15"'
    $officekey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
    $officeversion = "Office 2013 64-Bit"
}elseif((Test-Path -Path "C:\Program Files (x86)\Microsoft Office\Office15") -eq $true) {
    #Office 2013 - 32-Bit
    $officepath = '"C:\Program Files (x86)\Microsoft Office\Office15"'
    $officekey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
    $officeversion = "Office 2013 32-Bit"
}elseif((Test-Path -Path "C:\Program Files\Microsoft Office\Office16") -eq $true) {
    #Office 2016/Office 2019 - 64-Bit
    $officepath = '"C:\Program Files\Microsoft Office\Office16"'
    $officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    $officekey2 = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
    $officeversion = "Office 2016/2019 64-Bit"
}elseif((Test-Path -Path "C:\Program Files (x86)\Microsoft Office\Office16") -eq $true) {
    #Office 2016/Office 2019 - 32-Bit
    $officepath = '"C:\Program Files (x86)\Microsoft Office\Office16"'
    $officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    $officekey2 = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
    $officeversion = "Office 2016/2019 32-Bit"
}elseif((Test-Path -Path "C:\Program Files\Microsoft Office\Office21") -eq $true) {
    #Office 2021 - 64-Bit
    #assumed path need to verify
    $officepath = '"C:\Program Files\Microsoft Office\Office21"'
    $officekey = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
    $officeversion = "Office 2021 64-Bit"
}elseif((Test-Path -Path "C:\Program Files (x86)\Microsoft Office\Office21") -eq $true) {
    #Office 2021 - 32-Bit
    #assumed path need to verify
    $officepath = '"C:\Program Files (x86)\Microsoft Office\Office21"'
    $officekey = "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
    $officeversion = "Office 2021 32-Bit"
}

Write-Output "$officepath"
Write-Output "$officekey"
Write-Output "$officekey2"
Write-Output "$officeversion"

#Restart Software Protection Service
Write-Host "Restart Software Protection Service" -ForegroundColor Green
net stop sppsvc ; net start sppsvc

If ($null -ne ($officepath)) {
    If ($null -ne ($officekey)) {
        Write-Host "$OfficeVersion is installed" -ForegroundColor Green
        Write-Host "Attempting to activate" -ForegroundColor Green
        cmd /C cscript $officepath\ospp.vbs /inpkey:"$officekey"
    }
    else {
        Write-Host "No Office Products Installed" -ForegroundColor Yellow
    }
    If ($null -ne ($officekey2)) {
        cmd /C cscript $officepath\ospp.vbs /inpkey:"$officekey2"
    }
    cscript $officepath\ospp.vbs /act
    cscript $officepath\ospp.vbs /remhst
    cscript $officepath\ospp.vbs /dstatusall
}
else {
    Write-Host "No Office Products Installed" -ForegroundColor Yellow
}

If (($kmskey) -ne $null) {
    cmd /C cscript C:\windows\system32\slmgr.vbs /ipk "$kmskey"
    cmd /C cscript C:\windows\system32\slmgr.vbs /ato
    cmd /C cscript C:\windows\system32\slmgr.vbs /dlv
}


