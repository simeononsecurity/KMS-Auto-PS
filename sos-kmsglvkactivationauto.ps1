#SimeonOnSecurity - https://simeononsecurity.ch
#https://docs.microsoft.com/en-us/windows-server/get-started/kmsclientkeys
#GLVK Auto Install for KMS Activation

#Require elivation for script run
#Requires -RunAsAdministrator

#Windows Activation
#Windows 8
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.2%" AND ProductType="1"') -ne $null) {
    $kmskey = "32JNW-9KQ84-P47T8-D8GGY-CWCK7"
}
#Windows 8.1
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.3%" AND ProductType="1"') -ne $null) {
    $kmskey = "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7"
}
#Windows 10
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.%" AND ProductType = "1"') -ne $null) {
    $kmskey = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
}
#Windows Server 2012 R2 
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.3%" AND ProductType="2"') -ne $null) {
    $kmskey = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"
}
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version like "6.3%" AND ProductType="3"') -ne $null) {
    $kmskey = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"
}
#Windows Server 2016
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.0.14393" AND ProductType = "2"') -ne $null) {
    $kmskey = "CB7KF-BWN84-R7R2Y-793K2-8XDDG"
}
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.0.14393" AND ProductType = "3"') -ne $null) {
    $kmskey = "CB7KF-BWN84-R7R2Y-793K2-8XDDG"
}
#Windows Server 2019
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.0.17763" AND ProductType = "2"') -ne $null) {
    $kmskey = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
}
if ((Get-WmiObject -Query 'SELECT Version,ProductType FROM Win32_OperatingSystem WHERE Version LIKE "10.0.17763" AND ProductType = "3"') -ne $null) {
    $kmskey = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
}

#Office Activation
#https://theitbros.com/how-to-change-office-product-key-from-mak-to-kms/
#Office 2013 - 31-Bit
if ((Test-Path -Path "C:\Program Files\Microsoft Office\Office15") -eq $true) {
    $officepath = "C:\Program Files\Microsoft Office\Office15"
    $officekey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
}
#Office 2013 - 64-Bit
if ((Test-Path -Path "C:\Program Files (x86)\Microsoft Office\Office15") -eq $true) {
    $officepath = "C:\Program Files (x86)\Microsoft Office\Office15"
    $officekey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
}
#Office 2016/Office 2019 - 32-Bit
if ((Test-Path -Path "C:\Program Files\Microsoft Office\Office16") -eq $true) {
    $officepath = "C:\Program Files\Microsoft Office\Office16"
    $officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    $officekey2 = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
}
#Office 2016/Office 2019 - 64-Bit
if ((Test-Path -Path "C:\Program Files (x86)\Microsoft Office\Office16") -eq $true) {
    $officepath = "C:\Program Files (x86)\Microsoft Office\Office16"
    $officekey = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
    $officekey2 = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
}

If (($officepath) -ne $null) {
    If (($officekey) -ne $null) {
        cmd /C cscript "$officepath\ospp.vbs" /inpkey:"$officekey"
    }
    If (($officekey2) -ne $null) {
        cmd /C cscript "$officepath\ospp.vbs" /inpkey:"$officekey2"
    }
    cmd /C cscript "$officepath\ospp.vbs" /act
    cmd /C cscript "$officepath\ospp.vbs" /remhst
    cmd /C cscript "$officepath\ospp.vbs" /dstatus
}
cmd /C cscript C:\windows\system32\slmgr.vbs /ipk "$kmskey"
cmd /C cscript C:\windows\system32\slmgr.vbs /ato
cmd /C cscript C:\windows\system32\slmgr.vbs /dlv

