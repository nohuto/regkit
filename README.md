# WPR / Procmon Registry Activity Records

Records were made while using `24H2` / `IoT Enterprise LTSC 2024`- Subkeys are always included. Most activities were recorded during boot, there are some others, such as `Steam.txt`, `TLOU2.txt`, `StartAllBack.txt`, and `Lighshot.txt`, that were traced using Procmon during use. WPR is included in WADK:
```ps
winget install Microsoft.WindowsADK
```
- [Windows Performance Recorder](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/windows-performance-recorder)  
- [Process Monitor](https://learn.microsoft.com/en-us/sysinternals/downloads/procmon) ([*](https://live.sysinternals.com/))

## Records Table

| File | Path(s) |
|------|---------|
| [NVIDIA-0001.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/NVIDIA-0001.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\00XX` |
| [Disk-Storport-(990Pro).txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Disk-Storport-(990Pro).txt) | `HKLM\SYSTEM\ControlSet001\Enum\SCSI\Disk&Ven_NVMe&Prod_Samsung_SSD_990\5&33c33320&0&000000\Device Parameters\StorPort` (your path will be different) |
| [NIC-Intel.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/NIC-Intel.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00XX (Intel)` |
| [nvlddmkm.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/nvlddmkm.txt) | `HKLM\SYSTEM\ControlSet001\Services\nvlddmkm` |
| [kbdclass.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/kbdclass.txt) | `HKLM\SYSTEM\ControlSet001\Services\kbdclass` |
| [mouclass.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/mouclass.txt) | `HKLM\SYSTEM\ControlSet001\Services\mouclass` |
| [NVIDIA-Corp.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/NVIDIA-Corp.txt) | `HKLM\SOFTWARE\NVIDIA Corporation` |
| [Graphics-Drivers.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Graphics-Drivers.txt) | `HKLM\SYSTEM\ControlSet001\Control\GraphicsDrivers` |
| [stornvme.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/stornvme.txt) | `HKLM\SYSTEM\ControlSet001\Services\stornvme\Parameters` |
| [Power.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Power.txt) | `HKLM\SYSTEM\ControlSet001\Control\Power` |
| [USB-Flags.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/USB-Flags.txt) | `HKLM\SYSTEM\ControlSet001\Control\usbflags` |
| [FileSystem.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/FileSystem.txt) | `HKLM\SYSTEM\ControlSet001\Control\FileSystem` |
| [Enum-USB.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Enum-USB.txt) | `HKLM\SYSTEM\ControlSet001\Enum\USB` |
| [ControlPanel-Mouse.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/ControlPanel-Mouse.txt) | `HKCU\Control Panel\Mouse` |
| [ControlPanel-Desktop.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/ControlPanel-Desktop.txt) | `HKCU\Control Panel\Desktop` |
| [Input.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Input.txt) | `HKLM\SYSTEM\INPUT` |
| [Audio.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Audio.txt) | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio` |
| [Control-Wdf.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Control-Wdf.txt) | `HKLM\SYSTEM\ControlSet001\Control\Wdf` |
| [MultiMedia.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/MultiMedia.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MultiMedia` |
| [BrokerInfrastructure.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/BrokerInfrastructure.txt) | `HKLM\SYSTEM\ControlSet001\Services\BrokerInfrastructure` |
| [Dnscache-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Dnscache-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\Dnscache\Parameters` |
| [AFD-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/AFD-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\AFD\Parameters` |
| [Tcpip-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Tcpip-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters` |
| [NDIS-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/NDIS-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\NDIS\Parameters` |
| [Windows-Dwm.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Windows-Dwm.txt) | `HKLM\SOFTWARE\Microsoft\Windows\Dwm` |
| [Winows-NT.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Winows-NT.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows` |
| [Session-Manager.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Session-Manager.txt) | `HKLM\SYSTEM\ControlSet001\Control\Session Manager`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Memory Management`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Power`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Quota System` |
| [monitor.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/monitor.txt) | `HKLM\SYSTEM\ControlSet001\Services\monitor` |
| [mouhid.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/mouhid.txt) | `HKLM\SYSTEM\ControlSet001\Services\mouhid` |
| [kbdhid.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/kbdhid.txt) | `HKLM\SYSTEM\ControlSet001\Services\kbdhid` |
| [pci.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/pci.txt) | `HKLM\SYSTEM\ControlSet001\Enum\pci` |
| [Classpnp.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Classpnp.txt) | `HKLM\SYSTEM\ControlSet001\Control\Classpnp` |
| [Policies.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Policies.txt) | `HKLM\SYSTEM\ControlSet001\Policies` |
| [ACPI.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/ACPI.txt) | `HKLM\SYSTEM\ControlSet001\Services\ACPI \acpiex \AcpiDev \acpipagr \AcpiPmi \acpitime` |
| [WHEA.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/WHEA.txt) | `HKLM\SYSTEM\ControlSet001\Control\WHEA` |
| [PnP.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/PnP.txt) | `HKLM\SYSTEM\ControlSet001\Control\PnP` |
| [CrashControl.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/CrashControl.txt) | `HKLM\SYSTEM\ControlSet001\Control\CrashControl` |
| [Accessibility.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Accessibility.txt) | `HKCU\Control Panel\Accessibility` |
| [USBHUB3.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/USBHUB3.txt) | `HKLM\SYSTEM\ControlSet001\Services\USBHUB3` |
| [Lsa.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Lsa.txt) | `HKLM\SYSTEM\ControlSet001\Control\Lsa` |
| [OLE.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/OLE.txt) | `HKLM\SOFTWARE\Microsoft\OLE` ([*](https://learn.microsoft.com/en-us/windows/win32/com/hkey-local-machine-software-microsoft-ole)) |
| [wbem.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/wbem.txt) | `HKLM\SOFTWARE\Microsoft\wbem` |
| [disk.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/disk.txt) | `HKLM\SYSTEM\ControlSet001\Services\disk` |
| [BFE.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/BFE.txt) | `HKLM\SYSTEM\ControlSet001\Services\BFE` |
| [Cryptography.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Cryptography.txt) | `HKLM\SOFTWARE\Microsoft\Cryptography` |
| [Wisp.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Wisp.txt) | `HKCU\Software\Microsoft\Wisp` |
| [usbhub.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/usbhub.txt) | `HKLM\SYSTEM\ControlSet001\Services\usbhub` |
| [GRE-INITIALIZE.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/GRE-INITIALIZE.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\GRE_INITIALIZE` |
| [NlaSvc.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/NlaSvc.txt) | `HKLM\SYSTEM\ControlSet001\Services\NlaSvc` |
| [Policies.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Policies.txt) | `HKLM\Software\Microsoft\Windows\CurrentVersion\Policies` |
| [CV-Explorer.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/CV-Explorer.txt) | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer` |
| [StorPort.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/StorPort.txt) | `HKLM\SYSTEM\ControlSet001\Control\StorPort` |
| [Intel.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Intel.txt) | `HKLM\Software\Intel` |
| [Intel-00XX.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Intel-00XX.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\00XX` (Intel) |
| [TLOU2.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/TLOU2.txt) | `HKCU\Software\Naughty Dog` |
| [Windows-Defender.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Windows-Defender.txt) | `HKLM\SOFTWARE\Policies\Microsoft\Windows Defender` |
| [Error-Reporting.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Error-Reporting.txt) | `HKLM\SOFTWARE\Microsoft\WINDOWS\Windows Error Reporting` |
| [Internet-Settings.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Internet-Settings.txt) | `\Software\Microsoft\Windows\CurrentVersion\Internet Settings` |
| [LanmanServer.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/LanmanServer.txt) | `HKLM\SYSTEM\ControlSet001\Services\LanmanServer` |
| [Terminal-Server.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Terminal-Server.txt) | `HKLM\SYSTEM\ControlSet001\Control\Terminal Server` |
| [Policies-System.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Policies-System.txt) | `HKLM\SOFTWARE\Policies\Microsoft\WINDOWS\SYSTEM` |
| [Lighshot.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Lighshot.txt) | `HKCU\Software\SkillBrains\Lightshot` |
| [Steam.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/Steam.txt) | `HKCU\Software\Valve\Steam` |
| [StartAllBack.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/StartAllBack.txt) | `HKCU\Software\StartIsBack` |

</details>
