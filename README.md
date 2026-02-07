# WPR / [Procmon](https://github.com/nohuto/win-registry/blob/main/guide/procmon.md) Registry Activity Records

Records were made while using `24H2` / `IoT Enterprise LTSC 2024`- Subkeys are always included. Most activities were recorded during boot, there are some others, such as `Steam.txt`, `TLOU2.txt`, `StartAllBack.txt`, and `Lighshot.txt`, that were traced using Procmon during use. You can get the Windowos Performance Toolkit from [ADK](https://go.microsoft.com/fwlink/?linkid=2337875), or install it via `winget`, but this will install more than the Performance Toolkit.

```powershell
winget install Microsoft.WindowsADK
```

- [Windows Performance Recorder](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/windows-performance-recorder)  
- [Process Monitor](https://learn.microsoft.com/en-us/sysinternals/downloads/procmon) ([*](https://live.sysinternals.com/))

## ToC

- [Guides](https://github.com/nohuto/win-registry?tab=readme-ov-file#guides)
- [Records Table](https://github.com/nohuto/win-registry?tab=readme-ov-file#records-table)
- [Registry Values Research](https://github.com/nohuto/win-registry?tab=readme-ov-file#registry-values-research)
  - [DXG Kernel Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#dxg-kernel-values)
  - [Session Manager Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#session-manager-values)
  - [Power Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#power-values)
  - [DWM Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#dwm-values)
  - [USBFLAGS/USBHUB/USB Values](https://github.com/nohuto/win-registry#usbusbhubusbflags-values)
  - [BCD Edits](https://github.com/nohuto/win-registry#bcd-edits)
  - [Intel NIC Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#intel-nic-values)
  - [MMCSS Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#mmcss-values)
  - [StorNVMe Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#stornvme-values)
  - [Miscellaneous Values](https://github.com/nohuto/win-registry?tab=readme-ov-file#miscellaneous-values)

## Guides

- [procmon.md](https://github.com/nohuto/win-registry/blob/main/guide/procmon.md) - Guide on how to trace registry activity for a specific app.
- [wpr-wpa.md](https://github.com/nohuto/win-registry/blob/main/guide/wpr-wpa.md) - Guide on how to create a boot registry activity trace and how to format it so [regkit](https://github.com/nohuto/regkit) can use it.

## Records Table

| File | Path(s) |
|------|---------|
| [23H2.txt](https://github.com/nohuto/win-registry/blob/main/records/23H2.txt) | System trace of 23H2, used for [regkit](https://github.com/nohuto/regkit) (see [trace-menu](https://github.com/nohuto/regkit#trace-menu)) |
| [24H2.txt](https://github.com/nohuto/win-registry/blob/main/records/24H2.txt) | System trace of 24H2, used for [regkit](https://github.com/nohuto/regkit) (see [trace-menu](https://github.com/nohuto/regkit#trace-menu)) |
| [25H2.txt](https://github.com/nohuto/win-registry/blob/main/records/25H2.txt) | System trace of 25H2, used for [regkit](https://github.com/nohuto/regkit) (see [trace-menu](https://github.com/nohuto/regkit#trace-menu)) |
| [ACPI.txt](https://github.com/nohuto/win-registry/blob/main/records/ACPI.txt) | `HKLM\SYSTEM\ControlSet001\Services\ACPI`<br>`HKLM\SYSTEM\ControlSet001\Services\acpiex`<br>`HKLM\SYSTEM\ControlSet001\Services\AcpiDev`<br>`HKLM\SYSTEM\ControlSet001\Services\acpipagr`<br>`HKLM\SYSTEM\ControlSet001\Services\AcpiPmi`<br>`HKLM\SYSTEM\ControlSet001\Services\acpitime` |
| [AFD-Parameters.txt](https://github.com/nohuto/win-registry/blob/main/records/AFD-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\AFD\Parameters` |
| [Accessibility.txt](https://github.com/nohuto/win-registry/blob/main/records/Accessibility.txt) | `HKCU\Control Panel\Accessibility` |
| [Audio.txt](https://github.com/nohuto/win-registry/blob/main/records/Audio.txt) | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio` |
| [BFE.txt](https://github.com/nohuto/win-registry/blob/main/records/BFE.txt) | `HKLM\SYSTEM\ControlSet001\Services\BFE` |
| [BrokerInfrastructure.txt](https://github.com/nohuto/win-registry/blob/main/records/BrokerInfrastructure.txt) | `HKLM\SYSTEM\ControlSet001\Services\BrokerInfrastructure` |
| [CV-Explorer.txt](https://github.com/nohuto/win-registry/blob/main/records/CV-Explorer.txt) | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer` |
| [Classpnp.txt](https://github.com/nohuto/win-registry/blob/main/records/Classpnp.txt) | `HKLM\SYSTEM\ControlSet001\Control\Classpnp` |
| [Control-Wdf.txt](https://github.com/nohuto/win-registry/blob/main/records/Control-Wdf.txt) | `HKLM\SYSTEM\ControlSet001\Control\Wdf` |
| [ControlPanel-Desktop.txt](https://github.com/nohuto/win-registry/blob/main/records/ControlPanel-Desktop.txt) | `HKCU\Control Panel\Desktop` |
| [ControlPanel-Mouse.txt](https://github.com/nohuto/win-registry/blob/main/records/ControlPanel-Mouse.txt) | `HKCU\Control Panel\Mouse` |
| [CrashControl.txt](https://github.com/nohuto/win-registry/blob/main/records/CrashControl.txt) | `HKLM\SYSTEM\ControlSet001\Control\CrashControl` |
| [Cryptography.txt](https://github.com/nohuto/win-registry/blob/main/records/Cryptography.txt) | `HKLM\SOFTWARE\Microsoft\Cryptography` |
| [Disk-Storport-(990Pro).txt](https://github.com/nohuto/win-registry/blob/main/records/Disk-Storport-(990Pro).txt) | `HKLM\SYSTEM\ControlSet001\Enum\SCSI\Disk&Ven_NVMe&Prod_Samsung_SSD_990\5&33c33320&0&000000\Device Parameters\StorPort` (your path will be different) |
| [Dnscache-Parameters.txt](https://github.com/nohuto/win-registry/blob/main/records/Dnscache-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\Dnscache\Parameters` |
| [Enum-USB.txt](https://github.com/nohuto/win-registry/blob/main/records/Enum-USB.txt) | `HKLM\SYSTEM\ControlSet001\Enum\USB` |
| [Error-Reporting.txt](https://github.com/nohuto/win-registry/blob/main/records/Error-Reporting.txt) | `HKLM\SOFTWARE\Microsoft\WINDOWS\Windows Error Reporting` |
| [FileSystem.txt](https://github.com/nohuto/win-registry/blob/main/records/FileSystem.txt) | `HKLM\SYSTEM\ControlSet001\Control\FileSystem` |
| [GRE-INITIALIZE.txt](https://github.com/nohuto/win-registry/blob/main/records/GRE-INITIALIZE.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\GRE_INITIALIZE` |
| [Graphics-Drivers.txt](https://github.com/nohuto/win-registry/blob/main/records/Graphics-Drivers.txt) | `HKLM\SYSTEM\ControlSet001\Control\GraphicsDrivers` |
| [Input.txt](https://github.com/nohuto/win-registry/blob/main/records/Input.txt) | `HKLM\SYSTEM\INPUT` |
| [Intel-00XX.txt](https://github.com/nohuto/win-registry/blob/main/records/Intel-00XX.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\00XX` (Intel) |
| [Intel.txt](https://github.com/nohuto/win-registry/blob/main/records/Intel.txt) | `HKLM\Software\Intel` |
| [Internet-Settings.txt](https://github.com/nohuto/win-registry/blob/main/records/Internet-Settings.txt) | `\Software\Microsoft\Windows\CurrentVersion\Internet Settings` |
| [LanmanServer.txt](https://github.com/nohuto/win-registry/blob/main/records/LanmanServer.txt) | `HKLM\SYSTEM\ControlSet001\Services\LanmanServer` |
| [Lighshot.txt](https://github.com/nohuto/win-registry/blob/main/records/Lighshot.txt) | `HKCU\Software\SkillBrains\Lightshot` |
| [Lsa.txt](https://github.com/nohuto/win-registry/blob/main/records/Lsa.txt) | `HKLM\SYSTEM\ControlSet001\Control\Lsa` |
| [MultiMedia.txt](https://github.com/nohuto/win-registry/blob/main/records/MultiMedia.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MultiMedia` |
| [NDIS-Parameters.txt](https://github.com/nohuto/win-registry/blob/main/records/NDIS-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\NDIS\Parameters` |
| [NetBT.txt](https://github.com/nohuto/win-registry/blob/main/records/NetBT.txt) | `HKLM\SYSTEM\ControlSet001\Services\NetBT` |
| [NIC-Intel.txt](https://github.com/nohuto/win-registry/blob/main/records/NIC-Intel.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00XX` (Intel) |
| [NIC-Intel-IDA.txt](https://github.com/nohuto/win-registry/blob/main/records/NIC-Intel-IDA.txt) | Same path as above, but values were found via decompiling (some may not get read) |
| [NVIDIA-DispGUID.txt](https://github.com/nohuto/win-registry/blob/main/records/NVIDIA-DispGUID.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\00XX` |
| [NVIDIA-Corp.txt](https://github.com/nohuto/win-registry/blob/main/records/NVIDIA-Corp.txt) | `HKLM\SOFTWARE\NVIDIA Corporation` |
| [NlaSvc.txt](https://github.com/nohuto/win-registry/blob/main/records/NlaSvc.txt) | `HKLM\SYSTEM\ControlSet001\Services\NlaSvc` |
| [OLE.txt](https://github.com/nohuto/win-registry/blob/main/records/OLE.txt) | `HKLM\SOFTWARE\Microsoft\OLE` ([*](https://learn.microsoft.com/en-us/windows/win32/com/hkey-local-machine-software-microsoft-ole)) |
| [PnP.txt](https://github.com/nohuto/win-registry/blob/main/records/PnP.txt) | `HKLM\SYSTEM\ControlSet001\Control\PnP` |
| [Policies-System.txt](https://github.com/nohuto/win-registry/blob/main/records/Policies-System.txt) | `HKLM\SOFTWARE\Policies\Microsoft\WINDOWS\SYSTEM` |
| [Policies.txt](https://github.com/nohuto/win-registry/blob/main/records/Policies.txt) | `HKLM\SYSTEM\ControlSet001\Policies` |
| [Policies.txt](https://github.com/nohuto/win-registry/blob/main/records/CV-Policies.txt) | `HKLM\Software\Microsoft\Windows\CurrentVersion\Policies` |
| [Power.txt](https://github.com/nohuto/win-registry/blob/main/records/Power.txt) | `HKLM\SYSTEM\ControlSet001\Control\Power` |
| [Session-Manager.txt](https://github.com/nohuto/win-registry/blob/main/records/Session-Manager.txt) | `HKLM\SYSTEM\ControlSet001\Control\Session Manager`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Memory Management`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Power`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Quota System` |
| [StartAllBack.txt](https://github.com/nohuto/win-registry/blob/main/records/StartAllBack.txt) | `HKCU\Software\StartIsBack` |
| [Steam.txt](https://github.com/nohuto/win-registry/blob/main/records/Steam.txt) | `HKCU\Software\Valve\Steam` |
| [StorPort.txt](https://github.com/nohuto/win-registry/blob/main/records/StorPort.txt) | `HKLM\SYSTEM\ControlSet001\Control\StorPort` |
| [TLOU2.txt](https://github.com/nohuto/win-registry/blob/main/records/TLOU2.txt) | `HKCU\Software\Naughty Dog` |
| [Tcpip-Parameters.txt](https://github.com/nohuto/win-registry/blob/main/records/Tcpip-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters` |
| [Terminal-Server.txt](https://github.com/nohuto/win-registry/blob/main/records/Terminal-Server.txt) | `HKLM\SYSTEM\ControlSet001\Control\Terminal Server` |
| [USB-Flags.txt](https://github.com/nohuto/win-registry/blob/main/records/USB-Flags.txt) | `HKLM\SYSTEM\ControlSet001\Control\usbflags` |
| [USBHUB3.txt](https://github.com/nohuto/win-registry/blob/main/records/USBHUB3.txt) | `HKLM\SYSTEM\ControlSet001\Services\USBHUB3` |
| [WHEA.txt](https://github.com/nohuto/win-registry/blob/main/records/WHEA.txt) | `HKLM\SYSTEM\ControlSet001\Control\WHEA` |
| [Windows-Defender.txt](https://github.com/nohuto/win-registry/blob/main/records/Windows-Defender.txt) | `HKLM\SOFTWARE\Policies\Microsoft\Windows Defender` |
| [Windows-Dwm.txt](https://github.com/nohuto/win-registry/blob/main/records/Windows-Dwm.txt) | `HKLM\SOFTWARE\Microsoft\Windows\Dwm` |
| [Winows-NT.txt](https://github.com/nohuto/win-registry/blob/main/records/Winows-NT.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows` |
| [Wisp.txt](https://github.com/nohuto/win-registry/blob/main/records/Wisp.txt) | `HKCU\Software\Microsoft\Wisp` |
| [disk.txt](https://github.com/nohuto/win-registry/blob/main/records/disk.txt) | `HKLM\SYSTEM\ControlSet001\Services\disk` |
| [kbdclass.txt](https://github.com/nohuto/win-registry/blob/main/records/kbdclass.txt) | `HKLM\SYSTEM\ControlSet001\Services\kbdclass` |
| [kbdhid.txt](https://github.com/nohuto/win-registry/blob/main/records/kbdhid.txt) | `HKLM\SYSTEM\ControlSet001\Services\kbdhid` |
| [monitor.txt](https://github.com/nohuto/win-registry/blob/main/records/monitor.txt) | `HKLM\SYSTEM\ControlSet001\Services\monitor` |
| [mouclass.txt](https://github.com/nohuto/win-registry/blob/main/records/mouclass.txt) | `HKLM\SYSTEM\ControlSet001\Services\mouclass` |
| [mouhid.txt](https://github.com/nohuto/win-registry/blob/main/records/mouhid.txt) | `HKLM\SYSTEM\ControlSet001\Services\mouhid` |
| [nvlddmkm.txt](https://github.com/nohuto/win-registry/blob/main/records/nvlddmkm.txt) | `HKLM\SYSTEM\ControlSet001\Services\nvlddmkm` |
| [pci.txt](https://github.com/nohuto/win-registry/blob/main/records/pci.txt) | `HKLM\SYSTEM\ControlSet001\Enum\pci` |
| [stornvme.txt](https://github.com/nohuto/win-registry/blob/main/records/stornvme.txt) | `HKLM\SYSTEM\ControlSet001\Services\stornvme\Parameters` |
| [usbhub.txt](https://github.com/nohuto/win-registry/blob/main/records/usbhub.txt) | `HKLM\SYSTEM\ControlSet001\Services\usbhub` |
| [wbem.txt](https://github.com/nohuto/win-registry/blob/main/records/wbem.txt) | `HKLM\SOFTWARE\Microsoft\wbem` |

# Registry Values Research

Since many values/keys are unknown, I took some time to create several lists showing their default values, used keys, and additional notes. I created them using IDA, WinDbg ([Symbols Memory Dump](https://github.com/nohuto/sym-mem-dump)), WinObjEx, and Windows Internals E7 P1. See the `assets` folder for references.

> [Windows Internals E7](https://github.com/nohuto/windows-books/releases)  
> [WinObjEx64](https://github.com/hfiref0x/WinObjEx64)  
> [WinDbg](https://learn.microsoft.com/en-us/windows-hardware/drivers/debugger/)  
> [Symbols Memory Dump](https://github.com/nohuto/sym-mem-dump)  

---

## DXG Kernel Values

These are default values I found in `dxgkrnl.sys`, see [dxgkrnl.c](https://github.com/nohuto/win-registry/blob/main/assets/dxgkrnl.c) for pseudocode snippets I used / [records/Graphics-Drivers.txt](https://github.com/nohuto/win-registry/blob/main/records/Graphics-Drivers.txt) for all values that get read on boot.

The `GraphicsDrivers\Scheduler` / `GraphicsDrivers\MemoryManager` values are from `dxgmms2.sys`, I used the drivers from 23H2/25H2 since they differ at some point. See [dxgmms2](https://github.com/nohuto/win-registry/blob/main/assets/dxgmms2) for all used files.

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers"
    "MiracastUseIhvDriver"; v3 = 2;
    "MiracastForceDisable"; v2 = 2;

    "ContextNoPatchMode"; v38 = 0
    "CreateGdiPrimaryOnSlaveGpu"; v48 = 0
    "CrtcPhaseFrames"; v57 = 2
    "DeadlockPulse"; v54 = 5000
    "DeadlockPulseTolerance"; v55 = 500
    "DeadlockTimeout"; v53 = 30000
    "DisableBadDriverCheckForHwProtection"; v70 = 0 // DWORD
    "DisableBoostedVSyncVirtualization"; v59 = 0 // DWORD
    "DisableGdiContextGpuVa"; v41 = 0
    "DisableIndependentVidPnVSync"; v56 = 0 // DWORD
    "DisableMonitoredFenceGpuVa"; v43 = 0
    "DisableMultiSourceMPOCheck"; v76 = 0
    "DisableOverlays"; v67 = 0
    "DisablePagingContextGpuVa"; v42 = 0
    "DisableSecondaryIFlipSupport"; v71 = 0
    "DriverManagesResidencyOverride"; v46 = 1
    "DriverStoreCopyMode"; v33 = 1
    "EnableBasicRenderGpuPv"; v60 = 0
    "EnableDecodeMPO"; v69 = 1
    "EnableFbrValidation"; v58 = 1
    "EnableMultiPlaneOverlay3DDIs"; v73 = 0
    "EnableOfferReclaimOnDriver"; v37 = 1
    "EnablePanelFitterSupport"; v100 = 0
    "EnableTimedCalls"; v49 = 0
    "EnableWDDM23Synchronization"; v50 = 0
    "Force32BitFences"; v68 = 0
    "ForceDirectFlip"; v66 = 0
    "ForceEnableDxgMms2"; v39 = 0
    "ForceExplicitResidencyNotification"; v44 = 0 // DWORD
    "ForceInitPagingProcessVaSpace"; v40 = 0
    "ForceReplicateGdiContent"; v47 = 0
    "ForceSecondaryIFlipSupport"; v72 = 0
    "ForceSecondaryMPOSupport"; v97 = 0
    "ForceSurpriseRemovalSupport"; v75 = 0
    "ForceVariableRefresh"; v52 = 0
    "GdiPhysicalAdapterIndex"; v74 = 0
    "GpuPriorityChangeMode"; v64 = 1
    "HighPriorityCompletionMode"; v63 = 1
    "InitialPagingQueueFenceValue"; v45 = 7000
    "IoMmuFlags"; v51 = 0
    "KnownProcessBoostMode"; v61 = 1
    "LeanMemoryLimit"; v122 = 1395864371
    "LeanMemoryLimit"; v123 = 16
    "NumVirtualFunctions"; v65 = 0
    "SmallQuantumMode"; v62 = 1

    "DefaultActiveIdleThreshold"; v191 = 2000;
    "DefaultD3TransitionIdleLongTimeThreshold"; v195 = 60000;
    "DefaultD3TransitionIdleShortTimeThreshold"; v193 = 10000;
    "DefaultD3TransitionIdleVeryLongTimeThreshold"; v197 = 60000;
    "DefaultD3TransitionLatencyActivelyUsed"; v192 = 80;
    "DefaultD3TransitionLatencyIdleLongTime"; v196 = 140000;
    "DefaultD3TransitionLatencyIdleMonitorOff"; v200 = 250000;
    "DefaultD3TransitionLatencyIdleNoContext"; v199 = 250000;
    "DefaultD3TransitionLatencyIdleShortTime"; v194 = 80000;
    "DefaultD3TransitionLatencyIdleVeryLongTime"; v198 = 200000;
    "DefaultExpectedResidency"; v176 = 2000;
    "DefaultIdleThresholdIdle0"; v187 = 200;
    "DefaultIdleThresholdIdle0MonitorOff"; v222 = 100;
    "DefaultLatencyToleranceIdle0"; v184 = 80;
    "DefaultLatencyToleranceIdle0MonitorOff"; v188 = 2000;
    "DefaultLatencyToleranceIdle1"; v185 = 15000;
    "DefaultLatencyToleranceIdle1MonitorOff"; v189 = 50000;
    "DefaultLatencyToleranceMemory"; v201 = 15000;
    "DefaultLatencyToleranceMemoryNoContext"; v202 = 30000;
    "DefaultLatencyToleranceNoContext"; v186 = 35000;
    "DefaultLatencyToleranceNoContextMonitorOff"; v190 = 100000;
    "DefaultLatencyToleranceOther"; v175 = -1;
    "DefaultLatencyToleranceTimerPeriod"; v183 = 200;
    "DefaultMemoryRefreshLatencyToleranceActivelyUsed"; v203 = 80;
    "DefaultMemoryRefreshLatencyToleranceIdleShortTime"; v204 = 15000;
    "DefaultMemoryRefreshLatencyToleranceMonitorOff"; v206 = 80000;
    "DefaultMemoryRefreshLatencyToleranceNoContext"; v205 = 30000;
    "DefaultPowerNotRequiredTimeout"; v209 = 25000;
    "DisableDevicePowerRequired"; v179 = 0;
    "DisablePStateManagement"; v181 = 0;
    "EnablePODebounce"; v180 = 0;
    "EnableRuntimePowerManagement"; v178 = 1;
    "lowdebounce"; v182 = 3;
    "MonitorLatencyTolerance"; v208 = 300000;
    "MonitorRefreshLatencyTolerance"; v207 = 17000;
    "uglitch"; v168 = 900;
    "uhigh"; v169 = 700;
    "uideal"; v167 = 500;
    "ulow"; v170 = 300;

    "AllowAdvancedEtwLogging"; v72 = 0;
    "DiagnosticsBufferExpansionTime"; v58 = 300;
    "EnableFuzzing"; v64 = 0;
    "EnableHMDTestMode"; v67 = 0;
    "EnableIgnoreWin32ProcessStatus"; v66 = 0;
    "ExternalDiagnosticsBufferMultiplier"; v59 = 1;
    "ExternalDiagnosticsBufferSize"; v56 = 16384;
    "ForceUsb4MonitorSupport"; g_bDbgForceUsb4MonitorSupport = 0;
    "InternalDiagnosticsBufferMultiplier"; v57 = 2;
    "InternalDiagnosticsBufferSize"; v55 = 65536;
    "InvestigationDebugParameter"; v65 = 0;
    "MaximumAdapterCount"; v60 = 32;
    "NodeUsageTelemetryTimerInterval"; v73 = v73; // ?
    "PreserveFirmwareMode"; v68 = 0;
    "PreventFullscreenWireFormatChange"; v69 = 0;
    "RapidHpdMaxChainInMilliseconds"; v71 = 0;
    "RapidHpdTimeoutInMilliseconds"; v70 = 0;
    "TerminationListSizeLimit"; v62 = 67108864;
    "TreatUsb4MonitorAsNormal"; g_bDbgTreatUsb4MonitorAsNormal = 0;
    "Usb4MonitorDpcdDP_IN_Adapter_Number"; g_DbgUsb4MonitorDpcdDP_IN_Adapter_Number = 0;
    "Usb4MonitorDpcdUSB4_Driver_ID"; g_DbgUsb4MonitorDpcdUSB4_Driver_ID = 0;
    "Usb4MonitorPowerOnDelayInSeconds"; g_DbgUsb4MonitorPowerOnDelayInSeconds = 0;
    "Usb4MonitorTargetId"; g_DbgUsb4MonitorTargetId = 0;
    "ValidateWDDMCaps"; v63 = 0;
    "WDDM2LockManagement"; v61 = 1;

    "DisableVaBackedVm"; g_VgpuDisableVaBackedVm = 0;
    "DisableVersionMismatchCheck"; v52 = 0;
    "GpuVirtualizationFlags"; v50 = (g_VgpuReplaceWarp ? 0x8 : 0x0); // bit0: CreatePVGpu=0, bit2: ForceSvm=0, bit3: ReplaceWarp=default from g_VgpuReplaceWarp ?
    "LimitNumberOfVfs"; g_LimitNumberOfVfs = 0;
    "VirtualGpuOnly"; g_VirtualGpuOnly = 0;

    "ForceBddFallbackOnly"; v35 = 0;
    "HwSchMode"; v29 = 0;
    "HwSchOverrideBlockList"; v31 = 1;
    "HwSchTreatExperimentalAsStable"; v30 = 0;
    "MiracastDefaultRtspPort"; dword_1C0153F64 = 7236;
    "PlatformSupportMiracast"; v26 = 0; // Set to 1 on LTSC IoT Enterprise 2024 by default
    "SupportMultipleIntegratedDisplays"; v28 = 0;
    "SuspendAdapterTimerPeriod"; v27 = 500000;

    "EnableExperimentalRefreshRates"; v22 = 0;
    "RapidHPDThresholdCount"; *(_DWORD*)((char*)this + 544) = 5;
    "RapidHPDTime"; v16 = 1000;

    "TdrDdiDelay"; v11 = 5;
    "TdrDebugMode"; v12 = 2;
    "TdrDelay"; v8 = 2;
    "TdrDodPresentDelay"; v9 = 2;
    "TdrDodVSyncDelay"; v10 = 2;
    "TdrLevel"; v13 = 3;
    "TdrLimitCount"; v14 = 5;
    "TdrLimitTime"; v15 = 60;

    "DRTTestEnable"; v14 = 0; // 1484026436 = Enabled ?
    "EnableAcmSupportDeveloperPreview"; v7 = 0;
    "ForceEnableDWMClone"; v82 = 0
    "HybridInternalPanelOverrideEnable"; v13 = 0
    "IsInternalRelease"; v44 = 0
    "MultiMonSupport"; v39 = 1;
    "OutputDuplicationSessionApplicationLimit"; v14 = 4
    "TdrTestMode"; v14 = 0
    "UnsupportedMonitorModesAllowed"; v5 = 0;

    "PageFaultDebugMode" = 1; // REG_DWORD, missing/invalid or >1 -> 1

    // from procmon boot trace
    "DisableCABC" = ?;
    "ForceAccessedPhysically" = ?;
    "ForceToMapGpuVa" = ?;
    "WarpOverrideWDDMVersion" = ?;
    "WarpSupportHybridDiscrete" = = ?;
    "WarpSupportsResourceResidency" = ?;

    // miscellaneous
    "CddBootImageMode" = ?;
    "CddBootScreenMode" = ?;
    "DODPreferredPresentMoveRegeionsOverride" = ?;
    "DisableLddmSpriteTearDown" = ?;
    "DisplayBrokerShouldNotBeActive" = ?;
    "DxgKrnlVersion" = ?;
    "MinDxgKrnlVersion" = ?;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Scheduler";
    "AutoSyncToCPUPriority" = 0; // REG_DWORD
    "QuantumUnit" = 25000; // REG_DWORD
    "PreemptionQuantumUnit" = 50000; // REG_DWORD
    "EnablePreemption" = 1; // REG_DWORD
    "HwQueuedRenderPacketGroupLimit" = 2; // REG_DWORD
    "QueuedPresentLimit" = 3; // REG_DWORD
    "InitDriverFenceId" = 0; // REG_DWORD
    "CarryOverUsedQuantum" = 0; // REG_DWORD
    "EnableFlipImmediateSwFlipQueue" = 1; // REG_DWORD, found in 23H2 (not in 25H2)
    "AdjustWorkerThreadPriority" = 1; // REG_DWORD
    "CountFlipTowardHwLimit" = 0; // REG_DWORD
    "NumberOfDmaPacketPool" = 20; // REG_DWORD
    "ProfileLevel" = 2; // REG_DWORD
    "VSyncIdleTimeout" = 7; // REG_DWORD, becomes 1 when adapter version >= 1300 and flag set, <1300 min 4
    "EnableDirectSubmission" = ; // REG_DWORD, found in 25H2 (not in 23H2), default from adapter cap
    "CountPresentTowardHwLimit" = 0; // REG_DWORD
    "EnableContextDelay" = 1; // REG_DWORD, found in 23H2 (not in 25H2)
    "LogDriverVSyncCallback" = 0; // REG_DWORD, found in 23H2 (not in 25H2)
    "MaximumAllowedPreemptionDelay" = 900; // REG_DWORD
    "ContextSchedulingPenaltyDelay" = 1000; // REG_DWORD
    "BackgroundProcessMaximumAllowedPreemptionDelay" = 8; // REG_DWORD
    "ForceEnableFlipFenceModel" = 0; // REG_DWORD
    "YieldPercentage" = 10; // REG_DWORD, valid 1..0x53 else default 10
    "ForegroundPriorityBoost" = 1; // REG_DWORD
    "ForceFlipTrueImmediateMode" = 0; // REG_DWORD, values 0..2
    "MaxYieldInterval" = 16; // REG_DWORD
    "MinYieldInterval" = 8000; // REG_DWORD, found in 25H2 (not in 23H2)
    "MaxFocusGpuQuantumWithoutPresent" = 100; // REG_DWORD, 25H2 default 10 when flag set
    "HistoryLogSize" = 64; // REG_DWORD, clamped 16..0x10000, must be 16, 32, 64, 128, ... (doubling sequence)
    "HwQueuePacketCap" = ; // REG_DWORD, default from adapter cap, clamped 1..14
    "FlipDoNotFlipMode" = 0; // REG_DWORD, values 0..2
    "DdiSuspendMode" = 0; // REG_DWORD, values 0..2; found in 23H2 (not in 25H2)
    "PfnCpuOverride" = 0; // REG_DWORD, values 0..3
    "PerSourceCustomDuration" = ; // REG_DWORD, default 1 when adapter version >= 2000
    "FlipOverrideMode" = 0; // REG_DWORD, 1 or 2 override device mode
    "NpuContextSwitchQuantum" = 30000; // REG_DWORD, found in 25H2 (not in 23H2)
    "NpuPreemptionQuantum" = 60000; // REG_DWORD, found in 25H2 (not in 23H2)
    "HwSchThreadOffloadMode" = 2; // REG_DWORD, found in 25H2 (not in 23H2)
    "DebugLargeSmoothenedDuration" = 1; // REG_DWORD, found in 25H2 (not in 23H2)
    "AudioDgAutoBoostPriority" = 24; // REG_DWORD, found in 25H2 (not in 23H2)
    "FrameServerAutoBoostPriority" = 17; // REG_DWORD, found in 25H2 (not in 23H2)

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\MemoryManager";
    // ReadConfiguration
    "PinnedMemoryLimit" = 25; // REG_DWORD, >= 0x5A -> default 25
    "PinnedApertureMemoryLimit" = 40; // REG_DWORD, >= 0x5A -> default 40
    "PagesHistory" = 0; // REG_DWORD, max 0x7FFFFFF
    "MemTransferThreshold" = 10; // REG_DWORD
    "ExcessiveMemTransferFlipThreshold" = 15; // REG_DWORD
    "ExcessiveMemTransferPenalty" = 5; // REG_DWORD
    "EventThrottleThreshold" = 300; // REG_DWORD
    "DisablePrefetching" = 0; // REG_DWORD
    "NbDmaBufferLimitPerDevice" = 256; // REG_DWORD, 1024 if system memory > 0x20000000
    "NbCddDmaBufferLimitPerDevice" = 4; // REG_DWORD
    "DmaBufferBytesLimitAllDevices" = 0x800000; // REG_DWORD, 0x2000000 if system memory > 0x20000000
    "DmaBufferListBytesLimitAllDevices" = 0x400000; // REG_DWORD, 0x1000000 if system memory > 0x20000000
    "NbDmaBufferLimitCompareWatermark" = 10; // REG_DWORD
    "NbPagingHistoryRecords" = 0; // REG_DWORD, 0x40 if internal release
    "PinDWMAllocationBackingStore" = 1; // REG_DWORD
    "RemovePagesFromWorkingSetOnPagingForDwm" = 1; // REG_DWORD
    "UseUnreset" = 1; // REG_DWORD
    "PrivateHeapPackingThreshold" = 0x100000; // REG_DWORD
    "PrivateHeapPackingBlockSize" = 0x800000; // REG_DWORD
    "EvictTemporaryPeriod" = 60; // REG_DWORD
    "EvictUnusedPeriod" = 60; // REG_DWORD
    "ProcessPendingOfferPeriod" = 1; // REG_DWORD
    "ProcessSysmemOfferPeriod" = 8; // REG_DWORD
    "SegmentBalancingPolicy" = 2; // REG_DWORD
    "BugcheckOnApertureCorruption" = 0; // REG_DWORD
    "QuickApertureCorruptionCheck" = 0; // REG_DWORD
    "DirectFlipMemoryRequirement" = 200; // REG_DWORD
    "CommitProcessHeapOnDemand" = 1; // REG_DWORD
    "SegmentCleanupSizeThreshold" = 4096; // REG_DWORD
    "SegmentCleanupCountThreshold" = 6; // REG_DWORD
    "SegmentCleanupTime" = 20; // REG_DWORD
    "SelfRefreshVramForceEvictionTimer" = 900; // REG_DWORD
    "MaxSegmentSize<0..31>" = 0; // REG_DWORD, if set, aligns to 4K and clamps to 0x800000

    // unsure about the decomp defaults here
    "PhysicalHeapSize" = 16; // REG_QWORD
    "PhysicalHeapLowestAddress" = 16; // REG_QWORD
    "PhysicalHeapHighestAddress" = 4294967295; // REG_QWORD

    // ReadCommitLimitInformation
    "PinnedBackingStoreLimit" = 0; // REG_DWORD, MB (<< 20), 0 -> system memory / 8
    "MinimumSystemMemoryCommitLimit" = 0; // REG_DWORD, MB (<< 20), min 0x4000000
    "SmallSystemMemorySize" = 0; // REG_DWORD, MB (<< 20)
    "SystemPartitionCommitLimitPercentage" = 50; // REG_DWORD, clamped to 5..100
    "SecondaryPartitionCommitLimitPercentage" = 80; // REG_DWORD, clamped to 5..100

    // ReadWorkingSetConfiguration
    "WorkingSet.DefaultMaximumPercentile" = 90; // REG_DWORD
    "WorkingSet.DefaultMinimumPercentile" = 65; // REG_DWORD

    // ReadUnusedAllocationConfiguration
    "UnusedTrimmingPeriod" = 1; // REG_DWORD
    "Unused.MinimumThreshold" = 0; // REG_DWORD
    "Unused.LowThreshold" = 15; // REG_DWORD
    "Unused.NormalThreshold" = 45; // REG_DWORD
    "Unused.HighThreshold" = 120; // REG_DWORD
    "Unused.MaximumThreshold" = 1000000; // REG_DWORD
    "Unused.SelfTrimMinimumThreshold" = 0; // REG_DWORD
    "Unused.SelfTrimLowThreshold" = 1; // REG_DWORD
    "Unused.SelfTrimNormalThreshold" = 2; // REG_DWORD
    "Unused.SelfTrimHighThreshold" = 5; // REG_DWORD
    "Unused.SelfTrimMaximumThreshold" = 1000000; // REG_DWORD
    "Unused.EvictApertureOfferLowThreshold" = 15; // REG_DWORD
    "Unused.EvictApertureOfferNormalThreshold" = 15; // REG_DWORD
    "Unused.EvictApertureOfferHighThreshold" = 30; // REG_DWORD
    "Unused.EvictApertureOfferMaximumThreshold" = 30; // REG_DWORD

    // ReadPreparationPeriodConfiguration
    "PreparationPeriod" = 1; // REG_DWORD, scaled to 100ns
    "Period.MinimumPolicyHeldPeriod" = 4; // REG_DWORD
    "Period.MaximumPolicyHeldPeriod" = 64; // REG_DWORD
    "Period.AlwaysForceMemReset" = 1; // REG_DWORD
    "Period.EvictionThresholdForMemReset" = 32; // REG_DWORD, post query << 20
    "Period.NbOfAllocationsThresholdToMRU" = 0x7FFFFFFF; // REG_DWORD

    // ReadHeapConfiguration
    "DebouncedPageManagement" = 1; // REG_DWORD
    "DebouncedUnlockAge" = 15; // REG_DWORD
    "DebouncedDecommitAge" = 15; // REG_DWORD
    "RecycleHeapPackingThreshold" = 4; // REG_DWORD
    "RecycleHeapPackingBlockSize" = 32; // REG_DWORD
    "RecycleHeapPTDBlockSize" = 1024; // REG_DWORD
    "ZeroedRecyclePages" = 1; // REG_DWORD
    "LeanRecycleHeapPackingThreshold" = 4; // REG_DWORD
    "LeanRecycleHeapPackingBlockSize" = 8; // REG_DWORD
    "LeanRecycleHeapPTDBlockSize" = 64; // REG_DWORD
    "MaximumDecommitDebounce" = 256; // REG_DWORD, 64 if system memory <= 0x53333333
    "MaximumUnlockDebounce" = 256; // REG_DWORD, 64 if system memory <= 0x53333333
    "RecycleHistory" = 0; // REG_DWORD
    "RecycleHistorySize" = 64; // REG_DWORD
    "ZeroPageLockThreshold" = 0x200000; // REG_DWORD, found in 25H2 (not in 23H2)

    // ReadPowerConfiguration
    "MemoryComponentActiveThreshold" = 300; // REG_DWORD, MB (<< 20)
    "SelfRefreshMemoryEvictionThreshold" = 300; // REG_DWORD, MB (<< 20)

    // ReadGpuVaConfiguration
    "DisableUncommitGpuVaInPagingProcess" = 0; // REG_DWORD
    "EnableZeroFlagInPde" = 0; // REG_DWORD
    "DisableMakeIoMmuAddressValid" = 0; // REG_DWORD
    "PagingProcessVaSpaceBitCount" = 30; // REG_DWORD
    "GpuVaFirstValidAddress" = 0x10000; // REG_DWORD, masked to 4K
    "EnableGpuVaGuardPages" = 0; // REG_DWORD
    "AllocateGpuVaFromHighAddresses" = 0; // REG_DWORD
    "CompanionContextMaxPendingOperations" = 128; // REG_DWORD, found in 25H2 (not in 23H2)

    // ReadGpuVaPagingHistoryConfiguration
    "GpuVaPagingHistorySize" = ; // REG_DWORD, default 0x40 if system memory > 0x53333333 else 0
    "GpuVaPagingHistoryMask" = 391174; // REG_DWORD, derived, min 0x1000

    // ReadPagingConfiguration
    "DemotionWithinDeviceEnabled" = 1; // REG_DWORD
    "DeviceSuspendPeriodMin" = 500; // REG_DWORD
    "DeviceSuspendPeriodMax" = 500; // REG_DWORD
    "DeviceResumePeriodMin" = 1000; // REG_DWORD
    "DeviceResumePeriodMax" = 1000; // REG_DWORD
    "PagingQueueProcessingPeriodTime" = 50; // REG_DWORD, clamped 16..300
    "EnablePromotion" = 1; // REG_DWORD, found in 25H2 (not in 23H2)
    "InitialPromotionInterval" = 48; // REG_DWORD
    "MaximumPromotionInterval" = 5000; // REG_DWORD
    "PromotionTargetSizePerInterval" = 0x2000000; // REG_QWORD
    "PromotionNumberCapPerInterval" = 50; // REG_DWORD
    "TransferFlushThreshold" = 1; // REG_DWORD, MB (<< 20)
    "EnableAsyncResidency" = 1; // REG_DWORD
    "ForceUncommitGpuVAOnEvict" = 0; // REG_DWORD
    "ForceSynchronousEvict" = 0; // REG_DWORD
    "BreakOnPagingFailure" = 0; // REG_DWORD
    "TemporaryResourcePolicy" = 0; // REG_DWORD, found in 25H2 (not in 23H2)

    // ReadTestAndStagingConfiguration
    "BudgetThreshold" = 25; // REG_DWORD, clamped to <= 100
    "PagingQueueFenceIncrement" = 1; // REG_DWORD, 0 = 1, upper bound 0x51EB851
    "RestrictToPreferredSegment" = 0; // REG_DWORD
    "Use64KPages" = 0; // REG_DWORD
    "ExpandTo64KBAllocationSizeThreshold" = 0x400000; // REG_DWORD
    "AlwaysDecommitOnOffer" = 0; // REG_DWORD
    "LazyDecommitChunkSizeMB" = 32; // REG_DWORD, max 512
    "DecommitRepurposeMode" = 1; // REG_DWORD, values 0..2 else 0, found in 23H2 (not in 25H2)
    "DxgMms2OfferReclaim" = 4294967295; // REG_DWORD, allowed 0/1/2/4294967295, others = 0
    "LargifyUpgradeThresholdPercent" = 0; // REG_DWORD, found in 25H2 (not in 23H2)
    "LargifyUpgradeThresholdBytes" = 0; // REG_DWORD, found in 25H2 (not in 23H2)

    // ReadVPRConfiguration
    "VPRGrowRatioNumerator" = 4; // REG_DWORD
    "VPRGrowRatioDenominator" = 5; // REG_DWORD, if denominator <= numerator or numerator == 0 -> 4/5
    "VPRCapacityRatioNumerator" = 1; // REG_DWORD
    "VPRCapacityRatioDenominator" = 5; // REG_DWORD, if denominator <= numerator or numerator == 0 -> 4/5

    // ReadBudgetConfiguration
    "GlobalCommitmentBudget" = 0; // REG_QWORD
    "EnableTrimWnfCallback" = 1; // REG_DWORD
    "StartPeriodicTrimThreshold" = 40; // REG_DWORD
    "CriticalPeriodicTrimThreshold" = 10; // REG_DWORD
    "IdleTrimInterval" = 90000; // REG_DWORD
    "ForegroundTrimInterval" = 90000; // REG_DWORD
    "MaximumTrimInterval" = 10000; // REG_DWORD
    "MinimumTrimInterval" = 2000; // REG_DWORD
    "VideoMemoryFragmentationBuffer" = 10; // REG_DWORD
    "SystemMemoryFragmentationBuffer" = 5; // REG_DWORD
    "ProcessBudgetCapBuffer" = 5; // REG_DWORD
    "MaxVideoMemoryFragmentationBuffer" = 512; // REG_DWORD
    "MaxProcessBudgetCapBuffer" = 256; // REG_DWORD
    "L_LocalMemoryBudgetDWMTarget" = 30; // REG_DWORD
    "L_LocalMemoryBudgetFocusTarget" = 50; // REG_DWORD
    "LNL_LocalMemoryBudgetDWMTarget" = 30; // REG_DWORD
    "LNL_LocalMemoryBudgetFocusTarget" = 50; // REG_DWORD
    "LNL_NonLocalMemoryBudgetDWMTarget" = 30; // REG_DWORD
    "LNL_NonLocalMemoryBudgetFocusTarget" = 50; // REG_DWORD

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Power";
    "UseSelfRefreshVRAMInS3"; v166 = 1;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\BasicDisplay";
    "BasicDisplayUserNotified"; v2 = 0;

    "DisableBasicDisplayFallback"; v33 = -1;
    "EnableBasicDisplayFallback"; v32 = -1;
    "ForcePreserveBootDisplay"; v34 = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Smm";
    "DebugMode"; v11 = 0;
    "EnablePageTracking"; v8 = 0;
    "ForceDmaRemapping"; v9 = 0;
    "ForceEnableIommu"; v3 = 0;
    "IdentityMappedPassthrough"; v7 = 0;
    "LogicalAddressMode"; v4 = 0;
    "PreferHighLogicalAddresses"; v10 = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\DMM";
    "AssertOnDdiViolation"; g_DmmAssertOnDdiViolation = 0;
    "BadMonitorModeDiag"; v17 = 2;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\DMM";
    "EnableVirtualRefreshRateOnExternalMonitor"; *((_DWORD*)this + 134) = 0;
    "HPDFilterLimit"; *((_DWORD*)this + 133) = 20000000;
    "LongLinkTrainingTimeout"; *((_DWORD*)this + 132) = 1000;
    "ModeListCaching"; v81 = 1;
    "SetTimingsFlags"; *((_DWORD*)this + 130) = 0;
    "ShortLinkTrainingTimeout"; *((_DWORD*)this + 131) = 200;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Validation";
    "FailEscapeDDI"; v8 = 0
    "FailRenderDDI"; v9 = 0
    "FailReserveGPUVA"; v10 = 0
    "Level"; v7 = 0
    "ReportVirtualMachine"; v11 = 0

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\MonitorDataStore\\MONITOR-ID"
    "AdvancedColorEnabled"; v3 = 0;
    "AutoColorManagementEnabled"; v8 = 0;
    "AutoColorManagementSupported"; = ? // REG_DWORD, bool?
    "DockedOrientation" = ?;
    "EnableBoostRefreshRateByDefault" = ?;
    "EnableIntegratedPanelAcmByDefault"; v6 = 0;
    "EnableIntegratedPanelHdrByDefault"; v4 = 0;
    "HDREnabled"; v2 = 0;
    "MicrosoftApprovedAcmSupport"; v5 = 0;
    "MonitorOrientation" = ?;
    "OverrideWCGCapabilities" = ?;
    "PreferredScaleFactor" = ?;
    "SDRWhiteLevel" = ?;
    "VMSDisabled" = ?;

// the 3 keys below are based on a testing system monitor, therefore the defaults will be different for you
"HKLM\\System\\CurrentControlSet\\Control\\GraphicsDrivers\\Configuration\\<CONFIG_ID>";
    "SetId" = "MSBDD_MSI3CB01222_2E_07E4_FF_10DE_2482_00000007_00000000_0"; // REG_SZ
    "Timestamp" = = ?; // REG_QWORD

"HKLM\\System\\CurrentControlSet\\Control\\GraphicsDrivers\\Configuration\\<CONFIG_ID>\\00\\00";
    "ActiveSize.cx" = 2560; // REG_DWORD, horizontal pixels
    "ActiveSize.cy" = 1440; // REG_DWORD, vertical lines
    "BoostRefreshRateMultiplier" = 1; // REG_DWORD
    "ColorBasis" = 2; // REG_DWORD
    "DwmClipBox.bottom" = 1440; // REG_DWORD
    "DwmClipBox.left" = 0; // REG_DWORD
    "DwmClipBox.right" = 2560; // REG_DWORD
    "DwmClipBox.top" = 0; // REG_DWORD
    "Flags" = 25365391; // REG_DWORD
    "HSyncFreq.Denominator" = 2640; // REG_DWORD
    "HSyncFreq.Numerator" = 640000000; // REG_DWORD
    "PixelFormat" = 21; // REG_DWORD
    "PixelRate" = 640000000; // REG_DWORD
    "PrimSurfSize.cx" = 2560; // REG_DWORD
    "PrimSurfSize.cy" = 1440; // REG_DWORD
    "Rotation" = 1; // REG_DWORD
    "Scaling" = 4; // REG_DWORD
    "ScanlineOrdering" = 1; // REG_DWORD
    "Stride" = 10240; // REG_DWORD
    "VSyncFreq.Denominator" = 1000; // REG_DWORD
    "VSyncFreq.Numerator" = 164802; // REG_DWORD, refresh rate
    "VideoStandard" = 255; // REG_DWORD
    "VirtualRefreshRate.Denominator" = 1000; // REG_DWORD
    "VirtualRefreshRate.Numerator" = 164802; // REG_DWORD

"HKLM\\System\\CurrentControlSet\\Control\\GraphicsDrivers\\Configuration\\<CONFIG_ID>\\00";
    "CcdDbVersion" = 4; // REG_DWORD
    "ColorBasis" = 2; // REG_DWORD
    "PixelFormat" = 21; // REG_DWORD
    "Position.cx" = 0; // REG_DWORD
    "Position.cy" = 0; // REG_DWORD
    "PrimSurfSize.cx" = 2560; // REG_DWORD
    "PrimSurfSize.cy" = 1440; // REG_DWORD
    "Stride" = 10240; // REG_DWORD

"AdapterPnpKey";
    "EnableVirtualTopologySupport"; v84 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : EnableVirtualTopologySupport
    "NeedToSuspendVidSchBeforeSetGammaRamp"; v83 = (AdapterBuild < 8704 ? 1 : 0)
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : NeedToSuspendVidSchBeforeSetGammaRamp
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : NeedToSuspendVidSchBeforeSetGammaRamp

    "DisableNonPOSTDevice"; v40 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : DisableNonPOSTDevice
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicRender : DisableNonPOSTDevice

    "ACGSupported"; v165 = 0
    // Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : ACGSupported
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicRender : ACGSupported
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : ACGSupported
    "DxgkGpuVaIommuRequired"; v166 = 0
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : DxgkGpuVaIommuRequired
    "DxgkGpuVaIommuGlobalSupported"; v167 = 0
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : DxgkGpuVaIommuGlobalSupported

    "AllowUnspecifiedVSync"; v18 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : AllowUnspecifiedHSync
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : AllowUnspecifiedHSync
    "AllowUnspecifiedHSync"; v19 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : AllowUnspecifiedHSync
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : AllowUnspecifiedHSync
    "AllowUnspecifiedPixelRate"; v20 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : AllowUnspecifiedPixelRate
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : AllowUnspecifiedPixelRate
    "ForceDualViewBehavior"; v21 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : ForceDualViewBehavior
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : ForceDualViewBehavior

"<AdapterPnpKey>\\DxgkSettings";
    "UseSelfRefreshVRAMInS3"; v166 = 1;

// these are also in dxgmms2 but read from the pnp key, not from GraphicsDrivers - see https://github.com/nohuto/win-config/blob/0cbc8e153f7ea6bf4b640e51c53d235a5de67de8/power/desc.md#disable-device-powersavings
"<AdapterPnpKey>\\MemoryManager";
    "MaxLocalSegmentSize" = 0; // REG_DWORD, MB (<< 20), 0 allowed, 1..256 -> 256
    "MaxNonLocalSegmentSize" = 0; // REG_DWORD, MB (<< 20), 0 allowed, 1..512 -> 512
    "SelfRefreshVramForceEvictionTimerDC" = 900; // REG_DWORD, found in 25H2 (not in 23H2)
    "SelfRefreshVramForceEvictionTimerAC" = 900; // REG_DWORD, found in 25H2 (not in 23H2)
    "Supports64KBPages" = 0; // REG_DWORD, bit0 used
    "EnablePromotion" = 1; // REG_DWORD, found in 25H2 (not in 23H2)

"<AdapterPnpKey>";
    "HwQueuedRenderPacketGroupLimitPerNode" = ; // REG_BINARY, DWORD array, big endian
```

---

## Session Manager Values

See [session-manager-symbols](https://github.com/nohuto/win-registry/blob/main/assets/session-manager-symbols.txt) for reference.

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Kernel";
    "AdjustDpcThreshold" = 20; // KiAdjustDpcThreshold
    "AlwaysTrackIoBoosting" = 0; // PspAlwaysTrackIoBoosting
    "AmdTprLowerInterruptDelayConfig" = 0; // KiAmdTprLowerInterruptDelayConfig
    "BoostingPeriodMultiplier" = 3; // KiNormalPriorityBoostingPeriodMultiplier
    "BugCheckUnexpectedInterrupts" = 0; // KiBugCheckUnexpectedInterrupts
    "CacheAwareScheduling" = 47; // KiCacheAwareScheduling
    "CacheErrataOverride" = 0; // KiTLBCOverride
    "CacheIsoBitmap" = 0; // KiCacheIsoBitmap
    "DebuggerIsStallOwner" = 0; // KiDebuggerIsStallOwner
    "DebugPollInterval" = 2000; // KiDebugPollInterval
    "DefaultDynamicHeteroCpuPolicy" = 3; // (policy enum only)
    // Behavior of Dynamic hetero policy All (0) (all available) Large (1) LargeOrIdle (2) Small (3) SmallOrIdle (4) Dynamic (5) (use priority and other metrics to decide) BiasedSmall (6) (use priority and other metrics, but prefer small) BiasedLarge (7).
    "DefaultHeteroCpuPolicy" = 5; // KiDefaultHeteroCpuPolicy
    "DeviceOwnerProtectionDowngradeAllowed" = 0; // SeDeviceOwnerProtectionDowngradeAllowed
    "DisableControlFlowGuardExportSuppression" = 0; // PspDisableControlFlowGuardExportSuppression
    "DisableExceptionChainValidation" = 2; // PspSehValidationPolicy
    "DisableLightWeightSuspend" = 0; // KiDisableLightWeightSuspend
    "DisableLowQosTimerResolution" = 1; // KeDisableLowQosTimerResolution
    "DisablePointerParameterAlignmentValidation" = 0; // KiDisablePointerParameterAlignmentValidation
    "DisableTsx" = 0; // KiDisableTsx
    "DpcCumulativeSoftTimeout" = 120000; // KeDpcCumulativeSoftTimeoutMs
    "DpcQueueDepth" = 4; // KiMaximumDpcQueueDepth
    "DpcSoftTimeout" = 20000; // KeDpcSoftTimeoutMs
    "DPCTimeout" = 20000; // KeDpcTimeoutMs
    "DpcWatchdogPeriod" = 120000; // KeDpcWatchdogPeriodMs
    "DpcWatchdogProfileBufferSizeBytes" = 266240; // KeDpcWatchdogProfileBufferSizeBytes
    "DpcWatchdogProfileCumulativeDpcThreshold" = 110000; // KeDpcWatchdogProfileCumulativeDpcThresholdMs
    "DpcWatchdogProfileOffset" = 10000; // KeDpcWatchdogProfileOffsetMs
    "DpcWatchdogProfileSingleDpcThreshold" = 18333; // KeDpcWatchdogProfileSingleDpcThresholdMs
    "DriveRemappingMitigation" = 1; // ObpDriveRemappingMitigation
    "DynamicHeteroCpuPolicyExpectedRuntime" = 5200; // KiDynamicHeteroCpuPolicyExpectedRuntime
    "DynamicHeteroCpuPolicyImportant" = 2; // (LargeOrIdle)
    // Policy for a dynamic thread that is deemed important.
    "DynamicHeteroCpuPolicyImportantPriority" = 8; // KiDynamicHeteroCpuPolicyImportantPriority
    // Priority above which threads are considered important if prioritybased dynamic policy is chosen.
    "DynamicHeteroCpuPolicyImportantShort" = 3; // (Small)
    // Policy for dynamic thread that is deemed important but run a short amount of time.
    "DynamicHeteroCpuPolicyMask" = 7; //  (foreground status = 1, priority = 2, expected run time = 4)
    // Determine what is considered in assessing whether a thread is important.
    "EnablePerCpuClockTickScheduling" = 0; // KiEnableClockTimerPerCpuTickScheduling
    "EnableTickAccumulationFromAccountingPeriods" = 0; // KiEnableTickAccumulationFromAccountingPeriods
    "EnableWerUserReporting" = 1; // DbgkEnableWerUserReporting
    "ForceBugcheckForDpcWatchdog" = 0; // KiForceBugcheckForDpcWatchdog
    "ForceForegroundBoostDecay" = 0; // KiSchedulerForegroundBoostDecayPolicy
    "ForceIdleGracePeriod" = 5; // KiForceIdleGracePeriodInSec
    "ForceParkingRequested" = 1; // KiForceParkingConfiguration
    "GlobalTimerResolutionRequests" = 0; // KiGlobalTimerResolutionRequests
    "HeteroFavoredCoreFallback" = 0; // PpmHeteroFavoredCoreFallback
    "HeteroSchedulerOptions" = 0; // KiHeteroSchedulerOptions
    "HeteroSchedulerOptionsMask" = 0; // KiHeteroSchedulerOptionsMask
    "HgsPlusFeedbackUpdateThresholdNetRuntime" = 20; // dword_140FC33C0
    "HgsPlusFeedbackUpdateThresholdRuntime" = 20; // dword_140FC33B4
    "HgsPlusHigherPerfClassFeedbackThreshold" = 1; // dword_140FC33E0
    "HgsPlusInvalidFeedbackDefaultClass" = 0; // dword_140FC33D4
    "HgsPlusInvalidFeedbackDefaultClassSet" = 0; // dword_140FC33D8
    "HgsPlusInvalidFeedbackLimit" = 50; // dword_140FC33D0
    "HgsPlusLowerPerfClassFeedbackThreshold" = 4; // dword_140FC33DC
    "HgsPlusMinimumScoreDifferenceForSwap" = 25; // dword_140FC33E8
    "HgsPlusThreadCreationDefaultClass" = 0; // dword_140FC33E4
    "HotpatchTestMode" = 0; // KeHotpatchTestMode
    "HyperStartDisabled" = 0; // HvlVpStartDisabled
    "IdealDpcRate" = 20; // KiIdealDpcRate
    "IdealNodeRandomized" = 1; // PspIdealNodeRandomized
    "InterruptSteeringFlags" = 0; // KiInterruptSteeringFlags
    "LongDpcQueueThreshold" = 3; // KiLongDpcQueueThreshold
    "LongDpcRuntimeThreshold" = 100; // KiLongDpcRuntimeThreshold
    "MaxDynamicTickDuration" = 8; // KiMaxDynamicTickDurationSize
    "MaximumCooperativeIdleSearchWidth" = 16; // KiMaximumCooperativeIdleSearchWidth
    "MaximumSharedReadyQueueSize" = 260; // KiMaximumSharedReadyQueueSize
    "MinimumDpcRate" = 3; // KiMinimumDpcRate
    "MitigationAuditOptions" = 0; // PspSystemMitigationAuditOptions
    "MitigationOptions" = 0; // PspSystemMitigationOptions
    "ObCaseInsensitive" = 1; // ObpCaseInsensitive
    "ObObjectSecurityInheritance" = 0; // ObpObjectSecurityInheritance
    "ObTracePermanent" = 0; // ObpTracePermanent
    "ObTracePoolTags" = 0; // ObpTracePoolTagsBuffer / ObpTracePoolTagsLength
    "ObTraceProcessName" = 0; // ObpTraceProcessNameBuffer / ObpTraceProcessNameLength
    "ObUnsecureGlobalNames" = 6619246; // ObpUnsecureGlobalNamesBuffer / ObpUnsecureGlobalNamesLength
    "PassiveWatchdogTimeout" = 300; // KiPassiveWatchdogTimeout
    "PerfIsoEnabled" = 0; // KiPerfIsoEnabled
    "PoCleanShutdownFlags" = 0; // PopShutdownCleanly
    "PowerOffFrozenProcessors" = 1; // KiPowerOffFrozenProcessors
    "ReadyTimeTicks" = 6; // KiNormalPriorityBoostReadyTimeTicks
    "RebalanceMinPriority" = 1; // KiRebalanceMinPriority
    "ReservedCpuSets" = 0; // KiReservedCpuSets
    "ScanLatencyTicks" = 7; // KiNormalPriorityBoostScanLatencyTicks
    "SchedulerAssistThreadFlagOverride" = 0; // KiSchedulerAssistThreadFlagOverride
    "SeAllowAllApplicationAceRemoval" = 0; // SepAllowAllApplicationAceRemoval
    "SeAllowSessionImpersonationCapability" = 0; // SepAllowSessionImpersonationCap
    "SeCompatFlags" = 0; // SeCompatFlags
    "SeLpacEnableWatsonReporting" = 0; // SeLpacEnableWatsonReporting
    "SeLpacEnableWatsonThrottling" = 1; // SeLpacEnableWatsonThrottling
    "SerializeTimerExpiration" = 1; // KiSerializeTimerExpiration
    // This behavior is controlled by the kernel variable KiSerializeTimerExpiration, which is initialized based on a registry setting whose value is different between a server and client installation. By modifying or creating the value SerializeTimerExpiration under HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel other than 0 or 1, serialization can be disabled, enabling timers to be distributed among processors. Deleting the value, or keeping it as 0, allows the kernel to make the decision based on Modern Standby availability, and setting it to 1 permanently enables serialization even on non-Modern Standby systems.
    "SeTokenDoesNotTrackSessionObject" = 0; // SeTokenDoesNotTrackSessionObject
    "SeTokenLeakDiag" = 0; // SeTokenLeakTracking
    "SeTokenSingletonAttributesConfig" = 3; // SepTokenSingletonAttributesConfig
    "SplitLargeCaches" = 0; // KiSplitLargeCaches
    "ThreadDpcEnable" = 1; // KeThreadDpcEnable
    "ThreadReadyCount" = 1; // KiNormalPriorityBoostMaximumThreadReadyCount
    "TimerCheckFlags" = 1; // KeTimerCheckFlags
    "VerifierDpcScalingFactor" = 1; // KeVerifierDpcScalingFactor
    "VirtualHeteroHysteresis" = 4294967295; // PpmPerfQosTransitionHysteresisOverride
    "VpThreadSystemWorkPriority" = 30; // KiVpThreadSystemWorkPriority
    "WpsSimulationOverride" = 0; // PpmWpsSimulationOverride / PpmWpsSimulationOverrideSize
    "XStateContextLookasidePerProcMaxDepth" = 0; // KiXStateContextLookasidePerProcMaxDepth

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Kernel\\RNG";
    "RNGAuxiliarySeed" = ; // ExpRNGAuxiliarySeed - REG_DWORD, default of 1807947291? ("HKLM\System\CurrentControlSet\Control\Session Manager\kernel\RNG\RNGAuxiliarySeed","Type: REG_DWORD, Data: 1807947291", procmon boot trace)

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager";
    "AlpcMessageLog" = 0; // AlpcpMessageLogEnabled 
    "AlpcWakePolicy" = 1; // AlpcpWakePolicyDefault 
    "CriticalSectionTimeout" = 2592000; // dword_140FC3204 dd 278D00
    "CWDIllegalInDLLSearch" = 0; // PspCurDirDevicesSkippedForDlls 
    "Debugger Retries" = 20; // KdpContext (0x14) 
    "DisableIFEOCaching" = 0; // RtlpDisableIFEOCaching 
    "GlobalFlag" = 0; // CmNtGlobalFlag <> 0x7061006c ?
    "GlobalFlag2" = 0; // CmNtGlobalFlag2 <> 0x6c642e30 ?
    "HeapDeCommitFreeBlockThreshold" = 4096; // qword_140FC3210 dq 1000
    "HeapDeCommitTotalFreeThreshold" = 65536; // qword_140FC3218 dq 10000
    "HeapSegmentCommit" = 8192; // qword_140FC3220 dq 2000
    "HeapSegmentReserve" = 1048576; // qword_140FC3228 dq 100000
    "ImageExecutionOptions" = 0; // ViImageExecutionOptions 
    "InitConsoleFlags" = 0; // InitConsoleFlags 
    "MultiUsersInSessionSupported" = 0; // RtlpMultiUsersInSessionSupported 
    "ObjectSecurityMode" = 1; // ObpObjectSecurityMode 
    "PowerPolicySimulate" = 0; // PopSimulate 
    "ProtectionMode" = 1; // ObpProtectionMode , DWORD
    "ResourceCheckFlags" = 3; // ExResourceCheckFlags 
    "ResourceEnforceOwnerTransfer" = 0; // ExpResourceEnforceOwnerTransfer 
    "ResourceTimeoutCount" = 45; // ExResourceTimeoutCount (0x2d) 
    "SkipRegistryInit" = 0; // CmNtSkipRegistryInit 

    // procmon boot trace
    "ObjectDirectories" = \Windows, \RPC Control; // ? - REG_MULTI_SZ
    "BootExecute" = ?; // REG_SZ
    "BootExecuteNoPnpSync" = ?;
    "PlatformExecute" = ?;
    "SetupExecute" = ?;
    "SetupExecuteNoPnpSync" = ?;
    "S0InitialCommand" = ?;
    "NumberOfInitialSessions" = 2; // ? - REG_DWORD
    "PendingFileRenameOperations" = ?;
    "PendingFileRenameOperations2" = ?;
    "AllowProtectedRenames" = ?;
    "ClearTempFiles" = ?;
    "TempFileDirectory" = ?;
    "ExcludeFromKnownDlls" = ?; // REG_MULTI_SZ
    "BackgroundLoadKnownDlls" = ?;
    "DisableWpbtExecution" = ?; // REG_DWORD
    "RaiseExceptionOnPossibleDeadlock" = ?;
    "ResourcePolicies" = ?;
    "SafeDllSearchMode" = ?;
    "SafeProcessSearchMode" = ?;
    "SmtDelayBaseYield" = ?;
    "SmtDelayMaxYield" = ?;
    "SmtDelaySleepLoopWindowSize" = ?;
    "SmtDelaySpinCountThreshold" = ?;
    "SmtFactorYield" = ?;
    "SystemUpdateOnBoot" = ?;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Quota System";
    "ApplicationBlockedMessageLimit" = 50; // PspJobNoWakeChargeLimit (0x32) 
    "JobTimeLimitsPeriodSeconds" = 7; // PspJobTimeLimitsPeriodSeconds 
    "SystemBlockedMessageLimit" = 200; // PspSystemNoWakeChargeLimit (0xC8) 

    "DfssGenerationLengthMS" = 600; // PsDfssGenerationLengthMS dd 258
    "DfssLongTermFraction1024" = 512; // sDfssLongTermFraction1024 dd 200
    "DfssLongTermSharingMS" = 15; // PsDfssLongTermSharingMS dd 0F
    "DfssResolutionMS" = 4294967295; // PsDfssDesiredTimerResolutionMs dd 0FFFFFFFF
    "DfssShortTermSharingMS" = 30; // PsDfssShortTermSharingMS dd 1E
    "EnableCpuQuota" = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management";
    "AllocationPreference" = 0; // dword_140FC3200 dd 0
    "AllowUserHotPatchWithoutVbs" = 0; // dword_140FC3250 dd 0
    "CacheUnmapBehindLengthInMB" = 8388608; // CcUnmapBehindLength (0x00800000) 
    "CustomDTPDenominator" = 8; // CcClientDTPDenominator (0x8) 
    "DeadlockRecursionDepthLimit" = 0; // ViRecursionDepthLimitFromRegistry 
    "DeadlockSearchNodesLimit" = 0; // ViSearchedNodesLimitFromRegistry 
    "DifPluginConfigData" = 635710207; // DifPluginConfigData (0x25e8007f) 
    "DifPluginConfigDataLength" = 1276097421; // DifPluginConfigDataLength (0x4c084b8d) 
    "DisableCacheTelemetry" = 2; // CcDisableTelemetryRegKeyAtInit 
    "DisablePageCombining" = 0; // dword_140FC31E8 dd 0
    "DisablePagingExecutive" = 0; // dword_140FC31E4 dd 0
    "EnableAsyncLazywrite" = 2; // CcEnableAsyncLazywriteOverride 
    "EnableAsyncLazywriteMulti" = 2; // CcEnableAsyncLazywriteMultiOverride 
    "EnableCooling" = 0; // dword_140FC31F8 dd 0
    "EnablePerVolumeLazyWriter" = 2; // CcEnablePerVolumeLazyWriterOverride 
    "ForceValidateIo" = 0; // dword_140FC31F0 dd 0
    "HighMemoryThreshold" = 0; // qword_140FC3238 dq 0
    "KernelPadSectionsOverride" = 0; // dword_140FC3248 dd 0
    "LargeWriteSize" = 0; // CcAzure_LargeWriteSize 
    "LazyWriterPercentageOfNumProcs" = 0; // CcAzure_LazyWriterPercentageOfNumProcs 
    "LowMemoryThreshold" = 0; // qword_140FC3230 dq 0
    "MaxLazyWritePages" = 0; // CcMaxLazyWritePagesOverride 
    "MinimumStackCommitInBytes" = 0; // dword_140FC3208 dd 0
    "Mirroring" = 0; // dword_140FC31F4 dd 0
    "ModifiedWriteMaximum" = ?; // dword_140FC31FC
    "MoveImages" = 1; // MmRegistryState 
    "NonPagedPoolQuota" = 4294967295; // PspDefaultResourceLimits (4294967295) 
    "PagedPoolQuota" = ?; // unk_140FD7DE4
    "PageValidationAction" = 0; // MmPageValidationAction 
    "PageValidationFrequency" = 0; // MmPageValidationFrequency 
    "PagingFileQuota" = ?; // unk_140FD7DE8
    "PhysicalMemoryMapperEnforcementMode" = 0; // dword_140FC324C dd 0
    "PoolForceFullDecommit" = 0; // PoolForceFullDecommit 
    "PoolTag" = 0; // MmSpecialPoolTag 
    "PoolTagOverruns" = 1; // MmSpecialPoolCatchOverruns 
    "PoolTagSmallTableSize" = 4097; // PoolTrackTableSize (0x1001) 
    "ProtectNonPagedPool" = 0; // MmProtectFreedNonPagedPool 
    "RemoteFileDirtyPageThreshold" = 1310720; // CcRemoteFileDPInlineFlushThreshold (0x00140000) 
    "SimulateCommitSavings" = 0; // dword_140FC3240 dd 0
    "SoftThrottleDelayInMs" = 0; // CcAzure_SoftThrottleDelayInMs 
    "SoftThrottleLargeWriteAtPct" = 0; // CcAzure_SoftThrottleLargeWriteAtPct 
    "SpecialPurposeMemoryPages" = 0; // MmSpecialPurposeMemoryPages 
    "SpecialPurposeMemoryStartPage" = 0; // MmSpecialPurposeMemoryStartPage 
    "SpecialPurposeMemoryStartPageValueSize" = 4294967295; // MmSpecialPurposeMemoryStartPageValueSize (4294967295) 
    "TopBottomDPTEqual" = 0; // CcAzure_TopBottomDPTEqual 
    "TrackLockedPages" = 0; // MmTrackLockedPages 
    "TrackPtes" = 0; // dword_140FC31EC dd 0
    "VerifierDifPoolTags" = 0; // DifpPoolTags 
    "VerifierDifPoolTagsSizeBytes" = 4294967295; // DifpPoolTagsSizeBytes (4294967295) 
    "VerifierFaultApplications" = 0; // VerifierFaultApplicationsBuffer 
    "VerifierFaultApplicationsSize" = 4294967295; // VerifierFaultApplicationsBufferSize (4294967295) 
    "VerifierFaultBootMinutes" = 8; // VfFaultInjectionBootMinutes 
    "VerifierFaultProbability" = 600; // VfFaultInjectionProbability (0x258) 
    "VerifierFaultTags" = 0; // VerifierFaultTagsBuffer 
    "VerifierFaultTagsSize" = 4294967295; // VerifierFaultTagsBufferSize (4294967295) 
    "VerifierHandleTraces" = 16384; // VfHandleTracingEntries (0x4000) 
    "VerifierIrpStackTraces" = 16384; // IovIrpTracesLength (0x4000) 
    "VerifierIrpTimeout" = 0; // VfWdIrpTimeoutMsec 
    "VerifierNewRuleWorkaround" = 0; // VerifierNewRuleWorkaround 
    "VerifierOptions" = 0; // VfOptionFlags 
    "VerifierRandomTargets" = 0; // VfRandomVerifiedDrivers 
    "VerifierSettingState" = 0; // VfRuleClasses 
    "VerifierSettingStateSize" = 4294967295; // VfRuleClassesSize (4294967295) 
    "VerifierTipDisable" = 0; // VerifierTipDisable 
    "VerifierTipLimitDenominator" = 0; // DifiPluginControlDenominator 
    "VerifierTipLimitNumerator" = 0; // DifiPluginControlNumerator 
    "VerifierTipSparseness" = 0; // DifiPluginControlSparseness 
    "VerifierTriageContext" = 0; // VfTriageContext 
    "VerifyBTSBufferSize" = 0; // ViVerifyBTSBufferSize 
    "VerifyDriverLevel" = 4294967295; // MmVerifyDriverLevel (4294967295) 
    "VerifyDrivers" = 3905129288; // MmVerifyDriverBuffer (0xE8C38B48) 
    "VerifyDriversLength" = 1207968387; // MmVerifyDriverBufferLength (0x48002283) 
    "VerifyDriversSuppress" = 276138824; // VfXdvSuppressDriversBuffer (0x10758b48) 
    "VerifyDriversSuppressLength" = 3482011648; // VfXdvSuppressDriversBufferLength (0xCF8B4800) 
    "VerifyMode" = 4; // VfVerifyMode 
    "VerifyTriage" = 4294967295; // ViVerifyTriage (4294967295) 
    "VerifyTriageRules" = 0; // ViVerifyTriageRules 
    "VerifyTriageRulesSize" = 4294967295; // ViVerifyTriageRulesSize (4294967295) 
    "VmPauseOutswapSizeCapMB" = 512; // VmPauseOutswapSizeCapMB (0x200) 
    "WorkingSetPagesQuota" = ?; // unk_140FD7DEC
    "WorkingSetSwapSharedPages" = 0; // PspOutSwapSharedPages 
    "XdvTipTag" = 0; // CarTipTag 
    "XdvVerifierOptions" = 0; // CarXdvOptions 
    "XdvVerifierOptions" = 0; // VfFlightOptions

    // procmon boot trace
    "PagingFiles" = C:\pagefile.sys <int> <int> // REG_MULTI_SZ
    "PagefileOnOsVolume" = ?; // 4,094
    "WaitForPagingFiles" = ?; // 4,094
    "ExistingPageFiles" = \??\C:\pagefile.sys; // REG_MULTI_SZ
    "DisableDedicatedMemoryCaching" = ?;
    "DedicatedMemoryPagefileSizeMB" = ?
    "PagefileHybridPriority" = ?;
    "SwapfileControl" = ?;
    "SwapFile" = ?;
    "TempPageFile" = ?;
    "FeatureSettings" = ? // DWORD

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Executive";
    "AdditionalCriticalWorkerThreads" = 0; // ExpAdditionalCriticalWorkerThreads 
    "AdditionalDelayedWorkerThreads" = 0; // ExpAdditionalDelayedWorkerThreads 
    "ForceEnableMutantAutoboost" = 0; // ExpForceEnableMutantAutoboost 
    "KernelWorkerTestFlags" = 0; // ExpWorkerQueueTestFlags 
    "MaximumKernelWorkerThreads" = 4096; // ExpMaximumKernelWorkerThreads (0x1000) 
    "MaxTimeSeparationBeforeCorrect" = 60; // ExpMaxTimeSeperationBeforeCorrect (0x3C) 
    "WorkerFactoryThreadCreationTimeout" = 10; // ExpWorkerFactoryThreadCreationTimeoutInSeconds (0x0A) 
    "WorkerFactoryThreadIdleTimeout" = 67; // ExpWorkerFactoryThreadIdleTimeoutInSeconds (0x43) 
    "WorkerThreadTimeoutInSeconds" = 600; // ExpWorkerThreadTimeoutInSeconds (0x258)
    "TickcountRolloverDelay" = 0; // ? (InitTickRolloverDelay dd 0) - InitTickRolloverDelay <> 24848b00, InitTickRolloverDelayLength <> 5e4130c4, InitTickRolloverDelayType <> e2894460

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Power";
    "FlushPolicy" = 0; // PopFlushPolicy 
    "IdleScanInterval" = 30; // PopIdleScanInterval (0x1E) 
    "SkipTickOverride" = 1; // PopSkipTickPolicy
    "SleepStudyDeviceAccountingLevel" = 4; // PopSleepStudyDeviceAccountingLevel 
    "SleepStudyDisabled" = 0; // PopSleepStudyDisabled 
    "WatchdogResumeTimeout" = 120; // PopWatchdogResumeTimeout (0x78) 
    "WatchdogSleepTimeout" = 300; // PopWatchdogSleepTimeout (0x12C) 
    "Win32CalloutWatchdogBugcheckEnabled" = 0; // PopWin32CalloutWatchdogBugcheckEnabled 

    // PopOpenPowerKey
    "AwayModeEnabled" = 0; // REG_DWORD, range: 0-1
    "HiberbootEnabled" = 0; // REG_DWORD, range: 0-1
    "KernelResumeIoCpuTime" = 0; // REG_DWORD, milliseconds, range: 0-4294967295
    "MaxHuffRatio" = 1; // REG_DWORD, range: 1-98
    "MultiPhaseResumeDisabled" = 0; // REG_DWORD, range: 0-1
    "SystemPowerPolicy" = "<STRUCT 232 BYTES>"; // REG_BINARY, Size=232

    // HybridBootAnimationTime records the boot animation duration during fast boot, HiberIoCpuTime is CPU time spent on hibernation I/O during resume, ResumeCompleteTimestamp is the system timestamp when resume from hibernation completed. So all of them are just counters and chaning their data won't affect the boot.
    "HiberIoCpuTime" = 0; // REG_DWORD, milliseconds, range: 0-4294967295
    "HybridBootAnimationTime" = 1601; // REG_DWORD, milliseconds, range: 0-4294967295
    "ResumeCompleteTimestamp" = 0; // REG_QWORD, range: 0-4294967295FFFFFFFF

    // PpmInitIllegalThrottleLogging
    "ProcessorThrottleLogInterval" = 10000; // REG_DWORD, milliseconds, range: 0-10000 (values >10000 are clamped to 10000)

    // procmon boot trace
    "SleepStudyBufferSizeInMB" = ?;
    "SleepStudyHistoryDays" = ?;
    "SleepStudyPerfTrackDripsThresholdPercentage" = ?;
    "SleepStudyTraceDirectory" = ?;

"HKLM\\System\\CurrentControlSet\\Control\\Session Manager\\Throttle";
    "PerfEnablePackageIdle" = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Segment Heap";
    "Enabled" = 0; // if present with DataLength==4 and nonzero type:
                    //    RtlpLowFragHeapGlobalFlags |= 0x10;  // global segment heap enable
                    //    if (value & 0x2)                      // low byte, bit 1
                    //        RtlpLowFragHeapGlobalFlags |= 0x20; // extra option ?
                    // if the value exists but is stored as REG_NONE (type==0):
                    //    RtlpLowFragHeapGlobalFlags |= 0x8;   // global "disable/override"

// Miscellaneous values

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\LSA";
    "AuditBaseDirectories" = 0; // ObpAuditBaseDirectories 
    "AuditBaseObjects" = 0; // ObpAuditBaseObjects 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\LSA\\audit";
    "ProcessAccessesToAudit" = 0; // SepProcessAccessesToAudit 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\TimeZoneInformation";
    "ActiveTimeBias" = ?; // dword_140FCE974
    "Bias" = 480; // ExpAltTimeZoneBias (0x000001e0) 
    "RealTimeIsUniversal" = 0; // ExpRealTimeIsUniversal 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\I/O System";
    "DisableDiskCounters" = 0; // PsDisableDiskCounters 
    "IoAllowLoadCrashDumpDriver" = 0; // IopAllowLoadCrashDumpDriver 
    "IoBlockLegacyFsFilters" = 0; // IopBlockLegacyFsFilters 
    "IoCaseInsensitive" = 1; // IopCaseInsensitive 
    "IoEnableSessionZeroAccessCheck" = 0; // IopSessionZeroAccessCheckEnabled 
    "IoFailZeroAccessCreate" = 1; // IopFailZeroAccessCreate 
    "IoIrpCompletionTimeoutInSeconds" = 300; // IopIrpCompletionTimeoutInSeconds (0x12C) 
    "IoKeepAliveTimeMs" = 5000; // IopKeepAliveTimeMs (0x1388) 
    "LargeIrpStackLocations" = 14; // IopLargeIrpStackLocations (0x0E) 
    "MediumIrpStackLocations" = 2; // IopMediumIrpStackLocations 
    "RequireDeviceAccessCheck" = 1; // IopRequireDeviceAccessCheck 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Configuration Manager";
    "BugcheckRecoveryEnabled" = 0; // CmBugcheckRecoveryEnabled 
    "CallbackMemoryFromPerProcLookaside" = 1; // CmpAllocateCallbackMemoryFromPerProcLookaside 
    "CallbackMemoryFromPool" = 0; // CmpAllocateCallbackMemoryFromPool 
    "DelayCloseSize" = 2048; // CmpDelayedCloseSize (0x800) 
    "Enabled" = 0; // CmpLKGEnabled 
    "EnablePeriodicBackup" = 0; // CmpDoIdleProcessing 
    "FastBoot" = 1; // CmFastBoot 
    "FreezeThawTimeoutInSeconds" = 60; // CmFreezeThawTimeoutInSeconds (0x3C) 
    "RegistryFlushGlobalFlags" = 0; // CmpGlobalFlushControlFlags 
    "RegistryLazyFlushBootDelay" = 60; // CmpEnableLazyFlushBootDelayInterval (0x3C) 
    "RegistryLazyFlushInterval" = 60; // CmpLazyFlushIntervalInSeconds (0x3C) 
    "RegistryLazyLocalizeInterval" = 60; // CmpLazyLocalizeIntervalInSeconds (0x3C) 
    "RegistryLazyReconcileInterval" = 3600; // CmpLazyReconcileIntervalInSeconds (0x0E10) 
    "RegistryLogFileSizeCap" = 0; // CmpLogFileSizeCap 
    "RegistryReorganizationLimit" = 1048576; // CmpReorganizeLimit (0x00100000) 
    "RegistryReorganizationLimitDays" = 7; // CmpReorganizeDelayDays 
    "SelfHealingEnabled" = 1; // CmSelfHeal 
    "SystemHiveLimitSize" = 1610612736; // CmSystemHiveLimitSize (0x60000000) 
    "VirtualizationEnabled" = 1; // CmVEEnabled 
    "VolatileBoot" = 0; // CmpVolatileBoot 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\StateSeparation\\Policy";
    "AllHivesVolatile" = 0; // CmStateSeparationAllHivesVolatile 
    "DevelopmentMode" = 0; // CmStateSeparationDevMode 
    "Enabled" = 0; // CmStateSeparationEnabled 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\ValidationRunlevels";
    "Global" = 1210938368; // CmGlobalValidationRunlevel (0x482d7400) 

"HKLM\\System\\CurrentControlSet\\Control\\Processor";
    "AllowGuestPerfStates" = 0;
    "AllowPepPerfStates" = 0;
    "Capabilities" = 4294967288; // Fallback of 0 ?
    "DisableAsserts" = 0;
    "Overrides" = 0;
```

## Power Values

See [power-symbols](https://github.com/nohuto/win-registry/blob/main/assets/power-symbols.txt) for reference. The list doesn't include all existing values yet, but the listed ones do exist. [PopOpenPowerKey-all](https://github.com/nohuto/win-registry/blob/main/assets/PopOpenPowerKey-all.c) shows the pseudocode for several `Session Manager\\Power` values.

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power";
    "ActiveIdleLevel" = 1; // PopFxActiveIdleLevel 
    "ActiveIdleThreshold" = 5000000; // PopFxActiveIdleThreshold (0x004C4B40) 
    "ActiveIdleTimeout" = 1000; // PopFxActiveIdleTimeout (0x000003E8) 
    "AllowAudioToEnableExecutionRequiredPowerRequests" = 1; // PopPowerRequestActiveAudioEnablesExecutionRequired 
    "AllowHibernate" = 4294967295; // PopAllowHibernateReg (4294967295) - REG_DWORD
    "AllowSystemRequiredPowerRequests" = 1; // PopPowerRequestConvertSystemToExecution 
    "AlwaysComputeQosHints" = 0; // PpmPerfAlwaysComputeQosEnabled 
    "BootHeteroPolicyOverride" = 0; // PpmPerfBootHeteroPolicyOverrideEnabled 
    "CheckpointSystemSleep" = 0; // PopCheckpointSystemSleepEnabledReg 
    "CheckpointSystemSleepSimulateFlags" = 0; // PopCheckpointSystemSleepSimulateFlags 
    "CheckPowerSourceAfterRtcWakeTime" = 30; // PopCheckPowerSourceAfterRtcWakeTime (0x1E) 
    "Class1InitialUnparkCount" = 64; // PpmParkInitialClass1UnParkCount (0x40) 
    "CoalescingFlushInterval" = 60; // PopCoalescingFlushInterval (0x0000003C) 
    "CoalescingTimerInterval" = 1500; // PopCoalescingTimerInterval (0x000005DC) - Units: seconds (multiplies value by -10,000,000, one second in 100 ns units, so the default corresponds to a 25min cadence)
    "DeepIoCoalescingEnabled" = 0; // PopDeepIoCoalescingEnabled 
    "DirectedDripsAction" = 3; // PopDirectedDripsAction 
    "DirectedDripsDebounceInterval" = 120; // PopDirectedDripsDebounceInterval (0x78) 
    "DirectedDripsDfxEnforcementPolicy" = 1; // PopDirectedDripsDfxEnforcementPolicy 
    "DirectedDripsOverride" = 4294967295; // PopDirectedDripsOverride (4294967295) 
    "DirectedDripsSurprisePowerOnTimeout" = 5; // PopDirectedDripsSurprisePowerOnTimeoutSeconds 
    "DirectedDripsTimeout" = 300; // PopDirectedDripsTimeout (0x12C) 
    "DirectedDripsWaitWakeTimeout" = 5; // PopDirectedDripsWaitWakeTimeoutSeconds 
    "DirectedFxDefaultTimeout" = 120; // PopFxDirectedFxDefaultTimeout (0x00000078) 
    "DisableDisplayBurstOnPowerSourceChange" = 0; // PopDisableDisplayBurstOnPowerSourceChange 
    "DisableIdleStatesAtBoot" = 0; // PpmIdleDisableStatesAtBoot 
    "DisableInboxPepGeneratedConstraints" = 4294967295; // PopDisableInboxPepGeneratedConstraintsOverride (4294967295) 
    "DisableVsyncLatencyUpdate" = 0; // PpmDisableVsyncLatencyUpdate 
    "DozeDeferralChecksToIgnore" = 0; // PopDozeDeferralChecksToIgnore 
    "DozeDeferralMaxSeconds" = 259200; // PopDozeDeferralMaxSeconds (0x0003F480) 
    "DripsCallbackInterval" = 35; // PopDripsCallbackInterval (0x23) 
    "DripsSwHwDivergenceEnableLiveDump" = 0; // PopDripsSwHwDivergenceEnableLiveDump 
    "DripsSwHwDivergenceThreshold" = 270; // PopDripsSwHwDivergenceThreshold (0x010E) 
    "DripsWatchdogAction" = 198; // PopDripsWatchdogAction (0xC6) 
    "DripsWatchdogDebounceInterval" = 120; // PopDripsWatchdogDebounceInterval (0x78) 
    "DripsWatchdogTimeout" = 300; // PopDripsWatchdogTimeout (0x12C) 
    "EnableInputSuppression" = 4294967295; // PopEnableInputSuppressionOverride (4294967295) 
    "EnableMinimalHiberFile" = 0; // PopEnableMinimalHiberFile 
    "EnablePowerButtonSuppression" = 4294967295; // PopEnablePowerButtonSuppressionOverride (4294967295) 
    "EnergyEstimationEnabled" = 1; // PopEnergyEstimationEnabled 
    "EnforceAusterityMode" = 0; // PopEnforceAusterityMode 
    "EnforceConsoleLockScreenTimeout" = 0; // PopEnforceConsoleLockScreenTimeout 
    "EnforceDisconnectedStandby" = 0; // PopEnforceDisconnectedStandby 
    "EventProcessorEnabled" = 1; // PopEventProcessorEnabled 
    "ExitLatencyCheckEnabled" = 0; // PpmExitLatencyCheckEnabled 
    "ExperimentalClusterIdleMitigation" = 0; // PpmIdleClusterIdleMitigation 
    "ForceMinimalHiberFile" = 0; // PopForceMinimalHiberFile 
    "FxAccountingTelemetryDisabled" = 0; // PopDiagFxAccountingTelemetryDisabled 
    "FxRuntimeLogNumberEntries" = 64; // PopFxRuntimeLogNumberEntries (0x40) - Changing it to 0 will end up with a BSoD
    "HeteroFavoredCoreRotationTimeoutMs" = 30000; // PpmHeteroFavoredCoreRotationTimeoutMs (0x00007530) 
    "HeteroHgsEePerfHintsIndependentEnabled" = 0; // PpmHeteroHgsEePerfHintsIndependentEnabled 
    "HeteroHgsPlusDisabled" = 0; // PpmHeteroHgsThreadDisabled 
    "HeteroMultiClassParkingEnabled" = 4294967295; // PpmHeteroMultiClassParkingRegValue (4294967295) 
    "HeteroMultiCoreClassesEnabled" = 4294967295; // PpmHeteroMultiCoreClassesRegValue (4294967295) 
    "HeteroWpsContainmentEnumOverride" = 0; // PpmHeteroWpsContainmentEnumOverride 
    "HeteroWpsWorkloadProminenceCutoff" = 35; // PpmHeteroWpsWorkloadProminenceCutoff (0x23) 
    "HiberbootEnabled" = 0; // PopHiberbootEnabledReg 
    "HiberFileSizePercent" = 100; // PopHiberFileSizePercent dd 64h (IDA), but set to 0 by default on LTSC IoT Enterprise 2024 since hibernation is unsupported by default - REG_DWORD
    "HiberFileType" = 4294967295; // PopHiberFileTypeReg (4294967295)
    "HiberFileTypeDefault" = 4294967295; // PopHiberFileTypeDefaultReg (4294967295)
    "HibernateBootOptimizationEnabled" = 0; // PopHiberBootOptimizationEnabledReg 
    "HibernateChecksummingEnabled" = 1; // PopHiberChecksummingEnabledReg 
    "HibernateEnabledDefault" = 1; // PopHiberEnabledDefaultReg - REG_DWORD
    "HighPerfDurationBoot" = 90000; // PpmHighPerfDuration (0x00015F90) 
    "HighPerfDurationCSExit" = ?; // unk_140FC337C
    "HighPerfDurationSxExit" = ?; // unk_140FC3380
    "IdleDurationExpirationTimeout" = 4; // PpmIdleDurationExpirationTimeoutMs 
    "IdleProcessorsRequireQosManagement" = 4294967295; // PpmPerfQosManageIdleProcessors (4294967295) 
    "IdleStateTimeout" = 500; // PopPepIdleStateTimeout (0x000001F4) 
    "IgnoreCsComplianceCheck" = 0; // PopIgnoreCsComplianceCheck 
    "IgnoreLidStateForInputSuppression" = 4294967295; // PopLidStateForInputSuppressionOverride (4294967295) 
    "IpiLastClockOwnerDisable" = 0; // PpmIpiLastClockOwnerDisable 
    "LatencyToleranceDefault" = 100000; // PpmLatencyToleranceLimit (0x000186A0) 
    "LatencyToleranceFSVP" = 20000; // dword_140FC3428 dd 4E20
    "LatencyToleranceIdleResiliency" = 1500000; // dword_140FC342C dd 16E360
    "LatencyToleranceParked" = 0; // PpmIdleParkedLatencyLimit 
    "LatencyToleranceSoftParked" = 0; // PpmIdleSoftParkedLatencyLimit 
    "LatencyToleranceVSyncEnabled" = 13001; // dword_140FC3424 dd 32C9
    "LidReliabilityState" = 1; // REG_DWORD, range: 0-1
    "ManualDimTimeout" = 0; // PopAdaptiveManualDimTimeout 
    "MaximumFrequencyOverride" = 0; // PpmFrequencyOverride 
    "MfBufferingThreshold" = 0; // PpmMfBufferingThreshold 
    "MfOverridesDisabled" = 1; // PpmMfOverridesDisabled 
    "MSDisabled" = 0; // PopModernStandbyDisabled 
    "MultiparkGranularity" = 8; // PpmParkMultiparkGranularity 
    "PdcIdlePhaseDefaultWatchdogTimeoutSeconds" = 30; // PopPdcIdlePhaseDefaultWatchdogTimeoutSeconds (0x0000001E) 
    "PdcOneWayEntry" = 0; // PopPowerAggregatorOneWayEntry 
    "PerfArtificialDomain" = 4294967295; // PpmPerfArtificialDomainSetting (4294967295) 
    "PerfBoostAtGuaranteed" = 0; // PpmPerfBoostAtGuaranteed 
    "PerfCalculateActualUtilization" = 1; // PpmPerfCalculateActualUtilization 
    "PerfCheckTimerImplementation" = 0; // PpmCheckTimerImplementation 
    "PerfIdealAggressiveIncreasePolicyThreshold" = 90; // PpmPerfIdealAggressiveIncreaseThreshold (0x5A) 
    "PerfQueryOnDevicePowerChanges" = 0; // PopFxPerfQueryOnDevicePowerChanges 
    "PerfSingleStepSize" = 5; // PpmPerfSingleStepSize (0x05) 
    "PlatformAoAcOverride" = 4294967295; // PopPlatformAoAcOverride (4294967295) 
    "PlatformRoleOverride" = 4294967295; // PopPlatformRoleOverride (4294967295) 
    "PoFxSystemIrpWaitForReportDevicePowered" = 0; // PopPoFxSystemIrpWaitForReportDevicePoweredReg 
    "PowerActionResumeWatchdogTimeoutDefault" = 300; // PopPowerActionResumingWatchdogTimeoutDefault (0x0000012C) 
    "PowerActionTransitioningWatchdogTimeoutDefault" = 600; // PopPowerActionTransitioningWatchdogTimeoutDefault (0x00000258) 
    "PromoteHibernateToShutdown" = 0; // PopPromoteHibernateToShutdown 
    "ProximityEscapeMsec" = 0; // TtmpProximityEscapeMsec 
    "RestrictedStandbyDozeTimeoutSeconds" = 0; // PopPowerAggregatorRestrictedStandbyDozeTimeoutSeconds 
    "SkipHibernateMemoryMapValidation" = 4294967295; // PopEnableHibernateMemoryMapValidationOverride (4294967295) 
    "SleepstudyAccountingEnabled" = 1; // SleepstudyHelperAccountingEnabled 
    "SleepstudyGlobalBlockerLimit" = 3000; // SleepstudyHelperBlockerGlobalLimit (0x0BB8) 
    "SleepstudyLibraryBlockerLimit" = 200; // SleepstudyHelperBlockerLibraryLimit (0xC8) 
    "SmartUserPresenceAction" = 0; // PopSmartUserPresenceAction 
    "SmartUserPresenceCheckTimeout" = 10800; // PopSmartUserPresenceCheckTimeout (0x00002A30) 
    "SmartUserPresenceGracePeriod" = 1800; // PopSmartUserPresenceGracePeriod (0x00000708) 
    "SmartUserPresenceWakeOffset" = 300; // PopSmartUserPresenceWakeOffset (0x0000012C) 
    "StandbyConnectivityGracePeriod" = 0; // PopStandbyConnectivityGracePeriod 
    "SuppressResumePrompt" = 0; // PopSuppressResumePrompt 
    "ThermalPollingMode" = 0; // PopThermalPollingMode 
    "ThermalTelemetryVerbosity" = 1; // PopThermalTelemetryVerbosity 
    "TimerRebaseThresholdOnDripsExit" = 60; // PopTimerRebaseThresholdRegValue (0x3C) 
    "TtmEnabled" = 0; // TtmpEnabled 
    "UserBatteryChargeEstimator" = 0; // PopUserBatteryChargingEstimator 
    "UserBatteryDischargeEstimator" = 0; // PopDisableBatteryDischargeEstimator 
    "WatchdogWorkOrderTimeout" = 300000; // PopFxWatchdogWorkOrderTimeout (0x000493E0) 
    "Win32kCalloutWatchdogTimeoutSeconds" = 30; // PopWin32kCalloutWatchdogTimeoutSeconds (0x0000001E) 

    // UmpoRestoreEsOverrideState
    "EnergySaverState" = 2; // 1 = override state (more power savings)? if != 1 no override? (WNF_PO_ENERGY_SAVER_OVERRIDE/WNF_SEB_ENERGY_SAVER_STATE_V2)

    // InitializePowerWatchdogTimeoutDefaults
    "PowerWatchdogDrvSetMonitorTimeoutMsec" = 10000; // v10[13]
    "PowerWatchdogDwmSyncFlushTimeoutMsec" = 30000; // v10[10]
    "PowerWatchdogPoCalloutTimeoutMsec" = 10000;
    "PowerWatchdogPowerOnGdiTimeoutMsec" = 30000;
    "PowerWatchdogRequestQueueTimeoutMsec" = 30000;

    // from procmon boot trace
    "DisableHotKeyWhenConsoleOff" = ?;
    "EmiPollingInterval" = ?;
    "EmiTelemetryActivePollingInterval" = ?;
    "EmiTelemetryCsPollingInterval" = ?;
    "LidNotifyReliable" = ?;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\ForceHibernateDisabled";
    "GuardedHost" = ?; // unk_140FC5234
    "Policy" = 0; // PopHiberForceDisabledReg 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\HiberFileBucket";
    "Percent16GBFull" = ?; // unk_140FC36D0 - 28Hex/40Dec?
    "Percent16GBReduced" = ?; // unk_140FC36CC - 14Hex/20Dec?
    "Percent1GBFull" = ?; // unk_140FC3670 - 28Hex/40Dec?
    "Percent1GBReduced" = ?; // unk_140FC366C - 14Hex/20Dec?
    "Percent2GBFull" = ?; // unk_140FC3688 - 28Hex/40Dec?
    "Percent2GBReduced" = ?; // unk_140FC3684 - 14Hex/20Dec?
    "Percent32GBFull" = ?; // unk_140FC36E8 - 28Hex/40Dec?
    "Percent32GBReduced" = ?; // unk_140FC36E4 - 14Hex/20Dec?
    "Percent4GBFull" = ?; // unk_140FC36A0 - 28Hex/40Dec?
    "Percent4GBReduced" = ?; // unk_140FC369C - 14Hex/20Dec?
    "Percent8GBFull" = ?; // unk_140FC36B8 - 28Hex/40Dec?
    "Percent8GBReduced" = ?; // unk_140FC36B4 - 14Hex/20Dec?
    "PercentUnlimitedFull" = ?; // unk_140FC3700 - 28Hex/40Dec?
    "PercentUnlimitedReduced" = ?; // unk_140FC36FC - 14Hex/20Dec?

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\ModernSleep";
    "EnabledActions" = 0; // PopAggressiveStandbyActionsRegValue 
    "EnableDsNetRefresh" = 0; // PopEnableDsNetRefresh 

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerThrottling";
    "PowerThrottlingOff" = 0; // PpmPerfQosGroupPolicyDisable 
```

## DWM Values

See [dwm.c](https://github.com/nohuto/win-registry/blob/main/assets/dwm.c) for used snippets (taken from `dwmcore.dll`, `win32full.sys`, `dwm.exe`, `dwminit.dll`, `uDWM.dll`).

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

```c
"HKLM\\SOFTWARE\\Microsoft\\Windows\\DwmMicrosoft\\Windows\\Dwm";
    "BlackOutAllReadback" = 0;
    "ConfigureInput" = 1;
    "CpuClipAASinkEnableIntermediates" = 1;
    "CpuClipAASinkEnableOcclusion" = 1;
    "CpuClipAASinkEnableRender" = 1;
    "CpuClipAreaThreshold" = 20000;
    "CpuClipWarpPartitionThreshold" = 1024;
    "DisableDrawListCaching" = 0; // REG_DWORD
    "DisableProjectedShadows" = 0;
    "DisplayChangeTimeoutMs" = 1000;
    "EnableBackdropBlurCaching" = 1; // REG_DWORD
    "EnableCommonSuperSets" = 1;
    "EnableCpuClipping" = 1;
    "EnableDDisplayScanoutCaching" = 1;
    "EnableEffectCaching" = 1; // REG_DWORD
    "EnableFrontBufferRenderChecks" = 1;
    "EnableMegaRects" = 1;
    "EnablePrimitiveReordering" = 1;
    "ForceFullDirtyRendering" = 0;
    "GammaBlendPencil" = 1;
    "GammaBlendWithFP16" = 1;
    "InkGPUAccelOverrideVendorWhitelist" = 0;
    "LayerClippingMode" = 2;
    "LogExpressionPerfStats" = 0; // REG_DWORD
    "MajorityScreenTest_MinArea" = 80;
    "MajorityScreenTest_MinLength" = 80;
    "MaxD3DFeatureLevel" = 0;
    "MegaRectSearchCount" = 100;
    "MegaRectSize" = 100000;
    "MousewheelAnimationDurationMs" = 250;
    "MousewheelScrollingMode" = 0; // REG_DWORD
    "OptimizeForDirtyExpressions" = 1; // REG_DWORD
    "OverlayMinFPS" = 15; // If this value is present and set to zero, the Desktop Window Manager disables its minimum frame rate requirement for assigning DirectX swap chains to overlay planes in hardware that supports overlays. This makes it more likely that a low frame rate swap chain will get assigned and stay assigned to an overlay plane, if available. (https://github.com/MicrosoftDocs/win32/blob/docs/desktop-src/dwm/registry-values.md)
    "RenderThreadTimeoutMilliseconds" = 5000;
    "SuperWetExtensionTimeMicroseconds" = 1000;
    "TelemetryFramesReportPeriodMilliseconds" = 300000;
    "TelemetryFramesSequenceIdleIntervalMilliseconds" = 1000;
    "TelemetryFramesSequenceMaximumPeriodMilliseconds" = 1000;
    "UniformSpaceDpiMode" = 1;
    "UseFastestMonitorAsPrimary" = 0;
    "vBlankWaitTimeoutMonitorOffMs" = 250;
    "WarpEnableDebugColor" = 0;

    "BackdropBlurCachingThrottleMs" = 25; // 25ms if missing, clamped to <=1000ms when present?
    "CompositorClockPolicy" = 1; // range: 0-1
    "CpuClipFlatteningTolerance" = 0; // scaled /1000
    "CustomRefreshRateMode" = 0; // range: 0-2
    "DisableAdvancedDirectFlip" = 0; // REG_DWORD
    "DisableIndependentFlip" = 0;
    "DisableProjectedShadowsRendering" = 0;
    "FlattenVirtualSurfaceEffectInput" = 0;
    "ForceEffectMode" = 0; // range: 0-2, REG_DWORD
    "FrameCounterPosition" = 0;
    "InteractionOutputPredictionDisabled" = 0;
    "OverlayTestMode" = 0; // 5 = MPO disabled, REG_DWORD
    "ParallelModePolicy" = 1; // >=3 coerced to 1
    "ParallelModeRateThreshold" = 119; // divisor for g_qpcFrequency, missing key defaults to 119 Hz (units: Hz)? 0 disables
    "ResampleInLinearSpace" = 0;
    "ResampleModeOverride" = 0;
    "SDRBoostPercentOverride" = 0; // scaled /100
    "ShowDirtyRegions" = 0;

    "AnimationsShiftKey" = 0;
    "DisableLockingMemory" = 0;
    "ModeChangeCurtainUseDebugColor" = 0;
    "UseDPIScaling" = 1;

    "ChildWindowDpiIsolation" = 1; // range: 0-1
    "DisableDeviceBitmaps" = 0; // range: 0-1
    "EnableResizeOptimization" = 0; // range: 0-1, REG_DWORD
    "ResizeTimeoutGdi" = 0; // range: 0-4294967295 (ms)
    "ResizeTimeoutModern" = 0; // range: 0-4294967295 (ms)

    "DefaultColorizationColorState" = 0;
    "DisallowAnimations" = 0;
    "DisallowColorizationColorChanges" = 0;

    "DisableSessionTermination" = 0; // range: 0–1
    "ForceBasicDisplayAdapterOnDWMRestart" = 0; // range: 0–1
    "OneCoreNoBootDWM" = 0; // range: 0–1, REG_DWORD
    "OneCoreNoDWMRawGameController" = ? // didn't look into it yet, but it's related to OneCoreNoBootDWM

    "DisableHologramCompositor" = 0; // range: 0–1

    // Haven't looked into them yet
    "ForceUDwmSoftwareDevice" = ?;
    "ForceDisableModeChangeAnimation" = ?; // REG_DWORD

    // procmon boot trace
    "AccentColorInactive" = ?;
    "AnimationAttributionEnabled" = 1; // REG_DWORD
    "AnimationAttributionHashingEnabled" = 1; // REG_DWORD
    "ColorPrevalence" = ?;
    "CpuClipAASinkEnableDebugColors" = ?;
    "CpuClipAASinkForceEnable" = ?;
    "DebugFailFast" = ?;
    "DisableDeviceBitmapsForMultiAdapter" = ?;
    "DwmInitSessionActivityId_00000001" = ?; // a ID, REG_SZ
    "EnableDesktopOverlays" = ?;
    "EnableMPCPerfCounter" = ?;
    "EnableRenderPathTestMode" = ?;
    "EnableWindowColorization" = ?;
    "ForceDesktopTreeFullDirty" = ?;
    "MajorityScreenTest_MaxCoverage" = ?;
    "MarshalAllDebugInfo" = ?;
    "MPCInputRouterWaitForDebugger" = ?;
    "ShaderLinkingGPUBlacklist" = ?; // REG_SZ
    "SuperWetEnabled" = ?;
    "UseHWDrawListEntriesOnWARP" = ?;


"HKLM\\SOFTWARE\\Microsoft\\Windows\\Dwm\\Scene";
    "EnableBloom" = 0; // REG_DWORD
    "EnableDrawToBackbuffer" = 1; // REG_DWORD
    "EnableImageProcessing" = 1; // REG_DWORD
    "ImageProcessingResizeGrowth" = 200;
    "MsaaQualityMode" = 2;
    "SceneVisualCutoffCountOfConsecutiveIncidentsAllowed" = 5;
    "SceneVisualCutoffThresholdInMS" = 1000;

    "ForceNonPrimaryDisplayAdapter" = 0; // REG_DWORD
    "ImageProcessingResizeThreshold" = 0; // scaled /100

"HKLM\\SOFTWARE\\Microsoft\\Windows\\Dwm\\GpuAccelInkTiming";
    "ExtensionTimeMicroseconds" = 1000;
    "PeriodicFenceMinDifferenceMicroseconds" = 500;
    "RefreshRatePercentage" = 10;
```

## USB/USBHUB/USBFLAGS Values

For entries described as "any nonzero", the code treats the DWORD as a boolean, means any nonzero value is equivalent to `1`. Default data is unknown for most values as the driver code only reads the registry and handles fallbacks, note that this is currently based on USBHUB3.sys only, means it's not complete (USBXHCI.sys was used for DisableHCS0Idle & TestRunEsmInWorkItem, Ucx01000.sys for Allow64KLowOrFullSpeedControlTransfers).

This documentation doesn't include all details, since the repo is used for showing registry values, their default data, ranges and miscellaneous information. See [desc.md#usbflags-values](https://github.com/nohuto/win-config/blob/main/peripheral/desc.md#usbflags-values), [desc.md#usb-values](https://github.com/nohuto/win-config/blob/main/peripheral/desc.md#usb-values), [desc.md#usbhub-values](https://github.com/nohuto/win-config/blob/main/peripheral/desc.md#usbhub-values) for more details. The USBHUB3.sys included some values located in the device hardware key like `DeviceIdleEnabled`, `DefaultIdleState`, `DeviceIdleIgnoreWakeEnable`, I didn't add them here since it lacks on details, see [desc.md#disable-device-powersavings](https://github.com/nohuto/win-config/blob/main/power/desc.md#disable-device-powersavings).

`HUBDSM_QueryingRegistryValuesForDevice` -> `HUBMISC_QueryAndCacheRegistryValuesForDevice` -> `HUBREG_QueryUsbflagsValuesForDevice`

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\usbflags";
    "IgnoreHWSerNum<vvvvpppp>" = ?; // REG_DWORD, dynamic name, seems to use vvvvpppp (vendor, product) see documentation below
    "Allow64KLowOrFullSpeedControlTransfers"; ? = // REG_DWORD (bool), if 1 sets g_Allow64KLowOrFullSpeedControlTransfersFlag true, other values leave it false, missing/read failure leaves the global unchanged - "When enabled, low/full-speed control endpoints use 64KB max transfer size, otherwise table defaults apply. Query failures are non-fatal."
    "DisableHCS0Idle" = 0; // REG_DWORD (bool), any nonzero disables S0 idle, missing/read failure treated as 0 - probably means "Disable Host Controller S0 Idle" (S = System)
    "GenericCompositeUSBDeviceString" = // haven't looked into it yet
    "SetMultiTTBitDuringConfigureEndpoint" = ? // REG_DWORD, if 1 sets flag 0x2000000000000000 at a1+336? if 0 clears it, missing/read failure leaves flags unchanged
    "TestRunEsmInWorkItem" = 0; // REG_DWORD, if 1 sets bit 0 at a1+876? if 0 clears it, missing/read failure leaves bit cleared

// these are built by HUBREG_OpenCreateUsbflagsDeviceKey
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\usbflags\\<vvvvpppprrrr>";
    "IgnoreHWSerNum" = ?; // REG_DWORD, any nonzero value sets flag 0x1
    "UseWin8DescriptorValidation" = ?; // REG_DWORD, any nonzero value sets flag 0x200000
    "ResetOnResume" = ?; // REG_DWORD, any nonzero value sets flag 0x4
    "DisableOnSoftRemove" = ?; // REG_DWORD, if zero clears flag 0x8, any nonzero leaves it set, many inf files set it to 1 so I guess that the default is 1
    "RequestConfigDescOnReset" = ?; // REG_DWORD, any nonzero value sets flag 0x10. When set, device hack flag causes config descriptor request after reset
    "DisableRecoveryFromPowerDrain" = ?; // REG_DWORD, any nonzero value sets flag 0x800000
    "SkipContainerIdQuery" = ?; // REG_BINARY, any nonzero value sets flag 0x20 (1 means skip container ID query)
    "DisableLPM" = ?; // REG_DWORD (bool), any nonzero value sets flag 0x80 - When set, disables low power link states for the device = the driver skips enabling U2 and forces U1/U2 timeouts to 0 so link power management stays off. Also disabled when a parent hub indicates LPM should be off for all downstream devices. (this is how I understood it while reading through the W10 source code)
                      // "A link enters a low power state (consuming less power than the working state) only when the downstream device enters the suspended state through the selective suspend mechanism", "After remaining idle for a certain period of time, link partners progressively enter U1 (standby with fast exit) and then U2 (standby with slower exit)"
                      // https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/usb-3-0-lpm-mechanism- https://learn.microsoft.com/en-us/windows-hardware/drivers/usbcon/u1-and-u2-transitions
    "SkipBOSDescriptorQuery" = ?; // REG_DWORD, any nonzero value sets flag 0x8000
    "MsOs20DescriptorSetInfo" = ?; // REG_BINARY/REG_QWORD (8 bytes)??, on success stores QWORD and sets internal flag 0x4 for the descriptor info
    "osvc" = ?; // REG_BINARY, see documentation below (written by HUBMISC_StoreDeviceMSOSVendorCodeInRegsitry)

    "AlternateSettingFilter"; // ?
    "ResetTTOnCancel"; = ?;
    "NoClearTTBufferOnCancel" = ?;
    "PowerUpDelay" = ?;
```

> [peripheral/assets | usbflags-HUBDSM_QueryingRegistryValuesForDevice.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-HUBDSM_QueryingRegistryValuesForDevice.c)  
> [peripheral/assets | usbflags-HUBMISC_QueryAndCacheRegistryValuesForDevice.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-HUBMISC_QueryAndCacheRegistryValuesForDevice.c)  
> [peripheral/assets | usbflags-HUBREG_OpenCreateUsbflagsDeviceKey.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-HUBREG_OpenCreateUsbflagsDeviceKey.c)  
> [peripheral/assets | usbflags-HUBREG_QueryUsbflagsValuesForDevice.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-HUBREG_QueryUsbflagsValuesForDevice.c)  
> [peripheral/assets | usbflags-Controller_IsRegKeySetToDisableS0Idle.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-Controller_IsRegKeySetToDisableS0Idle.c)  
> [peripheral/assets | usbflags-Controller_PopulateRegistryOverrideForSetMultiTTBitFlag.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-Controller_PopulateRegistryOverrideForSetMultiTTBitFlag.c)  
> [peripheral/assets | usbflags-Controller_PopulateTestRegistrySettings.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-Controller_PopulateTestRegistrySettings.c)  
> [peripheral/assets | usbflags-Registry_InitializeAllow64KLowOrFullSpeedControlTransfersFlag.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbflags-Registry_InitializeAllow64KLowOrFullSpeedControlTransfersFlag.c)

```c
// HUBREG_QueryGlobalUsb20HardwareLpmSettings
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\Usb20HardwareLpm"; // g_Usb20HardwareLpmKeyName (aRegistryMachin_8)
    "Usb20HardwareLpmOverride" = ?; // REG_DWORD, any nonzero keeps global flag 0x8000 set, zero clears it
    "Usb20HardwareLpmTimeout" = 2; // REG_DWORD, if value == 0xFF, byte a1+72 set to 0xFF, didn't find other handles

// HUBREG_OpenQueryAttemptRecoveryFromUsbPowerDrainValue
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\AutomaticSurpriseRemoval"; // g_UsbAutomaticSurpriseRemovalKeyName (aRegistryMachin_6)
    "AttemptRecoveryFromUsbPowerDrain" = ?; // REG_DWORD, queried directly and via persisted state fallback under Control\usb

// HUBREG_QueryUsbHardwareVerifierValue
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\HardwareVerifier"; // g_HwVerifierKeyName (aRegistryMachin_7)
    "<VID><PID><REV>\\usbUpto20|usb2X|usb30\\device" = ?; // REG_DWORD, subkey chosen from bcdUSB, value name is g_HwVerifierDeviceName
    "<VID><PID>\\usbUpto20|usb2X|usb30\\device" = ?; // REG_DWORD, fallback if VID/PID/REV key missing
    "global\\usbUpto20|usb2X|usb30\\device" = ?; // REG_DWORD, final fallback if no device specific key

// HUBREG_QueryGlobalUsbLtmSettings
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\UsbLtm"; // g_UsbLtmKeyName (aRegistryMachin_4)
    "UsbLtmEnable" = 0; // REG_DWORD, any nonzero value sets global flag 0x40000, zero clears it

// these are taken from the W10 source, they seem to exist on latest builds (they do exist in usbport.sys on 23H2)
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\usb";
    "debuglevel" = 0; // REG_DWORD, default to 1/2 when DEBUG1/DEBUG2 builds. Trace verbosity, higher numbers enable more logs
    "debuglogmask" = 0xFFFFFFFE; // REG_DWORD, Bitmask for log categories
    "debuglogenable" = 1; // REG_DWORD (bool), enables debug log output
    "debugcatc" = 0; // REG_DWORD (bool), enables CATC analyzer trigger
    "DisableSelectiveSuspend" = 0; // REG_DWORD (bool), global disable for selective suspend (GlobalUsbhubLegacyValues?)
    "DisableCcDetect" = 0; // REG_DWORD (bool), global disable for CC detection
    "EnPMDebug" = 0; // REG_DWORD (bool), for debugging power management
    "ForceHcD3NoWakeArm" = 0 // REG_DWORD (bool), prevents wake-arming when forcing HC to D3
    "EnableDCA" = 0 // REG_DWORD (bool), enables direct controller access (HCT diagnostics)
    "ForcePortsHighSpeed" = 0; // REG_DWORD (bool), forces ports to remain under EHCI (HCT compatibility)

// "This class is reserved for USB host controllers and USB hubs", I'll add them here as they're also in usbport.sys and also taken from the W10 source
"HKLM\\System\\CurrentControlSet\\Control\\Class\\{36FC9E60-C465-11CF-8056-444553540000}\\<instance>";
    "HcFlavor" = ? // REG_DWORD, auto detect. Values are USB_CONTROLLER_FLAVOR enum (definition external)
    "TotalBusBandwidth" = ? // REG_DWORD, computed from miniport registration (bits/ms). Overrides bus bandwidth accounting
    "HcDisableAllSelectiveSuspend" = 0 (non-IA64), 1 (IA64); // REG_DWORD, non-zero disables selective suspend
    "CommonBuffer2GBLimit" = 0; // REG_DWORD, when non-zero, forces common buffers below 2GB ("Limit common buffer allocations for the miniport to the physical address range below 2GB.  Only bits 0 through 30 of the physical address can be set.  Bit 31 of the physical address cannot be set.")
    "ForceHCResetOnResume" = 0; // REG_DWORD, forces controller reset on resume
    "FastResumeEnable" = 0; // REG_DWORD, enables fast S0 resume

// miscellaneous note for future reference
"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Usb\\Ceip" // UsbhUpdateRegSurpriseRemovalCount
    "BootPathSurpriseRemovalCount" = ?;
```

> [peripheral/assets | usb-GetPersistedKeyPath.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usb-GetPersistedKeyPath.c)  
> [peripheral/assets | usb-HUBREG_OpenQueryAttemptRecoveryFromUsbPowerDrainValue.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usb-HUBREG_OpenQueryAttemptRecoveryFromUsbPowerDrainValue.c)  
> [peripheral/assets | usb-HUBREG_QueryGlobalUsb20HardwareLpmSettings.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usb-HUBREG_QueryGlobalUsb20HardwareLpmSettings.c)  
> [peripheral/assets | usb-HUBREG_QueryGlobalUsbLtmSettings.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usb-HUBREG_QueryGlobalUsbLtmSettings.c)  
> [peripheral/assets | usb-HUBREG_QueryUsbHardwareVerifierValue.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usb-HUBREG_QueryUsbHardwareVerifierValue.c)  
> [peripheral/assets | usb-ReadManifestAssignedValue.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usb-ReadManifestAssignedValue.c)

```c
// HUBREG_QueryGlobalHubValues
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\USBHUB\\hubg"; // g_HubGlobalKeyName (aRegistryMachin_10)
    "DisableSelectiveSuspendUI" = ?; // REG_DWORD, any nonzero value sets global flag 0x2
    "MsOsDescriptorMode" = ?; // REG_DWORD, stored to a1+8 when 0..3, out of range values are logged, MsOsDescriptorMode == 1 will force a MS OS descriptor query for all devices, regardless of USB version number. MsOsDescriptorMode == 2 will disable all MS OS descriptor queries. (there may exist a value named "DontSkipMsOsDescriptor" in this key)
    "EnableDiagnosticMode" = ?; // REG_DWORD, any nonzero value sets global flag 0x8
    "DisableOnSoftRemove" = ?; // REG_DWORD, if zero clears global flag 0x80 (default on) any nonzero leaves it set
    "DisableUxdSupport" = ?; // REG_DWORD, any nonzero value sets global flag 0x10
    "EnableExtendedValidation" = ?; // REG_DWORD bitmask, any nonzero value sets global flag 0x20, bit 0x8 sets 0x2000, bit 0x4 sets 0x4000
    "WakeOnConnectUI" = ?; // REG_DWORD, any nonzero value sets global flag 0x40 - This controls the UI check box 'Allow this device to wake the system'.  Essentially this is control for the wake on connect feature.
    "PreventDebounceTimeForSuperSpeedDevices" = ?; // REG_DWORD, any nonzero value sets global flag 0x10000 - Checks if we need to give extra time to SuperSpeed devices before talking to them

    // miscellaneous ones from GlobalUsbhubValues
    "UsbDebugModeEnable" = ?;
    "BreakOnHubException" = ?;
    "debuglevel" = ?;
    "DebugLogMask" = ?;
    "DebugLogEnable" = ?;
    "DisableHardReset" = ?;
    "BreakOnReplicant" = ?;
    "BreakOnEnumFailure" = ?;
    "UseIoErrorLog" = ?;
    "ForceResetOnResume" = ?;
    "DisableFastResume" = ?;
    "LogSize" = ?;
    "IdleTimeout" = ?;

// HUBREG_QueryGlobalUxdSettings (the defaults were taken from the W10 source)
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\uxd_control\\policy"; // g_UxdGlobalSettingsKey (aRegistryMachin_12)
    "UxdGlobalDeleteOnShutdown" = 0; // REG_DWORD (bool), any nonzero value sets global flag 0x100, global delete-on-shutdown policy
    "UxdGlobalDeleteOnReload" = 0; // REG_DWORD (bool), any nonzero value sets global flag 0x200, global policy to delete UXD keys on disable/reload events
    "UxdGlobalDeleteOnDisconnect" = 0; // REG_DWORD (bool), any nonzero value sets global flag 0x400, global policy to delete UXD keys on device disconnect
    "UxdGlobalEnable" = 0; // REG_DWORD (bool), any nonzero value sets global flag 0x800, "main" enable = if 0, UXD settings are ignored

// HUBREG_QueryUxdDeviceKey / HUBREG_DeleteUxdDeviceKey
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\uxd_control\\devices"; // g_UxdDeviceSettingsKey (aRegistryMachin_5)
    "%04X%04X%04X" = ?; // value name from VID/PID/REV

// HUBREG_GetUxdPnpValue
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\uxd_control\\pnp"; // g_UxdGuidSettingsKey (aRegistryMachin_3)
    "{GUID}" = ?; // value name from RtlStringFromGUID, data queried via WDF
```

> [peripheral/assets | usbhub-HUBREG_QueryUxdDeviceKey.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbhub-HUBREG_QueryUxdDeviceKey.c)  
> [peripheral/assets | usbhub-HUBREG_DeleteUxdDeviceKey.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbhub-HUBREG_DeleteUxdDeviceKey.c)  
> [peripheral/assets | usbhub-HUBREG_QueryGlobalUxdSettings.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbhub-HUBREG_QueryGlobalUxdSettings.c)  
> [peripheral/assets | usbhub-HUBREG_QueryGlobalHubValues.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbhub-HUBREG_QueryGlobalHubValues.c)  
> [peripheral/assets | usbhub-HUBREG_GetUxdPnpValue.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/usbhub-HUBREG_GetUxdPnpValue.c)  

```c
aRegistryMachin_1 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\USBFN";
aRegistryMachin_2 = // doesn't exist
aRegistryMachin_3 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\uxd_control\\pnp";
aRegistryMachin_4 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\UsbLtm";
aRegistryMachin_5 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\uxd_control\\devices";
aRegistryMachin_6 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\AutomaticSurpriseRemoval";
aRegistryMachin_7 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\HardwareVerifier";
aRegistryMachin_8 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb\\Usb20HardwareLpm";
aRegistryMachin_9 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\usbflags";
aRegistryMachin_10 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\USBHUB\\hubg";
aRegistryMachin_11 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\USB";
aRegistryMachin_12 = "HKLM\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\uxd_control\\policy";
aRegistryMachin_13 = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\usb";
```

## BCD Edits

See details about BCDEdit, used resources, related pseudocode etc. here [win-config/system/bcd-edits](https://github.com/nohuto/win-config/blob/main/system/desc.md#bcd-edits).

As kind of everything else, BCD edits are also stored in the registry:
```c
HKLM\BCD00000000\Objects

// Structure
HKLM\BCD00000000\Objects\{GUID} // {GUID} depends on the object, e.g. {bootmgr}, {current}, {globalsettings}
HKLM\BCD00000000\Objects\{GUID}\Elements\XXXXXXXX // XXXXXXXX is a specific setting for the object (8 digit)
HKLM\BCD00000000\Objects\{GUID}\Elements\XXXXXXXX : Element // (REG_BINARY/REG_MULTI_SZ/REG_SZ - depends on the setting) this value includes the state of the setting
```

See all object identifiers via `bcdedit /enum all /v` (`identifier`). Note that the list below uses `{bootmgr}`, `{current}` etc. which must be replaced by the actual GUID (see block above).

Here are elements which I tracked via Procmon (taken from default store and the MS documentation), note that this doesn't show default states (see block at the buttom), instead it shows several options and their possible states:
```c
"HKLM\\BCD00000000\\Objects\\{current}\\Elements";
    "\\26000141"; "Element" = 01; // REG_BINARY, event = true, false = 00
    "\\26000116"; "Element" = 01; // REG_BINARY, hypervisorusevapic = true, false = 00
    "\\260000F8"; "Element" = 01; // REG_BINARY, hypervisordisableslat = true, false = 00
    "\\260000FC"; "Element" = 01; // REG_BINARY, hypervisoruselargevtlb = true, false = 00 - Increases virtual Translation Lookaside Buffer (TLB) size.
    "\\260000F2"; "Element" = 01; // REG_BINARY, hypervisordebug = true, false = 00 - Controls whether the hypervisor debugger is enabled.
    "\\260000E1"; "Element" = 00; // REG_BINARY, disableelamdrivers = false, true = 01 - The OS loader removes this entry for security reasons. This option can only be triggered by using the F8 menu.
    "\\260000C3"; "Element" = 01; // REG_BINARY, onetimeadvancedoptions = true, false = 00 - Controls whether the system boots to the legacy menu (F8 menu) on the next boot.
    "\\260000C4"; "Element" = 01; // REG_BINARY, onetimeoptionsedit = true, false = 00
    "\\260000B0"; "Element" = 01; // REG_BINARY, ems = true, false = 00 - Indicates whether EMS should be enabled in the kernel.
    "\\260000A5"; "Element" = 01; // REG_BINARY, disabledynamictick = true, false = 00
    "\\260000A4"; "Element" = 01; // REG_BINARY, useplatformtick = true (forces platform clock source, often HPET), false = 00
    "\\260000A3"; "Element" = 01; // REG_BINARY, forcelegacyplatform = true, false = 00 - Forces the OS to assume the presence of legacy PC devices like CMOS and keyboard controllers.
    "\\260000A2"; "Element" = 01; // REG_BINARY, useplatformclock = true (forces the use of the platform clock as the system's performance counter), false = 00
    "\\260000A1"; "Element" = 01; // REG_BINARY, halbreakpoint = true, false = 00 - Indicates whether the HAL should call DbgBreakPoint at the start of HalInitSystem for phase 0 initialization of the kernel.
    "\\260000A0"; "Element" = 01; // REG_BINARY, debug = true, false = 00 - Indicates whether the kernel debugger should be enabled using the settings in the inherited debugger object.
    "\\26000091"; "Element" = 01; // REG_BINARY, sos = true, false = 00 - Indicates whether the system should display verbose information.
    "\\26000090"; "Element" = 01; // REG_BINARY, bootlog = true, false = 00 - Indicates whether the system should write logging information to %SystemRoot%\Ntbtlog.txt during initialization.
    "\\25000080"; "Element" = 0000000000000000; // REG_BINARY, safeboot = 0 (Minimal, SafeBoot\\Minimal), Network = 0100000000000000 (SafeBoot\\Network), DsRepair = 0200000000000000 (Directory Services Restore), Unset = not present
    "\\26000081"; "Element" = 01; // REG_BINARY, safebootalternateshell = true, false = 00 - Indicates whether the system should use the shell specified under the following registry key instead of the default shell: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SafeBoot\AlternateShell.
    "\\26000070"; "Element" = 01; // REG_BINARY, usefirmwarepcisettings = true, false = 00 - Indicates whether the system should use I/O and IRQ resources created by the system firmware instead of using dynamically configured resources.
    "\\26000065"; "Element" = 01; // REG_BINARY, groupaware = true, false = 00 - This setting makes drivers group aware and can be used to determine improper group usage.
    "\\26000064"; "Element" = 01; // REG_BINARY, maxgroup = true, false = 00 - Maximizes the number of groups created when assigning nodes to processor groups.
    "\\26000062"; "Element" = 01; // REG_BINARY, maxproc = true, false = 00 - Indicates whether the system should use the maximum number of processors.
    "\\26000060"; "Element" = 01; // REG_BINARY, onecpu = true, false = 00 - Indicates whether the operating system should initialize or start non-boot processors.
    "\\26000054"; "Element" = 01; // REG_BINARY, uselegacyapicmode = true (forces X2APICPOLICY=DISABLE), false = 00 (no-op) - Used to force legacy APIC mode, even if the processors and chipset support extended APIC mode.
    "\\26000051"; "Element" = 01; // REG_BINARY, usephysicaldestination = true, false = 00 - Indicates whether to enable physical-destination mode for all APIC messages.
    "\\26000043"; "Element" = 01; // REG_BINARY, novga = true, false = 00 - Disables the use of VGA modes in the OS.
    "\\26000042"; "Element" = 01; // REG_BINARY, novesa = true, false = 00 - Indicates whether the VGA driver should avoid VESA BIOS calls.
    "\\26000041"; "Element" = 01; // REG_BINARY, quietboot = true, false = 00 - Indicates whether the system should initialize the VGA driver responsible for displaying simple graphics during the boot process. If not, there is no display is presented during the boot process.
    "\\26000040"; "Element" = 01; // REG_BINARY, vga = true, false = 00 - Indicates whether the system should use the standard VGA display driver instead of a high-performance display driver.
    "\\26000030"; "Element" = 01; // REG_BINARY, nolowmem = true, false = 00 - Indicates whether the system should utilize the first 4GB of physical memory. This option requires 5GB of physical memory, and on x86 systems it requires PAE to be enabled.
    "\\26000027"; "Element" = 01; // REG_BINARY, allowprereleasesignatures = true, false = 00 - Indicates whether the test code signing certificate is supported.
    "\\26000025"; "Element" = 01; // REG_BINARY, lastknowngood = true, false = 00 - Indicates that the system should use the last-known good settings.
    "\\26000024"; "Element" = 01; // REG_BINARY, nocrashautoreboot = true, false = 00 - Indicates that the system should not automatically reboot when it crashes.
    "\\26000010"; "Element" = 01; // REG_BINARY, detectkernelandhal = true, false = 00 - Indicates whether the operating system loader should determine the kernel and HAL to load based on the platform features.
    "\\26000004"; "Element" = 01; // REG_BINARY, stampdisks = true, false = 00
    "\\25000142"; "Element" = 0100000000000000; // REG_BINARY, vsmlaunchtype = 1 (auto), Off = 0000000000000000
    "\\25000130"; "Element" = 0000000000000000; // REG_BINARY, claimedtpmcounter = 0
    "\\2500012B"; "Element" = 0100000000000000; // REG_BINARY, xsavedisable = 1, 0 = 0000000000000000 - When set to a value other than zero (0), disables XSAVE functionality in the kernel.
    "\\2500012A"; "Element" = 0000000000000000; // REG_BINARY, xsaveprocessorsmask = 0
    "\\25000129"; "Element" = 0000000000000000; // REG_BINARY, xsaveremovefeature = 0
    "\\25000128"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature7 = 0
    "\\25000127"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature6 = 0
    "\\25000126"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature5 = 0
    "\\25000125"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature4 = 0
    "\\25000124"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature3 = 0
    "\\25000123"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature2 = 0
    "\\25000122"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature1 = 0
    "\\25000121"; "Element" = 0000000000000000; // REG_BINARY, xsaveaddfeature0 = 0
    "\\25000120"; "Element" = 0000000000000000; // REG_BINARY, xsavepolicy = 0
    "\\25000115"; "Element" = 0100000000000000; // REG_BINARY, hypervisoriommupolicy = 1 (enable), Default = 0000000000000000, Disable = 0200000000000000 - Controls whether the hypervisor uses an Input Output Memory Management Unit (IOMMU).
    "\\25000113"; "Element" = 0200000000000000; // REG_BINARY, hypervisorrootproc = 2
    "\\25000100"; "Element" = 0200000000000000; // REG_BINARY, tpmbootentropy = 2 (forceenable), Default = 0000000000000000, ForceDisable = 0100000000000000 - Determines whether entropy is gathered from the trusted platform module (TPM) to help seed the random number generator in the OS.
    "\\250000FB"; "Element" = 0100000000000000; // REG_BINARY, hypervisorrootprocpernode = 1 - Specifies the total number of virtual processors in the root partition that can be started within a pre-split Non-Uniform Memory Architecture (NUMA) node.
    "\\250000FA"; "Element" = 0200000000000000; // REG_BINARY, hypervisornumproc = 2 - Specifies the total number of logical processors that can be started in the hypervisor.
    "\\250000F7"; "Element" = 0000000000000000; // REG_BINARY, bootuxpolicy = 0 (Disabled), Basic = 0100000000000000, Standard = 0200000000000000 (defunct) - Values are Disabled (0), Basic (1), and Standard (2).
    "\\250000F0"; "Element" = 0000000000000000; // REG_BINARY, hypervisorlaunchtype = 0 (Off), Auto = 0100000000000000 - Controls the hypervisor launch type. Options are HyperVisorLaunchOff (0) and HypervisorLaunchAuto (1).
    "\\250000E0"; "Element" = 0000000000000000; // REG_BINARY, bootstatuspolicy = 0 (displayallfailures), IgnoreAllFailures = 0100000000000000, IgnoreBootFailures = 0300000000000000, IgnoreCheckpointFailures = 0400000000000000, IgnoreShutdownFailures = 0200000000000000, DisplayBootFailures = 0600000000000000, DisplayCheckpointFailures = 0700000000000000, DisplayShutdownFailures = 0500000000000000
    "\\250000C2"; "Element" = 0000000000000000; // REG_BINARY, bootmenupolicy = 0 (Legacy), Standard = 0100000000000000 - Defines the type of boot menu the system will use. For Windows 10/11, Windows 8.1, Windows 8 and Windows RT the default is Standard. For Windows Server 2012 R2, Windows Server 2012, the default is Legacy. When Legacy is selected, the Advanced options menu (F8) is available. When Standard is selected, the boot menu appears but only under certain conditions: for example, if there is a startup failure, if you are booting up from a repair disk or installation media, if you have configured multiple boot entries, or if you manually configured the computer to use Advanced startup. When Standard is selected, the F8 key is ignored during boot. - Defines the type of boot menus the system will use. Possible values include menupolicylegacy (0) or menupolicystandard (1).
    "\\250000C1"; "Element" = 0000000000000000; // REG_BINARY, driverloadfailurepolicy = 0 (fatal), UseErrorControl = 0100000000000000 - Indicates the driver load failure policy. Zero (0) indicates that a failed driver load is fatal and the boot will not continue, one (1) indicates that the standard error control is used.
    "\\250000A6"; "Element" = 0100000000000000; // REG_BINARY, tscsyncpolicy = 1 (legacy), Default = 0000000000000000, Enhanced = 0200000000000000 (HalpTscSyncPolicy symbol not present means this doesn't do anything), this should exist on older Windows versions and controls the TSC synchronization policy
    "\\25000072"; "Element" = 0100000000000000; // REG_BINARY, pciexpress = 1 (forcedisable), Default = 0000000000000000
    "\\25000071"; "Element" = 0100000000000000; // REG_BINARY, msi = 1 (forcedisable), Default = 0000000000000000, ForceEnable only via loadoptions FORCEMSI - The PCI Message Signaled Interrupt (MSI) policy. Zero (0) indicates default, and one (1) indicates that MSI interrupts are disabled.
    "\\25000066"; "Element" = 4000000000000000; // REG_BINARY, groupsize = 64 - Specifies the size of all processor groups. Must be set to a power of 2 (max of 64, see pseudocode below).
    "\\25000063"; "Element" = 0100000000000000; // REG_BINARY, configflags = 1 - Indicates whether processor specific configuration flags are to be used.
    "\\25000061"; "Element" = 0200000000000000; // REG_BINARY, numproc = 2 - The maximum number of processors that can be utilized by the system; all other processors are ignored.
    "\\25000055"; "Element" = 0200000000000000; // REG_BINARY, x2apicpolicy = 2 (enable), Default = 0000000000000000, Disable = 0100000000000000 - Enables the use of extended APIC mode, if supported. Zero (0) indicates default behavior, one (1) indicates that extended APIC mode is disabled, and two (2) indicates that extended APIC mode is enabled. The system defaults to using extended APIC mode if available.
    "\\25000050"; "Element" = 0100000000000000; // REG_BINARY, clustermodeaddressing = 1 - Indicates that cluster-mode APIC addressing should be utilized, and the value is the maximum number of processors per cluster.
    "\\25000052"; "Element" = 0000000000000000; // REG_BINARY, restrictapicluster = 0 - The maximum number of APIC clusters that should be used by cluster-mode addressing.
    "\\25000032"; "Element" = 0004000000000000; // REG_BINARY, increaseuserva = 1024 - Increasing this value from the default 2GB decreases the amount of virtual address space available to the system and device drivers. The amount of memory that should be utilized by the process address space, in bytes. This value should be between 2GB and 3GB.
    "\\25000033"; "Element" = 0000000000000000; // REG_BINARY, perfmem = 0 - BcdOSLoaderInteger_PerformaceDataMemory (integer)
    "\\25000031"; "Element" = 8000000000000000; // REG_BINARY, removememory = 128 - The amount of memory the system should ignore.
    "\\25000021"; "Element" = 0100000000000000; // REG_BINARY, pae = 1 (forceenable), Default = 0000000000000000, ForceDisable = 0200000000000000 - If this value is not specified, the default is PaePolicyDefault which follows the rule "enable PAE if hot-pluggable memory is above 4GB"
    "\\25000020"; "Element" = 0000000000000000; // REG_BINARY, nx = 0 (OptIn, NX off by default), OptOut = 0100000000000000 (NX on by default), AlwaysOff = 0200000000000000, AlwaysOn = 0300000000000000 - If this value is not specified, the default is NxPolicyAlwaysOff.
    "\\23000003"; "Element" = {resume}; // REG_SZ, resumeobject = {resume} - The default boot environment application to load if the user does not select one.
    "\\22000053"; "Element" = \EFI\Microsoft\Boot\EVStore.dat; // REG_SZ, evstore = \EFI\Microsoft\Boot\EVStore.dat
    "\\22000041"; "Element" = recovery message; // REG_SZ, fverecoverymessage = recovery message
    "\\22000040"; "Element" = https://example.com/recovery; // REG_SZ, fverecoveryurl = https://example.com/recovery
    "\\22000013"; "Element" = kdcom.dll; // REG_SZ, dbgtransport = kdcom.dll - The transport DLL to be loaded by the operating system loader. This value overrides the default Kdcom.dll.
    "\\22000012"; "Element" = hal.dll; // REG_SZ, hal = hal.dll - The HAL to be loaded by the operating system loader. This value overrides the default HAL.
    "\\22000011"; "Element" = ntoskrnl.exe; // REG_SZ, kernel = ntoskrnl.exe - The kernel to be loaded by the operating system loader. This value overrides the default kernel.
    "\\22000002"; "Element" = \Windows; // REG_SZ, systemroot = \Windows - This value is reserved.
    "\\21000001"; "Element" = partition=C:; // REG_BINARY, filedevice = partition=C: - This value is reserved.
    "\\17000077"; "Element" = 7500001500000000; // REG_BINARY, allowedinmemorysettings = 0x15000075 - Indicates whether or not an in-memory BCD setting passed between boot apps will trigger BitLocker recovery. This value should not be modified as it could trigger a BitLocker recovery action.
    "\\1600007B"; "Element" = 01; // REG_BINARY, forcefipscrypto = true, false = 00 (BitLocker validation profile)
    "\\16000079"; "Element" = 01; // REG_BINARY, forcefipscrypto = true, false = 00 (BCD library enum, the one above is probably the used one) - Force the use of FIPS cryptography checks on boot applications.
    "\\16000074"; "Element" = 01; // REG_BINARY, bootshutdowndisabled = true, false = 00 - Disables the 1-minute timer that triggers shutdown on boot error screens, and the F8 menu, on UEFI systems.
    "\\16000072"; "Element" = 01; // REG_BINARY, nokeyboard = true, false = 00
    "\\1600006C"; "Element" = 01; // REG_BINARY, bootuxdisabled = true, false = 00 - This setting disables the progress bar and default Windows logo. If a custom text string has been defined, it is also disabled by this setting.
    "\\16000069"; "Element" = 01; // REG_BINARY, custom:16000069 = true, false = 00 - disables the loading circle while booting (see image at the bottom)
    "\\16000067"; "Element" = 01; // REG_BINARY, custom:16000067 = true, false = 00 - disables the Windows logo while booting (see image at the bottom)
    "\\16000060"; "Element" = 01; // REG_BINARY, isolatedcontext = true, false = 00 - Do not modify this setting. If this setting is removed from a Windows 8 installation, it will not boot. If this setting is added to a Windows 7 installation, it will not boot. - This setting is used to differentiate between the Windows 7 and Windows 8 implementations of UEFI.
    "\\16000054"; "Element" = 01; // REG_BINARY, highestmode = true, false = 00 - Forces highest available graphics resolution at boot. This value can only be used on UEFI systems.
    "\\16000053"; "Element" = 01; // REG_BINARY, restartonfailure = true, false = 00 - If enabled, specifies that boot error screens are not shown when OS launch errors occur, and the system is reset rather than exiting directly back to the firmware.
    "\\16000050"; "Element" = 01; // REG_BINARY, consoleextendedinput = true, false = 00 - Specifies that legacy BIOS systems should use INT 16h Function 10h for console input instead of INT 16h Function 0h.
    "\\16000049"; "Element" = 01; // REG_BINARY, testsigning = true, false = 00 - Indicates whether the test code signing certificate is supported.
    "\\16000048"; "Element" = 01; // REG_BINARY, nointegritychecks = true, false = 00 - This value is ignored by Windows 7 and Windows 8. - Disables integrity checks. Cannot be set when secure boot is enabled.
    "\\16000046"; "Element" = 01; // REG_BINARY, graphicsmodedisabled = true, false = 00 - Indicates whether graphics mode is disabled and boot applications must use text mode display.
    "\\16000041"; "Element" = 01; // REG_BINARY, optionsedit = true, false = 00 - Indicates whether the boot options editor is enabled.
    "\\16000040"; "Element" = 01; // REG_BINARY, advancedoptions = true, false = 00 - Indicates whether the advanced options boot menu (F8) is displayed.
    "\\1600001E"; "Element" = 01; // REG_BINARY, vm = true, false = 00
    "\\1600000F"; "Element" = 01; // REG_BINARY, traditionalkseg = true, false = 00
    "\\16000009"; "Element" = 01; // REG_BINARY, recoveryenabled = true, false = 00 - Indicates whether the recovery sequence executes automatically if the boot application fails. Otherwise, the recovery sequence only runs on demand.
    "\\15000081"; "Element" = 0000000000000000; // REG_BINARY, logcontrol = 0
    "\\15000088"; "Element" = 0000000000000000; // REG_BINARY, linearaddress57 = 0 (Default), OptOut = 0100000000000000, OptIn = 0200000000000000
    "\\15000066"; "Element" = 0300000000000000; // REG_BINARY, displaymessageoverride = 3 (Recovery), Resume = 0100000000000000
    "\\15000052"; "Element" = 0000000000000000; // REG_BINARY, graphicsresolution = 0 (1024x768), 800x600 = 0100000000000000, 1024x600 = 0200000000000000 - Forces a specific graphics resolution at boot. Possible values include GraphicsResolution1024x768 (0), GraphicsResolution800x600 (1), and GraphicsResolution1024x600 (2).
    "\\15000051"; "Element" = 0000000000000000; // REG_BINARY, initialconsoleinput = 0
    "\\1500004C"; "Element" = 0000000000000000; // REG_BINARY, volumebandid = 0 (fvebandid) - This value (if present) should not be modified.
    "\\1500004B"; "Element" = 0000000000000000; // REG_BINARY, integrityservices = 0 (Default, Enabled - Kernel Mode Code Signing), Enable = 0100000000000000, Disable = 0200000000000000
    "\\15000047"; "Element" = 0000000000000000; // REG_BINARY, configaccesspolicy = 0 (Default, allow MMCONFIG), DisallowMmConfig = 0100000000000000 (use CF8/CFC instead) - Indicates the access policy for PCI configuration space.
    "\\15000042"; "Element" = 0000000000000000; // REG_BINARY, keyringaddress = 0
    "\\1500000E"; "Element" = 0000000000000000; // REG_BINARY, avoidlowphysicalmemory = 0 - Specifies a minimum physical address to use in the boot environment.
    "\\1500000D"; "Element" = 0000000000000000; // REG_BINARY, relocatephysicalmemory = 0 - This value is not used in Windows 8 or Windows Server 2012. - Relocates physical memory on certain AMD processors.
    "\\1500000C"; "Element" = 0000000000000000; // REG_BINARY, firstmegabytepolicy = 0 (UseNone, use none of first MB), UseAll = 0100000000000000 (use all of first MB), UsePrivate = 0200000000000000 (reserved) - Indicates how the first megabyte of memory is to be used.
    "\\15000007"; "Element" = 0000008000000000; // REG_BINARY, truncatememory = 2147483648 - Maximum physical address a boot environment application should recognize. All memory above this address is ignored.
    "\\14000008"; "Element" = {winre}; // REG_MULTI_SZ, recoverysequence = {winre} - List of boot environment applications to be executed if the associated application fails. The applications are executed in the order they appear in this list.
    "\\14000006"; "Element" = {bootloadersettings}; // REG_MULTI_SZ, inherit = {bootloadersettings} - List of BCD objects from which the current object should inherit elements.
    "\\1200004A"; "Element" = \Windows\Fonts; // REG_SZ, fontpath = \Windows\Fonts - Use caution when modifying this setting. Boot screens will not work if the correct fonts are not present. - Overrides the default location of the boot fonts.
    "\\12000044"; "Element" = \Boot\BCD-Log; // REG_SZ, bsdlogpath = \Boot\BCD-Log - Allows a path override for the bootstat.dat log file in the boot manager and winload.exe.
    "\\12000030"; "Element" = NOCRASHONCTRL; // REG_SZ, loadoptions = NOCRASHONCTRL - String that is appended to the load options string passed to the kernel to be consumed by kernel-mode components. This is useful for communicating with kernel-mode components that are not BCD-aware.
    "\\12000005"; "Element" = en-US; // REG_SZ, locale = en-US - Preferred locale, in RFC 3066 format.
    "\\12000004"; "Element" = Windows 11; // REG_SZ, description = Windows 11 - Display name of the boot environment application.
    "\\12000002"; "Element" = \Windows\system32\winload.efi; // REG_SZ, path = \Windows\system32\winload.efi - Path to a boot environment application.
    "\\11000043"; "Element" = partition=C:; // REG_BINARY, bsdlogdevice = partition=C: - Allows a device override for the bootstat.dat log in the boot manager and winload.exe.
    "\\11000001"; "Element" = partition=C:; // REG_BINARY, device = partition=C: - Device on which a boot environment application resides.
"HKLM\\BCD00000000\\Objects\\{resume}\\Elements";
    "\\26000006"; "Element" = 00; // REG_BINARY, debugoptionenabled = false, true = 01 - Enables kernel debugging on resume from hibernate.
    "\\26000004"; "Element" = 01; // REG_BINARY, pae = true, false = 00
    "\\26000003"; "Element" = 01; // REG_BINARY, usecustomsettings = true, false = 00 - Allows the resume loader BCD object to use custom settings. If this setting is not specified or is not enabled, default settings are applied by the OS before resume.
    "\\25000008"; "Element" = 0100000000000000; // REG_BINARY, bootmenupolicy = 1 (Standard), Legacy = 0000000000000000 - Defines the type of boot menus the system will use. Possible values are menupolicylegacy (0) or menupolicystandard (1). The default setting is menupolicylegacy (0).
    "\\25000007"; "Element" = 0000000000000000; // REG_BINARY, bootux = 0 (Disabled), Basic = 0100000000000000, Standard = 0200000000000000 (defunct)
    "\\22000002"; "Element" = \hiberfil.sys; // REG_SZ, filepath = \hiberfil.sys - This value is reserved.
    "\\21000026"; "Element" = partition=C:; // REG_BINARY, custom:21000026 = partition=C:
    "\\21000005"; "Element" = partition=C:; // REG_BINARY, associatedosdevice = partition=C: - Specifies the name of the OS device associated with the hibernated OS. This is only used if the hibernation file is not stored on the OS device.
    "\\21000001"; "Element" = partition=C:; // REG_BINARY, filedevice = partition=C: - This value is reserved.
    "\\17000077"; "Element" = 7500001500000000; // REG_BINARY, allowedinmemorysettings = 0x15000075 - Indicates whether or not an in-memory BCD setting passed between boot apps will trigger BitLocker recovery. This value should not be modified as it could trigger a BitLocker recovery action.
    "\\16000060"; "Element" = 01; // REG_BINARY, isolatedcontext = true, false = 00 - Do not modify this setting. If this setting is removed from a Windows 8 installation, it will not boot. If this setting is added to a Windows 7 installation, it will not boot. - This setting is used to differentiate between the Windows 7 and Windows 8 implementations of UEFI.
    "\\16000009"; "Element" = 01; // REG_BINARY, recoveryenabled = true, false = 00 - Indicates whether the recovery sequence executes automatically if the boot application fails. Otherwise, the recovery sequence only runs on demand.
    "\\14000008"; "Element" = {winre}; // REG_MULTI_SZ, recoverysequence = {winre} - List of boot environment applications to be executed if the associated application fails. The applications are executed in the order they appear in this list.
    "\\14000006"; "Element" = {resumeloadersettings}; // REG_MULTI_SZ, inherit = {resumeloadersettings} - List of BCD objects from which the current object should inherit elements.
    "\\12000005"; "Element" = en-US; // REG_SZ, locale = en-US - Preferred locale, in RFC 3066 format.
    "\\12000004"; "Element" = Windows Resume Application; // REG_SZ, description = Windows Resume Application - Display name of the boot environment application.
    "\\12000002"; "Element" = \Windows\system32\winresume.efi; // REG_SZ, path = \Windows\system32\winresume.efi - Path to a boot environment application.
    "\\11000001"; "Element" = partition=C:; // REG_BINARY, device = partition=C: - Device on which a boot environment application resides.
"HKLM\\BCD00000000\\Objects\\{bootmgr}\\Elements";
    "\\27000030"; "Element" = 0000000000000000; // REG_BINARY, customactionslist = <integer list> - For more information see Custom Bootstrap Actions in Windows Vista.
    "\\26000031"; "Element" = 01; // REG_BINARY, persistbootsequence = true, false = 00 - Controls whether a boot sequence persists across multiple boots.
    "\\26000028"; "Element" = 01; // REG_BINARY, processcustomactionsfirst = true, false = 00 - Controls whether custom actions are processed before a boot sequence.
    "\\26000021"; "Element" = 01; // REG_BINARY, noerrordisplay = true, false = 00 - Indicates whether the display of errors should be suppressed. If this setting is enabled, the boot manager exits to the multi-OS menu on OS launch error.
    "\\26000020"; "Element" = 01; // REG_BINARY, displaybootmenu = true, false = 00 - Forces the display of the legacy boot menu, regardless of the number of OS entries in the BCD store and their BcdOSLoaderInteger_BootMenuPolicy.
    "\\26000005"; "Element" = 01; // REG_BINARY, attemptresume = true, false = 00 - Indicates that a resume operation should be attempted during a system restart.
    "\\25000004"; "Element" = 0300000000000000; // REG_BINARY, timeout = 3 - The boot menu time-out determines how long the boot menu is displayed before the default boot entry is loaded. It is calibrated in seconds. If you want extra time to choose the operating system that loads on your computer, you can extend the time-out value. Or, you can shorten the time-out value so that the default operating system starts faster. - If this value is not specified, the boot manager waits for the user to make a selection. - The maximum number of seconds a boot selection menu is to be displayed to the user. The menu is displayed until the user selects an option or the time-out expires.
    "\\24000010"; "Element" = {memdiag}; // REG_MULTI_SZ, toolsdisplayorder = {memdiag} - The boot manager tools display order list.
    "\\24000002"; "Element" = {current}; // REG_MULTI_SZ, bootsequence = {current} - If the firmware boot manager does not support loading multiple applications, this list cannot contain more than one entry. - List of boot environment applications the boot manager should execute. The applications are executed in the order they appear in this list.
    "\\24000001"; "Element" = {current}; // REG_MULTI_SZ, displayorder = {current} - The order in which BCD objects should be displayed. Objects are displayed using the string specified by the BcdLibraryString_Description element.
    "\\23000006"; "Element" = {resume}; // REG_SZ, resumeobject = {resume} - The resume application object.
    "\\23000003"; "Element" = {current}; // REG_SZ, default = {current} - The default boot environment application to load if the user does not select one.
    "\\22000023"; "Element" = \EFI\Microsoft\Boot\BCD; // REG_SZ, bcdfilepath = \EFI\Microsoft\Boot\BCD - The boot application.
    "\\21000022"; "Element" = partition=\Device\HarddiskVolume1; // REG_BINARY, bcddevice = partition=\Device\HarddiskVolume1 - The device on which the boot application resides.
    "\\14000006"; "Element" = {globalsettings}; // REG_MULTI_SZ, inherit = {globalsettings} - List of BCD objects from which the current object should inherit elements.
    "\\12000005"; "Element" = en-US; // REG_SZ, locale = en-US - Preferred locale, in RFC 3066 format.
    "\\12000004"; "Element" = Windows Boot Manager; // REG_SZ, description = Windows Boot Manager - Display name of the boot environment application.
    "\\12000002"; "Element" = \EFI\MICROSOFT\BOOT\BOOTMGFW.EFI; // REG_SZ, path = \EFI\MICROSOFT\BOOT\BOOTMGFW.EFI - Path to a boot environment application.
    "\\11000001"; "Element" = partition=\Device\HarddiskVolume1; // REG_BINARY, device = partition=\Device\HarddiskVolume1 - Device on which a boot environment application resides.
"HKLM\\BCD00000000\\Objects\\{memdiag}\\Elements";
    "\\26000004"; "Element" = 01; // REG_BINARY, failuresenabled = true, false = 00
    "\\25000009"; "Element" = 0000000000000000; // REG_BINARY, chckrfailcount = 0
    "\\25000007"; "Element" = 0000000000000000; // REG_BINARY, matsfailcount = 0
    "\\25000006"; "Element" = 0000000000000000; // REG_BINARY, invcfailcount = 0
    "\\25000005"; "Element" = 0000000000000000; // REG_BINARY, stridefailcount = 0
    "\\25000003"; "Element" = 0000000000000000; // REG_BINARY, failurecount = 0 - The number of pages that contain errors. This is useful for simulating error flows in the absence of bad physical memory.
    "\\25000002"; "Element" = 0000000000000000; // REG_BINARY, testmix = 0
    "\\25000001"; "Element" = 0000000000000000; // REG_BINARY, passcount = 0 - If this value is not specified, the default is to run memory diagnostic tests until the computer is powered off or the user logs off. - The number of passes for the current test mix.
    "\\1600000B"; "Element" = 01; // REG_BINARY, badmemoryaccess = true, false = 00 - If TRUE, indicates that a boot application can use memory listed in the BcdLibraryIntegerList_BadMemoryList.
    "\\14000006"; "Element" = {globalsettings}; // REG_MULTI_SZ, inherit = {globalsettings} - List of BCD objects from which the current object should inherit elements.
    "\\12000005"; "Element" = en-US; // REG_SZ, locale = en-US - Preferred locale, in RFC 3066 format.
    "\\12000004"; "Element" = Windows Memory Diagnostic; // REG_SZ, description = Windows Memory Diagnostic - Display name of the boot environment application.
    "\\12000002"; "Element" = \EFI\Microsoft\Boot\memtest.efi; // REG_SZ, path = \EFI\Microsoft\Boot\memtest.efi - Path to a boot environment application.
    "\\11000001"; "Element" = partition=\Device\HarddiskVolume1; // REG_BINARY, device = partition=\Device\HarddiskVolume1 - Device on which a boot environment application resides.
"HKLM\\BCD00000000\\Objects\\{badmemory}\\Elements";
    "\\1700000A"; "Element" = 0000000000000000; // REG_BINARY, badmemorylist = <integer list> - List of page frame numbers describing faulty memory in the system.
"HKLM\\BCD00000000\\Objects\\{winre}\\Elements";
    "\\46000010"; "Element" = 01; // REG_BINARY, custom:46000010 = true, false = 00
    "\\26000022"; "Element" = 01; // REG_BINARY, winpe = true, false = 00 - Indicates that the system should be started in Windows Preinstallation Environment (Windows PE) mode.
    "\\250000C2"; "Element" = 0100000000000000; // REG_BINARY, bootmenupolicy = 1 (Standard), Legacy = 0000000000000000 - Defines the type of boot menus the system will use. Possible values include menupolicylegacy (0) or menupolicystandard (1). The default value is menupolicylegacy (0).
    "\\25000020"; "Element" = 0000000000000000; // REG_BINARY, nx = 0 (OptIn), OptOut = 0100000000000000, AlwaysOff = 0200000000000000, AlwaysOn = 0300000000000000 - If this value is not specified, the default is NxPolicyAlwaysOff. - The no-execute page protection policy.
    "\\22000002"; "Element" = \windows; // REG_SZ, systemroot = \windows - This value is reserved.
    "\\21000001"; "Element" = ramdisk=[C:]\Recovery\WindowsRE\Winre.wim,{ramdiskoptions}; // REG_BINARY, osdevice = ramdisk=[C:]\Recovery\WindowsRE\Winre.wim,{ramdiskoptions} - This value is reserved.
    "\\15000065"; "Element" = 0300000000000000; // REG_BINARY, displaymessage = 3 (Recovery), Resume = 0100000000000000
    "\\14000006"; "Element" = {bootloadersettings}; // REG_MULTI_SZ, inherit = {bootloadersettings} - List of BCD objects from which the current object should inherit elements.
    "\\12000005"; "Element" = en-us; // REG_SZ, locale = en-us - Preferred locale, in RFC 3066 format.
    "\\12000004"; "Element" = Windows Recovery Environment; // REG_SZ, description = Windows Recovery Environment - Display name of the boot environment application.
    "\\12000002"; "Element" = \windows\system32\winload.efi; // REG_SZ, path = \windows\system32\winload.efi - Path to a boot environment application.
    "\\11000001"; "Element" = ramdisk=[C:]\Recovery\WindowsRE\Winre.wim,{ramdiskoptions}; // REG_BINARY, device = ramdisk=[C:]\Recovery\WindowsRE\Winre.wim,{ramdiskoptions} - Device on which a boot environment application resides.
"HKLM\\BCD00000000\\Objects\\{ramdiskoptions}\\Elements";
    "\\3600000B"; "Element" = 01; // REG_BINARY, ramdisktftpvarwindow = true, false = 00 - Enables or disables the TFTP variable window size extension.
    "\\3600000A"; "Element" = 01; // REG_BINARY, ramdiskmulticasttftpfallback = true, false = 00 (ramdiskmctftpfallback) - Enables fallback to TFTP if multicast fails.
    "\\36000009"; "Element" = 01; // REG_BINARY, ramdiskmulticastenabled = true, false = 00 (ramdiskmcenabled) - Enables or disables multicast for the RAM disk WIM file.
    "\\36000008"; "Element" = 0000000000000000; // REG_BINARY, ramdisktftpwindowsize = 0 - Defines the TFTP window size for the RAM disk WIM file.
    "\\36000007"; "Element" = 0000000000000000; // REG_BINARY, ramdisktftpblocksize = 0 - Defines the TFTP block size for the RAM disk Windows Imaging (WIM) file.
    "\\36000006"; "Element" = 01; // REG_BINARY, ramdiskexportascd = true, false = 00 - Enables exporting the RAM disk as a CD.
    "\\35000008"; "Element" = 0000000000000000; // REG_BINARY, ramdisktftpwindowsize = 0 (BitLocker list uses 0x35000008)
    "\\35000007"; "Element" = 0000000000000000; // REG_BINARY, ramdisktftpblocksize = 0 (BitLocker list uses 0x35000007)
    "\\35000005"; "Element" = 0000000000000000; // REG_BINARY, ramdiskimagelength = 0 - The length of the image for the RAM disk.
    "\\35000002"; "Element" = 0000000000000000; // REG_BINARY, tftpclientport = 0 - If this value is not specified, the default TFTP protocol port is used. - The IP port number to be used for Trivial File Transfer Protocol (TFTP) reads.
    "\\35000001"; "Element" = 0000000000000000; // REG_BINARY, ramdiskimageoffset = 0 - The RAM disk image offset.
    "\\32000004"; "Element" = \Recovery\WindowsRE\boot.sdi; // REG_SZ, ramdisksdipath = \Recovery\WindowsRE\boot.sdi - The path from the root of the SDI device to the RAM disk file.
    "\\31000003"; "Element" = partition=C:; // REG_BINARY, ramdisksdidevice = partition=C: - The device that contains the SDI object.
    "\\12000004"; "Element" = Windows Recovery; // REG_SZ, description = Windows Recovery - Display name of the boot environment application.
"HKLM\\BCD00000000\\Objects\\{globalsettings}\\Elements";
    "\\16000069"; "Element" = 01; // REG_BINARY, custom:16000069 = true, false = 00
    "\\16000067"; "Element" = 01; // REG_BINARY, custom:16000067 = true, false = 00
    "\\14000006"; "Element" = {dbgsettings};{emssettings};{badmemory}; // REG_MULTI_SZ, inherit = {dbgsettings}, {emssettings}, {badmemory} - List of BCD objects from which the current object should inherit elements.
"HKLM\\BCD00000000\\Objects\\{resumeloadersettings}\\Elements";
    "\\14000006"; "Element" = {globalsettings}; // REG_MULTI_SZ, inherit = {globalsettings} - List of BCD objects from which the current object should inherit elements.
"HKLM\\BCD00000000\\Objects\\{bootloadersettings}\\Elements";
    "\\14000006"; "Element" = {globalsettings};{hypervisorsettings}; // REG_MULTI_SZ, inherit = {globalsettings}, {hypervisorsettings} - List of BCD objects from which the current object should inherit elements.
"HKLM\\BCD00000000\\Objects\\{dbgsettings}\\Elements";
    "\\1600001C"; "Element" = 01; // REG_BINARY, debuggernetdhcp = true, false = 00 - Controls the use of DHCP by the network debugger. Setting this to false causes the OS to only use link-local addresses.
    "\\16000017"; "Element" = 01; // REG_BINARY, debuggerignoreusermodeexceptions = true, false = 00 - If TRUE, the debugger will ignore user mode exceptions and only stop for kernel mode exceptions.
    "\\16000010"; "Element" = 01; // REG_BINARY, debuggerenabled = true, false = 00 - Indicates whether the boot debugger should be enabled.
    "\\1500001B"; "Element" = 0000000000000000; // REG_BINARY, debuggernetport = 0 - Defines the network port for the network debugger.
    "\\1500001A"; "Element" = 0000000000000000; // REG_BINARY, debuggernethostip = 0 - Defines the host IP address for the network debugger.
    "\\15000018"; "Element" = 0000000000000000; // REG_BINARY, debuggerstartpolicy = 0 - Indicates the debugger start policy.
    "\\15000015"; "Element" = 0000000000000000; // REG_BINARY, debugger1394channel = 0 - Channel number for 1394 debugging.
    "\\15000014"; "Element" = 00c2010000000000; // REG_BINARY, debuggerserialbaudrate = 115200 - If this value is not specified, the default is specified by the DBGP ACPI table settings. - Baud rate for serial debugging.
    "\\15000013"; "Element" = 0100000000000000; // REG_BINARY, debuggerserialport = 1 - If this value is not specified, the default is specified by the DBGP ACPI table settings. - Serial port number for serial debugging.
    "\\15000012"; "Element" = 0000000000000000; // REG_BINARY, debuggerserialportaddress = 0 - I/O port address for the serial debugger.
    "\\15000011"; "Element" = 0400000000000000; // REG_BINARY, debugtype = 4 (Local - undocumented), Serial = 0000000000000000, 1394 = 0100000000000000, USB = 0200000000000000, NET = 0300000000000000 - Debugger type.
    "\\1200001D"; "Element" = testkey; // REG_SZ, debuggernetkey = testkey - Holds the key used to encrypt the network debug connection.
    "\\12000019"; "Element" = 0.25.0; // REG_SZ, debuggerbusparams = 0.25.0 - Defines the PCI bus, device, and function numbers of the debugging device. For example, 1.5.0 describes the debugging device on bus 1, device 5, function 0.
    "\\12000016"; "Element" = usbtarget; // REG_SZ, debuggerusbtargetname = usbtarget - The target name for the USB debugger. The target name is arbitrary but must match between the debugger and the debug target.
"HKLM\\BCD00000000\\Objects\\{emssettings}\\Elements";
    "\\16000020"; "Element" = 00; // REG_BINARY, bootems = false, true = 01 - Indicates whether EMS redirection should be enabled.
    "\\15000023"; "Element" = 00c2010000000000; // REG_BINARY, emsbaudrate = 115200 - Baud rate for EMS redirection.
    "\\15000022"; "Element" = 0100000000000000; // REG_BINARY, emsport = 1 - If this value is not specified, the default is specified by the SPCR ACPI table settings. - COM port number for EMS redirection.
"HKLM\\BCD00000000\\Objects\\{fwbootmgr}\\Elements";
    "\\25000004"; "Element" = 0100000000000000; // REG_BINARY, timeout = 1 - If this value is not specified, the boot manager waits for the user to make a selection. - The maximum number of seconds a boot selection menu is to be displayed to the user. The menu is displayed until the user selects an option or the time-out expires.
    "\\24000001"; "Element" = {bootmgr}; // REG_MULTI_SZ, displayorder = {bootmgr} - The order in which BCD objects should be displayed. Objects are displayed using the string specified by the BcdLibraryString_Description element.
"HKLM\\BCD00000000\\Objects\\{hypervisorsettings}\\Elements";
    "\\26000114"; "Element" = 00; // REG_BINARY, hypervisordhcp = false, true = 01 - Controls use of DHCP by the network debugger used with the hypervisor. Setting this to false forces local link only address.
    "\\250000FE"; "Element" = 50c3000000000000; // REG_BINARY, hypervisorhostport = 50000 - Defines the network UDP port for the network debugger.
    "\\250000FD"; "Element" = 0201a8c000000000; // REG_BINARY, hypervisorhostip = 3232235778 (192.168.1.2) - Defines the host IPv4 address for the network debugger.
    "\\250000F6"; "Element" = 0000000000000000; // REG_BINARY, hypervisorchannel = 0 - Specifies the channel number for 1394 debugging.
    "\\250000F5"; "Element" = 00c2010000000000; // REG_BINARY, hypervisorbaudrate = 115200 - If this value is not specified, the default is specified by the DBGP ACPI table settings. - Specifies the baud rate for serial debugging.
    "\\250000F4"; "Element" = 0100000000000000; // REG_BINARY, hypervisordebugport = 1 - If this value is not specified, the default is specified by the DBGP ACPI table settings. - Specifies the serial port number for serial debugging.
    "\\250000F3"; "Element" = 0000000000000000; // REG_BINARY, hypervisordebugtype = 0 (Serial), 1394 = 0100000000000000, NET = 0300000000000000 - Controls the hypervisor debugger type. Can be set to SERIAL (0), 1394 (1), or NET (2).
    "\\22000110"; "Element" = testkey; // REG_SZ, hypervisorusekey = testkey - Holds the key used to encrypt the network debug connection used with the hypervisor.
    "\\220000F9"; "Element" = 0.25.0; // REG_SZ, hypervisorbusparams = 0.25.0 - Defines the PCI bus, device, and function numbers of the debugging device used with the hypervisor. For example, 1.5.0 describes the debugging device on bus 1, device 5, function 0.
```

`{bootmgr}` - Windows Boot Manager  
`{fwbootmgr}` - Firmware Boot Manager  
`{current}` - Windows Boot Loader (current OS entry)  
`{resume}` - Windows Resume Application (resumeobject)  
`{winre}` - Windows Recovery Environment loader (recoverysequence)  
`{memdiag}` - Windows Memory Diagnostic  
`{ramdiskoptions}` - Device options for ramdisk (boot.sdi)  
`{globalsettings}` - Global settings  
`{bootloadersettings}` - Boot loader settings  
`{resumeloadersettings}` - Resume loader settings  
`{dbgsettings}` - Debugger settings  
`{emssettings}` - EMS settings  
`{badmemory}` - RAM defects  
`{hypervisorsettings}` - Hypervisor settings

> https://learn.microsoft.com/en-us/previous-versions/windows/desktop/bcd/bcd-enumerations  
> https://learn.microsoft.com/en-us/windows/security/operating-system-security/data-protection/bitlocker/bcd-settings-and-bitlocker

## Intel NIC Values

See [intelnet6x.c](https://github.com/nohuto/win-registry/blob/main/assets/intelnet6x.c) for reference, most of them were found via xrefs of `REGISTRY::RegReadRegTable`. Defaults depends on the adapter/driver, these were found on "Intel(R) Ethernet Controller I225-V". Many of them aren't read ([NIC-Intel.txt](https://github.com/nohuto/win-registry/blob/main/records/NIC-Intel.txt)).

Many parts aren't structered as they should be after decompiling via IDA, which made it impossible to get their data. See [NIC-Intel-IDA.txt](https://github.com/nohuto/win-registry/blob/main/records/NIC-Intel-IDA.txt) for a list of values which I found in IDA (within a single driver). The list below shows values which included their data.

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Class\\{4D36E972-E325-11CE-BFC1-08002bE10318}\\00XX";
    "*DeviceSleepOnDisconnect" = 0; // range 0-1
    "*EnableDynamicPowerGating" = 1; // range 0-1
    "*EncapsulatedPacketTaskOffloadVxlan" = 0; // range 0-1
    "*FlowControl" = 4; // range 0-4
    "*HeaderDataSplit" = 0; // range 0-1
    "*PMARPOffload" = 0; // range 0-1
    "*PMNSOffload" = 0; // range 0-1
    "*ReceiveBuffers" = 512; // range 128-4096
    "*RSCIPv4" = 0; // range 0-1
    "*RSCIPv6" = 0; // range 0-1
    "*RssOrVmqPreference" = 0; // range 0-1
    "*SpeedDuplex" = 0; // range 0-50000
    "*Sriov" = 0; // range 0-1
    "*SriovPreferred" = 0; // range 0-1
    "*VMQ" = 0; // range 0-1
    "*VMQLookaheadSplit" = 0; // range 0-1
    "*VMQVlanFiltering" = 1; // range 0-1
    "*VxlanUDPPortNumber" = 4789; // range 1-65535
    "*WakeOnMagicPacket" = 1; // range 0-1
    "*WakeOnPattern" = 1; // range 0-1
    "AdaptiveQHysteresis" = 64; // range 16-1024
    "AdaptiveQSize" = 128; // range 64-8192
    "AdaptiveQWorkSet" = 96; // range 32-8192
    "CheckForHangTime" = 2; // range 0-60
    "DisableIntelRST" = 1; // range 0-1
    "DisableReset" = 0; // range 0-1
    "DMACoalescing" = 0; // range 0-10240
    "EnableAdaptiveQueuing" = 1; // range 0-1
    "EnableDisconnectedStandby" = 0; // range 0-1
    "EnableHWAutonomous" = 0; // range 0-1
    "EnableModernStandby" = 0; // range 0-1
    "EnablePME" = 0; // range 0-1
    "EnablePowerManagement" = 1; // range 0-1
    "EnableRxDescriptorChaining" = 1; // range 0-1
    "FecMode" = 0; // range 0-3
    "ForceHostExitUlp" = 0; // range 0-1
    "ForceLtrValue" = 65535; // range 0-65535
    "ForceRscEnabled" = 0; // range 0-1
    "HDSplitAlways" = 0; // range 0-1
    "HDSplitBufferPad" = 2; // range 0-2
    "HDSplitLocation" = 2; // range 0-3
    "HDSplitSize" = 128; // range 128-960
    "I218DisablePLLShut" = 0; // range 0-1
    "I218DisablePLLShutGiga" = 0; // range 0-1
    "I219DisableK1Off" = 0; // range 0-1
    "MaxPacketCountPerDPC" = 256; // range 8-65535
    "MaxPacketCountPerIndicate" = 64; // range 1-65535
    "MinHardwareOwnedPacketCount" = 32; // range 8-4096
    "PadReceiveBuffer" = 0; // range 0-1
    "ReceiveBuffersOverride" = 1; // range 0-1
    "RegForceRxPathSerialization" = 0; // range 0-1
    "ResetTest" = 0; // range 0-1
    "ResetTestTime" = 300; // range 20-604800
    "RscMode" = 1; // range 0-2
    "RxBufferPad" = 10; // range 0-63
    "RxDescriptorCountPerTailWrite" = 8; // range 4-4096
    "SidebandUngateOverride" = 0; // range 0-1
    "StoreBadPackets" = 0; // range 0-1
    "ULPMode" = 1; // range 0-1
    "VMQSupported" = 0; // range 0-1
    "WakeFromS5" = 2; // range 0-65535
    "WakeOn" = 0; // range 0-4
    "WakeOnLink" = 0; // range 0-2

    "*IPChecksumOffloadIPv4" = 3; // range 0-3
    "*LsoV1IPv4" = 1; // range 0-1
    "*LsoV2IPv4" = 1; // range 0-1
    "*LsoV2IPv6" = 1; // range 0-1
    "*TCPChecksumOffloadIPv4" = 3; // range 0-3
    "*TCPChecksumOffloadIPv6" = 3; // range 0-3
    "*UDPChecksumOffloadIPv4" = 3; // range 0-3
    "*UDPChecksumOffloadIPv6" = 3; // range 0-3
```

## MMCSS Values

All values are read via `CiConfigReadDWORD()`, so the type is DWORD for all listed ones. CiConfigInitializeFromRegistry probably handles the `\Tasks\` values.

See [mmcss-CiConfigInitialize.c](https://github.com/nohuto/win-registry/blob/main/assets/mmcss-CiConfigInitialize.c) for notes and [system/desc.md#mmcss-values](https://github.com/nohuto/win-config/blob/main/system/desc.md#mmcss-values) for details on SystemResponsiveness.
```c
"HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\multimedia\\systemprofile";
    "SystemResponsiveness" = 100; // addr 0x1C0011090LL
    "NetworkThrottlingIndex" = 10; // addr 0x1C00110A0LL, 0 = 1, 1..70 keep, 71..0xFFFFFFFE -> 70, 4294967295 (CsInitialize skips CiNdisThrottleWorkItem allocation and CiNdisOpenDevice) keep
    "NoLazyMode" = 0; // addr 0x1C0011080LL, non-zero = true, see below for more
    "IdleDetectionCycles" = 2; // addr 0x1C00110B0LL, valid 1..31 else -> 2
    "LazyModeTimeout" = 1000000; // addr 0x1C00110C0LL, 0 -> 1000000
    "SchedulerTimerResolution" = 10000; // addr 0x1C00110D0LL, >10000 -> 10000
    "SchedulerPeriod" = 100000; // addr 0x1C00110E0LL, valid 50000..1000000 else -> 100000
    "MaxThreadsPerProcess" = 32; // addr 0x1C00110F0LL, valid 8..128 else -> 32
    "MaxThreadsTotal" = 256; // addr 0x1C0011100LL, valid 64..65535 else -> 256
```

## StorNVMe Values

Most values are read via `ReadMultiSzRegistryValueAndCompareId` using the device id match string `VEN_vvvv&DEV_dddd&REV_rr`, so I currently assume that their type is REG_MULTI_SZ. Several values are set to 0 which sometimes also means "ignore" for example. Note that the information in this list is based on `stornvme.sys` only (if value has comment).

[Database-engine/database-file-operations](https://learn.microsoft.com/en-us/troubleshoot/sql/database-engine/database-file-operations/troubleshoot-os-4kb-disk-sector-size?tabs=registry-editor#resolution-steps-for-disk-sector-size-errors-in-sql-server) validates that `ForcedPhysicalSectorSizeInBytes` is a Multi-String value, which confirms a part of my assumption, but I'm still not 100% sure about the rest. Feel free to correct me.

See [stornvme-GetRegistrySettings23H2.c](https://github.com/nohuto/win-registry/blob/main/assets/stornvme-GetRegistrySettings23H2.c) & [stornvme-GetRegistrySettings24H2.c](https://github.com/nohuto/win-registry/blob/main/assets/stornvme-GetRegistrySettings24H2.c) for details.

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\stornvme\\Parameters\\Device";
    "MaxTransferSize" = 0; // REG_MULTI_SZ, range 1..2048, clamp to 2048 (value << 10) 0 = ignore
    "IoQueueDepth" = 0; // REG_MULTI_SZ
    "IoSubmissionQueueCount" = 0; // REG_MULTI_SZ
    "IoCompletionQueueCount" = 0; // REG_MULTI_SZ
    "InterruptCoalescingTime" = 0; // REG_MULTI_SZ
    "InterruptCoalescingEntry" = 0; // REG_MULTI_SZ
    "ArbitrationBurst" = 255; // REG_MULTI_SZ
    "ContiguousMemoryFromAnyNode" = ?; // REG_MULTI_SZ
    "ShutdownTimeout" = 0; // REG_MULTI_SZ, >0xFF coerced to 0xFF, 0 ignored
    "DeallocateMaxLbaCount" = 0; // REG_MULTI_SZ
    "DisableDeallocate" = ?; // REG_MULTI_SZ
    "ControllerBasicInit" = ?; // REG_MULTI_SZ
    "AsyncEventMask" = 1823; // REG_MULTI_SZ, override masked to 0x1F, 0 ignored
    "IdlePowerMode" = 0; // REG_MULTI_SZ, applied only if <=6, skipped when StorPortExtendedFunction(97) sets mode=2
    "DiagnosticFlags" = 0; // REG_MULTI_SZ, bit 1 (0x2) forces LogSize default to 0x100000 bytes
    "LogSize" = 0; // REG_MULTI_SZ, stored as bytes (value << 10) 0 ignored (unless DiagnosticFlags set)
    "IoStripeAlignment" = 0; // REG_MULTI_SZ, applied only if value << 10
    "MedPowerFxIdleTimeout" = 4294967295; // REG_MULTI_SZ
    "LowestPowerFxIdleTimeout" = 50; // REG_MULTI_SZ
    "MedPowerD3IdleTimeout" = 3000; // REG_MULTI_SZ
    "LowestPowerD3IdleTimeout" = 1000; // REG_MULTI_SZ
    "MedPowerResumeLatency" = 4294967295; // REG_MULTI_SZ
    "LowestPowerResumeLatency" = 4294967295; // REG_MULTI_SZ
    "HostMemoryBufferBytes" = 4294967295; // REG_MULTI_SZ
    "BypassSgl" = ?; // REG_MULTI_SZ
    "TestMdlDataBufferOffsetInBytes" = 0; // REG_MULTI_SZ
    "UseDumpPointers" = ?; // REG_MULTI_SZ, presence enables?
    "ReservedQueuePairCount" = 0; // REG_MULTI_SZ, valid 1..65535 (check v69-1 <= 0xFFFE)
    "NvmeTestSwitch" = 1; // REG_MULTI_SZ
    "IoQueuePercentageInPollingMode" = 0; // REG_MULTI_SZ, >100 coerced to 100
    "IoPollingInterval" = 0; // REG_MULTI_SZ, >100000 coerced to 100000
    "IoCompletionCapInDPC" = 100; // REG_MULTI_SZ, if nonzero clamp to 128
    "IoPollingSize" = 0x4000; // REG_MULTI_SZ
    "ErrorEtwThrottleInterval" = 0xD693A400; // REG_MULTI_SZ, if nonzero clamp to max 0xD693A400
    "ResetEnableMask" = ?; // REG_MULTI_SZ
    "ReliabilityDegraded" = ?; // REG_MULTI_SZ
    "ReadOnly" = ?; // REG_MULTI_SZ
    "VolatileMemoryBackupDeviceFailed" = ?; // REG_MULTI_SZ
    "AvailableSpare" = 0; // REG_MULTI_SZ
    "AvailableSpareThreshold" = 0; // REG_MULTI_SZ
    "HostIdentifier" = ?; // REG_BINARY, length 1..16
    "ForcedPhysicalSectorSizeInBytes" = 0; // REG_MULTI_SZ
    "RetainAsyncEventControlMask" = 0; // REG_MULTI_SZ
    "ShutdownTimeoutForSurpriseRemove" = 0; // REG_MULTI_SZ, >0xFF coerced to 0xFF, 0 ignored
    "MaxIoCountLimit" = 0; // REG_MULTI_SZ, if nonzero clamp to 128
    "SubmissionQueueAssignmentPolicy" = 0; // REG_MULTI_SZ
    "DisableMFNDCCDuringRemoval" = ?; // REG_MULTI_SZ
    "EnableSingleDpcForIoCompletion" = ?; // REG_MULTI_SZ
    "DisableNamespacePreferredValueCheck" = ?; // REG_MULTI_SZ
    "IgnoreNamespacePreferredValues" = ?; // REG_MULTI_SZ0
    "DisableBypassIO" = ?; // REG_MULTI_SZ
    "DisableGetActiveNSIDList" = ?; // REG_MULTI_SZ
    "ForceCryptoEraseToUseFormatNVM" = ?; // REG_MULTI_SZ

"HKLM\\SYSTEM\\CurrentControlSet\\Services\\stornvme\\Parameters\\Device";
    "ControllerResetWaitTimeCushion" = 20000; // REG_MULTI_SZ, dynamic read in GetDynamicRegistrySettings, applied even if 0?
    "DisableActivateFWWithoutReset" = 0; // REG_MULTI_SZ, read in GetRegistrySettingsForSpecificKey and returned directly

// these don't exist in the driver on 23H2, but do exist in for example 24H2
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\stornvme\\Parameters\\Device";
    "DisableDSTThrottle" = ?; // REG_MULTI_SZ
    "DisableF0TimestampSync" = ?; // REG_MULTI_SZ
    "DisableForwardedIO" = ?; // REG_MULTI_SZ
    "EnableIntelTSESplitIOWorkaround" = ?; // REG_MULTI_SZ
    "EnforceActiveNamespaceIdentification" = ?; // REG_MULTI_SZ
    "SupportZeroActiveNamespace" = ?; // REG_MULTI_SZ
    "WeightedRoundRobinEnabled" = ?; // REG_MULTI_SZ

// didn't search for them yet
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\stornvme\\Parameters";
    "BusType" = ?;
    "BusyPauseTimeInMs" = ?;
    "BusyRetryCount" = ?;
    "IoLatencyCap" = ?;
    "IoTimeoutValue" = ?;
    "PnpAsyncNewDevices" = ?;

// these seem to exist but not on my current 23H2 nor in the 10.0.26100.7309 version
"HKLM\\SYSTEM\\CurrentControlSet\\Services\\stornvme\\Parameters\\Device";
    "DriverParameter" = ?;
    "LinkTimeout" = ?;
    "MaximumLogicalUnit" = ?;
    "MaximumUCXAddress" = ?;
    "MinimumUCXAddress" = ?;
    "NumberOfRequests" = ?;
    "UncachedExtAlignment" = ?;
```

## Miscellaneous Values

Several values from different keys that I've gathered over time for specific documentation purposes.

> [!WARNING]
> Everything listed below is based on personal research. Mistakes may exist, but I don't think I've made any.

Documented for the [peripheral/disable-touch--tablet](https://github.com/nohuto/win-config/blob/main/peripheral/desc.md#disable-touch--tablet) option, see [peripheral/assets | touch-twinui.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/touch-twinui.c)/[peripheral/assets | touch-InitializeInputSettingsGlobals.c](https://github.com/nohuto/win-config/blob/main/peripheral/assets/touch-InitializeInputSettingsGlobals.c) for details.

```c
"HKCU\\Software\\Microsoft\\Wisp\\Touch";
    "PanningDisabled" = 0;
    "Inertia" = 1;
    "Bouncing" = 1;
    "Friction" = 50;
    "TouchModeN_DtapDist" = 50;
    "TouchModeN_DtapTime" = 50;
    "TouchGate" = 1;
    "TouchModeN_HoldTime_Animation" = 50;
    "TouchModeN_HoldTime_BeforeAnimation" = 50;
    "TouchMode_hold" = 1;
    "Mobile_Inertia_Enabled" = 0;
    "Minimum_Velocity" = 0;
    "Thumb_Flick_Enabled" = 1;

"HKCU\\Software\\Microsoft\\Wisp\\MultiTouch";
    "MultiTouchEnabled" = 1;

"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\PrecisionTouchPad";
    "AAPThreshold" = 2; // range 0–4, touchpad sensitivity
    "CursorSpeed" = 10; // range 1–20, pointer speed
    "FeedbackIntensity" = 50; // range 0–100 (%), haptic feedback strength
    "ClickForceSensitivity" = 50; // range 0–100 (%), relative click-force sensitivity
    "LeaveOnWithMouse" = 1; // 0 = disable touchpad when mouse present, 1 = leave enabled
    "FeedbackEnabled" = 1; // 0 = no haptics, 1 = haptics on
    "TapsEnabled" = 1; // 0/1, single-finger tap-to-click
    "TapAndDrag" = 1; // 0/1, double-tap-and-drag
    "TwoFingerTapEnabled" = 1; // 0/1
    "RightClickZoneEnabled" = 1; // 0/1
    "PanEnabled" = 1; // 0/1, two-finger scrolling
    "ScrollDirection" = 0; // 0 = natural, 1 = reversed
    "ZoomEnabled" = 1;
    "HonorMouseAccelSetting" = 0; // 0 = always apply acceleration, 1 = honor SPI mouse accel?
    "RightClickZoneWidth" = 0;
    "RightClickZoneHeight" = 0;

"HKCU\\Software\\Microsoft\\Wisp\\Pen\\SysEventParameters";
    "Splash" = 50;
    "DblDist" = 50;
    "DblTime" = 300;
    "TapTime" = 100;
    "WaitTime" = 300;
    "HoldTime" = 2300;
    "FlickMode" = 1;
    "FlickTolerance" = 50;
    "Latency" = 8;
    "SampleTime" = 8;
    "UseHWTimeStamp" = 1;
    "SguiMode" = 0;
    "HoldMode" = 1;
    "MouseInputResolutionX" = 0;
    "MouseInputResolutionY" = 0;
    "MouseInputFrequency" = 0;
    "EraseEnable" = 1;
    "RightMaskEnable" = 1;
    "Color" = 0xC0000000C0000000; // ?

"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\TabletMode";
    "STCDefaultMigrationCompleted" = 0; // SHRegValueExists

"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\ImmersiveShell";
    "TabletMode" = 0; // 0 = desktop mode, 1 = tablet mode?
    "ExitedTabletModeWhileCSMActive" = 0; // set to 1 when a3 == 4, HasConvertibleSlateModeChanged() is true
    "TabletModeActivated" = 0; // set to 1 when SetModeInternal() switches into tablet mode

"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ImmersiveShell";
    "AllowPPITabletModeExit" = 0; // SHRegGetBOOLWithREGSAM, non-zero allows the mode switch

"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ImmersiveShell\\OverrideScaling";
    "SmallScreen" = 83; // ?
    "VerySmallScreen" = 71; // ?
    "TabletSmallScreen" = 83; // ?

"HKCU\\Software\\Microsoft\\Wisp\\Pen\\SysEventParameters\\FlickCommands";
    "Left" = { 0x4846455758C33841, 0x9F7145B888BB26B8 };
    "UpLeft" = { 0x47F38E42CEFA51BC, 0xEBDFECA56A8CB1AC };
    "Up"= { 0x450285124653D974, 0x8090833CF6D41AA0 };
    "UpRight" = { 0x47F38E42CEFA51BC, 0x6A8CB1ACEBDFECA5 };
    "Right" = { 0xC267B8DE4FA8068E, 0x4E301EF93B324FAB };
    "DownRight" = { 0x47F38E42CEFA51BC, 0x6A8CB1ACEBDFECA5 };
    "Down" = { 0x441A7051435776E6, 0xF7C82D37F0853D9B };
    "DownLeft" = { 0x47F38E42CEFA51BC, 0xEBDFECA56A8CB1AC };
```

Documented for [system/disable-notifications](https://github.com/nohuto/win-config/blob/main/system/desc.md#disable-notifications). All `NOC_GLOBAL_SETTING_*` values that I found in `NotificationController.dll`.
```c
"HKLM\\SOFTWARE\\Microsoft\\WINDOWS\\CurrentVersion\\Notifications\\Settings"
  'NOC_GLOBAL_SETTING_SUPRESS_TOASTS_WHILE_DUPLICATING'; // Hide notifications when I'm duplicating my screen
  'NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK'; // Show notifications on the lock screen
  'NOC_GLOBAL_SETTING_ALLOW_CRITICAL_TOASTS_ABOVE_LOCK'; // Show reminders and incoming VoIP calls on the lock screen
  'NOC_GLOBAL_SETTING_CORTANA_MANAGED_NOTIFICATIONS';
  'NOC_GLOBAL_SETTING_ALLOW_ACTION_CENTER_ABOVE_LOCK';
  'NOC_GLOBAL_SETTING_HIDE_NOTIFICATION_CONTENT';
  'NOC_GLOBAL_SETTING_TOASTS_ENABLED';
  'NOC_GLOBAL_SETTING_BADGE_ENABLED'; // Don't show number of notifications
  'NOC_GLOBAL_SETTING_GLEAM_ENABLED'; // App icons (Action Center)
  'NOC_GLOBAL_SETTING_ALLOW_HMD_NOTIFICATIONS'; // Show notifications on my head mounted display
  'NOC_GLOBAL_SETTING_ALLOW_CONTROL_CENTER_ABOVE_LOCK';
  'NOC_GLOBAL_SETTING_ALLOW_NOTIFICATION_SOUND'; // Allow notification to play sounds
```