# WPR / [Procmon](https://github.com/5Noxi/wpr-reg-records/blob/main/promon/tracing.md) Registry Activity Records

Records were made while using `24H2` / `IoT Enterprise LTSC 2024`- Subkeys are always included. Most activities were recorded during boot, there are some others, such as `Steam.txt`, `TLOU2.txt`, `StartAllBack.txt`, and `Lighshot.txt`, that were traced using Procmon during use. WPR is included in WADK:
```ps
winget install Microsoft.WindowsADK
```
- [Windows Performance Recorder](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/windows-performance-recorder)  
- [Process Monitor](https://learn.microsoft.com/en-us/sysinternals/downloads/procmon) ([*](https://live.sysinternals.com/))

Guide on how to trace registry activity for a specific app - [procmon.md](https://github.com/5Noxi/wpr-reg-records/blob/main/promon/tracing.md).

## Records Table

| File | Path(s) |
|------|---------|
| [ACPI.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/ACPI.txt) | `HKLM\SYSTEM\ControlSet001\Services\ACPI`<br>`HKLM\SYSTEM\ControlSet001\Services\acpiex`<br>`HKLM\SYSTEM\ControlSet001\Services\AcpiDev`<br>`HKLM\SYSTEM\ControlSet001\Services\acpipagr`<br>`HKLM\SYSTEM\ControlSet001\Services\AcpiPmi`<br>`HKLM\SYSTEM\ControlSet001\Services\acpitime` |
| [AFD-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/AFD-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\AFD\Parameters` |
| [Accessibility.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Accessibility.txt) | `HKCU\Control Panel\Accessibility` |
| [Audio.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Audio.txt) | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Audio` |
| [BFE.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/BFE.txt) | `HKLM\SYSTEM\ControlSet001\Services\BFE` |
| [BrokerInfrastructure.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/BrokerInfrastructure.txt) | `HKLM\SYSTEM\ControlSet001\Services\BrokerInfrastructure` |
| [CV-Explorer.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/CV-Explorer.txt) | `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer` |
| [Classpnp.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Classpnp.txt) | `HKLM\SYSTEM\ControlSet001\Control\Classpnp` |
| [Control-Wdf.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Control-Wdf.txt) | `HKLM\SYSTEM\ControlSet001\Control\Wdf` |
| [ControlPanel-Desktop.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/ControlPanel-Desktop.txt) | `HKCU\Control Panel\Desktop` |
| [ControlPanel-Mouse.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/ControlPanel-Mouse.txt) | `HKCU\Control Panel\Mouse` |
| [CrashControl.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/CrashControl.txt) | `HKLM\SYSTEM\ControlSet001\Control\CrashControl` |
| [Cryptography.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Cryptography.txt) | `HKLM\SOFTWARE\Microsoft\Cryptography` |
| [Disk-Storport-(990Pro).txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Disk-Storport-(990Pro).txt) | `HKLM\SYSTEM\ControlSet001\Enum\SCSI\Disk&Ven_NVMe&Prod_Samsung_SSD_990\5&33c33320&0&000000\Device Parameters\StorPort` (your path will be different) |
| [Dnscache-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Dnscache-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\Dnscache\Parameters` |
| [Enum-USB.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Enum-USB.txt) | `HKLM\SYSTEM\ControlSet001\Enum\USB` |
| [Error-Reporting.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Error-Reporting.txt) | `HKLM\SOFTWARE\Microsoft\WINDOWS\Windows Error Reporting` |
| [FileSystem.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/FileSystem.txt) | `HKLM\SYSTEM\ControlSet001\Control\FileSystem` |
| [GRE-INITIALIZE.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/GRE-INITIALIZE.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\GRE_INITIALIZE` |
| [Graphics-Drivers.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Graphics-Drivers.txt) | `HKLM\SYSTEM\ControlSet001\Control\GraphicsDrivers` |
| [Input.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Input.txt) | `HKLM\SYSTEM\INPUT` |
| [Intel-00XX.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Intel-00XX.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\00XX` (Intel) |
| [Intel.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Intel.txt) | `HKLM\Software\Intel` |
| [Internet-Settings.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Internet-Settings.txt) | `\Software\Microsoft\Windows\CurrentVersion\Internet Settings` |
| [LanmanServer.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/LanmanServer.txt) | `HKLM\SYSTEM\ControlSet001\Services\LanmanServer` |
| [Lighshot.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Lighshot.txt) | `HKCU\Software\SkillBrains\Lightshot` |
| [Lsa.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Lsa.txt) | `HKLM\SYSTEM\ControlSet001\Control\Lsa` |
| [MultiMedia.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/MultiMedia.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\MultiMedia` |
| [NDIS-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/NDIS-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\NDIS\Parameters` |
| [NetBT.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/NetBT.txt) | `HKLM\SYSTEM\ControlSet001\Services\NetBT` |
| [NIC-Intel.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/NIC-Intel.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\00XX` (Intel) |
| [NVIDIA-DispGUID.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/NVIDIA-DispGUID.txt) | `HKLM\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\00XX` |
| [NVIDIA-Corp.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/NVIDIA-Corp.txt) | `HKLM\SOFTWARE\NVIDIA Corporation` |
| [NlaSvc.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/NlaSvc.txt) | `HKLM\SYSTEM\ControlSet001\Services\NlaSvc` |
| [OLE.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/OLE.txt) | `HKLM\SOFTWARE\Microsoft\OLE` ([*](https://learn.microsoft.com/en-us/windows/win32/com/hkey-local-machine-software-microsoft-ole)) |
| [PnP.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/PnP.txt) | `HKLM\SYSTEM\ControlSet001\Control\PnP` |
| [Policies-System.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Policies-System.txt) | `HKLM\SOFTWARE\Policies\Microsoft\WINDOWS\SYSTEM` |
| [Policies.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Policies.txt) | `HKLM\SYSTEM\ControlSet001\Policies` |
| [Policies.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/CV-Policies.txt) | `HKLM\Software\Microsoft\Windows\CurrentVersion\Policies` |
| [Power.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Power.txt) | `HKLM\SYSTEM\ControlSet001\Control\Power` |
| [Session-Manager.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Session-Manager.txt) | `HKLM\SYSTEM\ControlSet001\Control\Session Manager`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Memory Management`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Power`<br>`HKLM\SYSTEM\ControlSet001\Control\Session Manager\Quota System` |
| [StartAllBack.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/StartAllBack.txt) | `HKCU\Software\StartIsBack` |
| [Steam.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Steam.txt) | `HKCU\Software\Valve\Steam` |
| [StorPort.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/StorPort.txt) | `HKLM\SYSTEM\ControlSet001\Control\StorPort` |
| [TLOU2.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/TLOU2.txt) | `HKCU\Software\Naughty Dog` |
| [Tcpip-Parameters.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Tcpip-Parameters.txt) | `HKLM\SYSTEM\ControlSet001\Services\Tcpip\Parameters` |
| [Terminal-Server.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Terminal-Server.txt) | `HKLM\SYSTEM\ControlSet001\Control\Terminal Server` |
| [USB-Flags.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/USB-Flags.txt) | `HKLM\SYSTEM\ControlSet001\Control\usbflags` |
| [USBHUB3.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/USBHUB3.txt) | `HKLM\SYSTEM\ControlSet001\Services\USBHUB3` |
| [WHEA.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/WHEA.txt) | `HKLM\SYSTEM\ControlSet001\Control\WHEA` |
| [Windows-Defender.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Windows-Defender.txt) | `HKLM\SOFTWARE\Policies\Microsoft\Windows Defender` |
| [Windows-Dwm.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Windows-Dwm.txt) | `HKLM\SOFTWARE\Microsoft\Windows\Dwm` |
| [Winows-NT.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Winows-NT.txt) | `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows` |
| [Wisp.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Wisp.txt) | `HKCU\Software\Microsoft\Wisp` |
| [disk.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/disk.txt) | `HKLM\SYSTEM\ControlSet001\Services\disk` |
| [kbdclass.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/kbdclass.txt) | `HKLM\SYSTEM\ControlSet001\Services\kbdclass` |
| [kbdhid.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/kbdhid.txt) | `HKLM\SYSTEM\ControlSet001\Services\kbdhid` |
| [monitor.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/monitor.txt) | `HKLM\SYSTEM\ControlSet001\Services\monitor` |
| [mouclass.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/mouclass.txt) | `HKLM\SYSTEM\ControlSet001\Services\mouclass` |
| [mouhid.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/mouhid.txt) | `HKLM\SYSTEM\ControlSet001\Services\mouhid` |
| [nvlddmkm.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/nvlddmkm.txt) | `HKLM\SYSTEM\ControlSet001\Services\nvlddmkm` |
| [pci.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/pci.txt) | `HKLM\SYSTEM\ControlSet001\Enum\pci` |
| [stornvme.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/stornvme.txt) | `HKLM\SYSTEM\ControlSet001\Services\stornvme\Parameters` |
| [usbhub.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/usbhub.txt) | `HKLM\SYSTEM\ControlSet001\Services\usbhub` |
| [wbem.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/wbem.txt) | `HKLM\SOFTWARE\Microsoft\wbem` |

# Kernel / DXG Kernel Values

Since many people don't yet know which values exist and what default value they have, here's a list. I used IDA, WinDbg, WinObjEx, Windows Internals E7 P1 to create it.

- [Windows Internels E7](https://github.com/5Noxi/windows-books/releases)  
- [WinObjEx64](https://github.com/hfiref0x/WinObjEx64)  
- [WinDbg](https://learn.microsoft.com/en-us/windows-hardware/drivers/debugger/)  
- [Symbols Memory Dump](https://github.com/5Noxi/sym-mem-dump)  

---

## DXG Kernel Values

These are default values I found in `dxgkrnl.sys`, see [dxgkrnl.c](https://github.com/5Noxi/wpr-reg-records/blob/main/dxgkrnl.c) for pseudocode snippets I used / [records/Graphics-Drivers.txt](https://github.com/5Noxi/wpr-reg-records/blob/main/records/Graphics-Drivers.txt) for all values that get read on boot.

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
    "DisableBadDriverCheckForHwProtection"; v70 = 0
    "DisableBoostedVSyncVirtualization"; v59 = 0
    "DisableGdiContextGpuVa"; v41 = 0
    "DisableIndependentVidPnVSync"; v56 = 0
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
    "ForceExplicitResidencyNotification"; v44 = 0
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
    "EnableIntegratedPanelAcmByDefault"; v6 = 0;
    "EnableIntegratedPanelHdrByDefault"; v4 = 0;
    "HDREnabled"; v2 = 0;
    "MicrosoftApprovedAcmSupport"; v5 = 0;

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
```

---

## Session Manager Values

See [session-manager-symbols](https://github.com/5Noxi/wpr-reg-records/blob/main/session-manager-symbols.txt) for reference.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Kernel";
    "AdjustDpcThreshold"; = 20; // KiAdjustDpcThreshold
    "AlwaysTrackIoBoosting"; = 0; // PspAlwaysTrackIoBoosting
    "AmdTprLowerInterruptDelayConfig"; = 0; // KiAmdTprLowerInterruptDelayConfig
    "BoostingPeriodMultiplier"; = 3; // KiNormalPriorityBoostingPeriodMultiplier
    "BugCheckUnexpectedInterrupts"; = 0; // KiBugCheckUnexpectedInterrupts
    "CacheAwareScheduling"; = 47; // KiCacheAwareScheduling
    "CacheErrataOverride"; = 0; // KiTLBCOverride
    "CacheIsoBitmap"; = 0; // KiCacheIsoBitmap
    "DebuggerIsStallOwner"; = 0; // KiDebuggerIsStallOwner
    "DebugPollInterval"; = 2000; // KiDebugPollInterval
    "DefaultDynamicHeteroCpuPolicy"; = 3; // (policy enum only)
    // Behavior of Dynamic hetero policy All (0) (all available) Large (1) LargeOrIdle (2) Small (3) SmallOrIdle (4) Dynamic (5) (use priority and other metrics to decide) BiasedSmall (6) (use priority and other metrics, but prefer small) BiasedLarge (7).
    "DefaultHeteroCpuPolicy"; = 5; // KiDefaultHeteroCpuPolicy
    "DeviceOwnerProtectionDowngradeAllowed"; = 0; // SeDeviceOwnerProtectionDowngradeAllowed
    "DisableControlFlowGuardExportSuppression"; = 0; // PspDisableControlFlowGuardExportSuppression
    "DisableExceptionChainValidation"; = 2; // PspSehValidationPolicy
    "DisableLightWeightSuspend"; = 0; // KiDisableLightWeightSuspend
    "DisableLowQosTimerResolution"; = 1; // KeDisableLowQosTimerResolution
    "DisablePointerParameterAlignmentValidation"; = 0; // KiDisablePointerParameterAlignmentValidation
    "DisableTsx"; = 0; // KiDisableTsx
    "DpcCumulativeSoftTimeout"; = 120000; // KeDpcCumulativeSoftTimeoutMs
    "DpcQueueDepth"; = 4; // KiMaximumDpcQueueDepth
    "DpcSoftTimeout"; = 20000; // KeDpcSoftTimeoutMs
    "DPCTimeout"; = 20000; // KeDpcTimeoutMs
    "DpcWatchdogPeriod"; = 120000; // KeDpcWatchdogPeriodMs
    "DpcWatchdogProfileBufferSizeBytes"; = 266240; // KeDpcWatchdogProfileBufferSizeBytes
    "DpcWatchdogProfileCumulativeDpcThreshold"; = 110000; // KeDpcWatchdogProfileCumulativeDpcThresholdMs
    "DpcWatchdogProfileOffset"; = 10000; // KeDpcWatchdogProfileOffsetMs
    "DpcWatchdogProfileSingleDpcThreshold"; = 18333; // KeDpcWatchdogProfileSingleDpcThresholdMs
    "DriveRemappingMitigation"; = 1; // ObpDriveRemappingMitigation
    "DynamicHeteroCpuPolicyExpectedRuntime"; = 5200; // KiDynamicHeteroCpuPolicyExpectedRuntime
    "DynamicHeteroCpuPolicyImportant"; = 2; // (LargeOrIdle)
    // Policy for a dynamic thread that is deemed important.
    "DynamicHeteroCpuPolicyImportantPriority"; = 8; // KiDynamicHeteroCpuPolicyImportantPriority
    // Priority above which threads are considered important if prioritybased dynamic policy is chosen.
    "DynamicHeteroCpuPolicyImportantShort"; = 3; // (Small)
    // Policy for dynamic thread that is deemed important but run a short amount of time.
    "DynamicHeteroCpuPolicyMask"; = 7; //  (foreground status = 1, priority = 2, expected run time = 4)
    // Determine what is considered in assessing whether a thread is important.
    "EnablePerCpuClockTickScheduling"; = 0; // KiEnableClockTimerPerCpuTickScheduling
    "EnableTickAccumulationFromAccountingPeriods"; = 0; // KiEnableTickAccumulationFromAccountingPeriods
    "EnableWerUserReporting"; = 1; // DbgkEnableWerUserReporting
    "ForceBugcheckForDpcWatchdog"; = 0; // KiForceBugcheckForDpcWatchdog
    "ForceForegroundBoostDecay"; = 0; // KiSchedulerForegroundBoostDecayPolicy
    "ForceIdleGracePeriod"; = 5; // KiForceIdleGracePeriodInSec
    "ForceParkingRequested"; = 1; // KiForceParkingConfiguration
    "GlobalTimerResolutionRequests"; = 0; // KiGlobalTimerResolutionRequests
    "HeteroFavoredCoreFallback"; = 0; // PpmHeteroFavoredCoreFallback
    "HeteroSchedulerOptions"; = 0; // KiHeteroSchedulerOptions
    "HeteroSchedulerOptionsMask"; = 0; // KiHeteroSchedulerOptionsMask
    "HgsPlusFeedbackUpdateThresholdNetRuntime"; = 20; // dword_140FC33C0
    "HgsPlusFeedbackUpdateThresholdRuntime"; = 20; // dword_140FC33B4
    "HgsPlusHigherPerfClassFeedbackThreshold"; = 1; // dword_140FC33E0
    "HgsPlusInvalidFeedbackDefaultClass"; = 0; // dword_140FC33D4
    "HgsPlusInvalidFeedbackDefaultClassSet"; = 0; // dword_140FC33D8
    "HgsPlusInvalidFeedbackLimit"; = 50; // dword_140FC33D0
    "HgsPlusLowerPerfClassFeedbackThreshold"; = 4; // dword_140FC33DC
    "HgsPlusMinimumScoreDifferenceForSwap"; = 25; // dword_140FC33E8
    "HgsPlusThreadCreationDefaultClass"; = 0; // dword_140FC33E4
    "HotpatchTestMode"; = 0; // KeHotpatchTestMode
    "HyperStartDisabled"; = 0; // HvlVpStartDisabled
    "IdealDpcRate"; = 20; // KiIdealDpcRate
    "IdealNodeRandomized"; = 1; // PspIdealNodeRandomized
    "InterruptSteeringFlags"; = 0; // KiInterruptSteeringFlags
    "LongDpcQueueThreshold"; = 3; // KiLongDpcQueueThreshold
    "LongDpcRuntimeThreshold"; = 100; // KiLongDpcRuntimeThreshold
    "MaxDynamicTickDuration"; = 8; // KiMaxDynamicTickDurationSize
    "MaximumCooperativeIdleSearchWidth"; = 16; // KiMaximumCooperativeIdleSearchWidth
    "MaximumSharedReadyQueueSize"; = 260; // KiMaximumSharedReadyQueueSize
    "MinimumDpcRate"; = 3; // KiMinimumDpcRate
    "MitigationAuditOptions"; = 0; // PspSystemMitigationAuditOptions
    "MitigationOptions"; = 0; // PspSystemMitigationOptions
    "ObCaseInsensitive"; = 1; // ObpCaseInsensitive
    "ObObjectSecurityInheritance"; = 0; // ObpObjectSecurityInheritance
    "ObTracePermanent"; = 0; // ObpTracePermanent
    "ObTracePoolTags"; = 0; // ObpTracePoolTagsBuffer / ObpTracePoolTagsLength
    "ObTraceProcessName"; = 0; // ObpTraceProcessNameBuffer / ObpTraceProcessNameLength
    "ObUnsecureGlobalNames"; = 6619246; // ObpUnsecureGlobalNamesBuffer / ObpUnsecureGlobalNamesLength
    "PassiveWatchdogTimeout"; = 300; // KiPassiveWatchdogTimeout
    "PerfIsoEnabled"; = 0; // KiPerfIsoEnabled
    "PoCleanShutdownFlags"; = 0; // PopShutdownCleanly
    "PowerOffFrozenProcessors"; = 1; // KiPowerOffFrozenProcessors
    "ReadyTimeTicks"; = 6; // KiNormalPriorityBoostReadyTimeTicks
    "RebalanceMinPriority"; = 1; // KiRebalanceMinPriority
    "ReservedCpuSets"; = 0; // KiReservedCpuSets
    "ScanLatencyTicks"; = 7; // KiNormalPriorityBoostScanLatencyTicks
    "SchedulerAssistThreadFlagOverride"; = 0; // KiSchedulerAssistThreadFlagOverride
    "SeAllowAllApplicationAceRemoval"; = 0; // SepAllowAllApplicationAceRemoval
    "SeAllowSessionImpersonationCapability"; = 0; // SepAllowSessionImpersonationCap
    "SeCompatFlags"; = 0; // SeCompatFlags
    "SeLpacEnableWatsonReporting"; = 0; // SeLpacEnableWatsonReporting
    "SeLpacEnableWatsonThrottling"; = 1; // SeLpacEnableWatsonThrottling
    "SerializeTimerExpiration"; = 1; // KiSerializeTimerExpiration
    // This behavior is controlled by the kernel variable KiSerializeTimerExpiration, which is initialized based on a registry setting whose value is different between a server and client installation. By modifying or creating the value SerializeTimerExpiration under HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel other than 0 or 1, serialization can be disabled, enabling timers to be distributed among processors. Deleting the value, or keeping it as 0, allows the kernel to make the decision based on Modern Standby availability, and setting it to 1 permanently enables serialization even on non-Modern Standby systems.
    "SeTokenDoesNotTrackSessionObject"; = 0; // SeTokenDoesNotTrackSessionObject
    "SeTokenLeakDiag"; = 0; // SeTokenLeakTracking
    "SeTokenSingletonAttributesConfig"; = 3; // SepTokenSingletonAttributesConfig
    "SplitLargeCaches"; = 0; // KiSplitLargeCaches
    "ThreadDpcEnable"; = 1; // KeThreadDpcEnable
    "ThreadReadyCount"; = 1; // KiNormalPriorityBoostMaximumThreadReadyCount
    "TimerCheckFlags"; = 1; // KeTimerCheckFlags
    "VerifierDpcScalingFactor"; = 1; // KeVerifierDpcScalingFactor
    "VirtualHeteroHysteresis"; = 4294967295; // PpmPerfQosTransitionHysteresisOverride
    "VpThreadSystemWorkPriority"; = 30; // KiVpThreadSystemWorkPriority
    "WpsSimulationOverride"; = 0; // PpmWpsSimulationOverride / PpmWpsSimulationOverrideSize
    "XStateContextLookasidePerProcMaxDepth"; = 0; // KiXStateContextLookasidePerProcMaxDepth

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Kernel\\RNG";
    "RNGAuxiliarySeed"; = ; // ExpRNGAuxiliarySeed = 742978275?

"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager";
    "ProtectionMode"; = 1; // ObpProtectionMode 
    "ObjectSecurityMode"; = 1; // ObpObjectSecurityMode 
    "GlobalFlag"; = 2336816981; // CmNtGlobalFlag 
    "GlobalFlag2"; = 1399515354; // CmNtGlobalFlag2 
    "SkipRegistryInit"; = 0; // CmNtSkipRegistryInit 
    "CWDIllegalInDLLSearch"; = 0; // PspCurDirDevicesSkippedForDlls 
    "ResourceTimeoutCount"; = 45; // ExResourceTimeoutCount 
    "ResourceCheckFlags"; = 3; // ExResourceCheckFlags 
    "ResourceEnforceOwnerTransfer"; = 0; // ExpResourceEnforceOwnerTransfer 
    "CriticalSectionTimeout"; = ?; // dword_140FC3204
    "HeapSegmentReserve"; = ?; // qword_140FC3228
    "HeapSegmentCommit"; = ?; // qword_140FC3220
    "HeapDeCommitTotalFreeThreshold"; = ?; // qword_140FC3218
    "HeapDeCommitFreeBlockThreshold"; = ?; // qword_140FC3210
    "PowerPolicySimulate"; = 0; // PopSimulate 
    "Debugger Retries"; = 20; // KdpContext 
    "AlpcMessageLog"; = 0; // AlpcpMessageLogEnabled 
    "AlpcWakePolicy"; = 1; // AlpcpWakePolicyDefault 
    "ImageExecutionOptions"; = 0; // ViImageExecutionOptions 
    "InitConsoleFlags"; = 0; // InitConsoleFlags 
    "MultiUsersInSessionSupported"; = 0; // RtlpMultiUsersInSessionSupported 
    "DisableIFEOCaching"; = 0; // RtlpDisableIFEOCaching 

"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Quota System";
    "JobTimeLimitsPeriodSeconds"; = 7; // PspJobTimeLimitsPeriodSeconds 
    "ApplicationBlockedMessageLimit"; = 50; // PspJobNoWakeChargeLimit 
    "SystemBlockedMessageLimit"; = 200; // PspSystemNoWakeChargeLimit 

"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management";
    "PagedPoolQuota"; = ?; // unk_140FD7DE4
    "NonPagedPoolQuota"; = 4294967295; // PspDefaultResourceLimits (0xFFFFFFFF) 
    "PagingFileQuota"; = ?; // unk_140FD7DE8
    "WorkingSetPagesQuota"; = ?; // unk_140FD7DEC
    "AllocationPreference"; = ?; // dword_140FC3200
    "Mirroring"; = ?; // dword_140FC31F4
    "MoveImages"; = 1; // MmRegistryState 
    "LowMemoryThreshold"; = ?; // qword_140FC3230
    "HighMemoryThreshold"; = ?; // qword_140FC3238
    "DisablePageCombining"; = ?; // dword_140FC31E8
    "DisablePagingExecutive"; = ?; // dword_140FC31E4
    "EnableCooling"; = ?; // dword_140FC31F8
    "ModifiedWriteMaximum"; = ?; // dword_140FC31FC
    "PoolTagSmallTableSize"; = 4097; // PoolTrackTableSize 
    "PoolForceFullDecommit"; = 0; // PoolForceFullDecommit 
    "PoolTag"; = 0; // MmSpecialPoolTag 
    "PoolTagOverruns"; = 1; // MmSpecialPoolCatchOverruns 
    "ProtectNonPagedPool"; = 0; // MmProtectFreedNonPagedPool 
    "TrackLockedPages"; = 0; // MmTrackLockedPages 
    "TrackPtes"; = ?; // dword_140FC31EC
    "PageValidationAction"; = 0; // MmPageValidationAction 
    "PageValidationFrequency"; = 0; // MmPageValidationFrequency 
    "ForceValidateIo"; = ?; // dword_140FC31F0
    "PhysicalMemoryMapperEnforcementMode"; = ?; // dword_140FC324C
    "VerifierOptions"; = 0; // VfOptionFlags 
    "VerifierTipDisable"; = 0; // VerifierTipDisable 
    "VerifierNewRuleWorkaround"; = 0; // VerifierNewRuleWorkaround 
    "VerifierIrpStackTraces"; = 16384; // IovIrpTracesLength 
    "VerifierHandleTraces"; = 16384; // VfHandleTracingEntries 
    "VerifierIrpTimeout"; = 0; // VfWdIrpTimeoutMsec 
    "VerifyDriverLevel"; = 4294967295; // MmVerifyDriverLevel (0xFFFFFFFF) 
    "VerifyDrivers"; = 594044221; // MmVerifyDriverBuffer (0x23706B3D) 
    "VerifyDriversLength"; = 1447310913; // MmVerifyDriverBufferLength (0x56415741) 
    "VerifierSettingState"; = 0; // VfRuleClasses 
    "VerifierSettingStateSize"; = 4294967295; // VfRuleClassesSize (0xFFFFFFFF) 
    "XdvVerifierOptions"; = 0; // VfFlightOptions 
    "VerifyMode"; = 4; // VfVerifyMode 
    "VerifyTriage"; = 4294967295; // ViVerifyTriage (0xFFFFFFFF) 
    "VerifyTriageRules"; = 0; // ViVerifyTriageRules 
    "VerifyTriageRulesSize"; = 4294967295; // ViVerifyTriageRulesSize (0xFFFFFFFF) 
    "DifPluginConfigData"; = 3506164991; // DifPluginConfigData (0xD0E9FFFF) 
    "DifPluginConfigDataLength"; = 1413828185; // DifPluginConfigDataLength (0x54415541) 
    "VerifierFaultProbability"; = 600; // VfFaultInjectionProbability (0x00000258) 
    "VerifierFaultBootMinutes"; = 8; // VfFaultInjectionBootMinutes 
    "VerifierFaultApplications"; = 0; // VerifierFaultApplicationsBuffer 
    "VerifierFaultApplicationsSize"; = 4294967295; // VerifierFaultApplicationsBufferSize (0xFFFFFFFF) 
    "VerifierFaultTags"; = 0; // VerifierFaultTagsBuffer 
    "VerifierFaultTagsSize"; = 4294967295; // VerifierFaultTagsBufferSize (0xFFFFFFFF) 
    "VerifierDifPoolTags"; = 0; // DifpPoolTags 
    "VerifierDifPoolTagsSizeBytes"; = 4294967295; // DifpPoolTagsSizeBytes (0xFFFFFFFF) 
    "VerifierRandomTargets"; = 0; // VfRandomVerifiedDrivers 
    "VerifierTipLimitDenominator"; = 0; // DifiPluginControlDenominator 
    "VerifierTipLimitNumerator"; = 0; // DifiPluginControlNumerator 
    "VerifierTipSparseness"; = 0; // DifiPluginControlSparseness 
    "VerifierTriageContext"; = 0; // VfTriageContext 
    "XdvTipTag"; = 0; // CarTipTag 
    "XdvVerifierOptions"; = 0; // CarXdvOptions 
    "VerifyBTSBufferSize"; = 0; // ViVerifyBTSBufferSize 
    "VerifyDriversSuppress"; = 418795416; // VfXdvSuppressDriversBuffer (0x18E8ABE8) 
    "VerifyDriversSuppressLength"; = 1212432054; // VfXdvSuppressDriversBufferLength (0x48535756) 
    "DeadlockRecursionDepthLimit"; = 0; // ViRecursionDepthLimitFromRegistry 
    "DeadlockSearchNodesLimit"; = 0; // ViSearchedNodesLimitFromRegistry 
    "MinimumStackCommitInBytes"; = ?; // dword_140FC3208
    "WorkingSetSwapSharedPages"; = 0; // PspOutSwapSharedPages 
    "VmPauseOutswapSizeCapMB"; = 512; // VmPauseOutswapSizeCapMB 
    "RemoteFileDirtyPageThreshold"; = 1310720; // CcRemoteFileDPInlineFlushThreshold (0x00140000) 
    "MaxLazyWritePages"; = 0; // CcMaxLazyWritePagesOverride 
    "EnablePerVolumeLazyWriter"; = 2; // CcEnablePerVolumeLazyWriterOverride 
    "EnableAsyncLazywrite"; = 2; // CcEnableAsyncLazywriteOverride 
    "EnableAsyncLazywriteMulti"; = 2; // CcEnableAsyncLazywriteMultiOverride 
    "CacheUnmapBehindLengthInMB"; = 8388608; // CcUnmapBehindLength (0x00800000) 
    "DisableCacheTelemetry"; = 2; // CcDisableTelemetryRegKeyAtInit 
    "TopBottomDPTEqual"; = 0; // CcAzure_TopBottomDPTEqual 
    "LazyWriterPercentageOfNumProcs"; = 0; // CcAzure_LazyWriterPercentageOfNumProcs 
    "LargeWriteSize"; = 0; // CcAzure_LargeWriteSize 
    "SoftThrottleLargeWriteAtPct"; = 0; // CcAzure_SoftThrottleLargeWriteAtPct 
    "SoftThrottleDelayInMs"; = 0; // CcAzure_SoftThrottleDelayInMs 
    "CustomDTPDenominator"; = 8; // CcClientDTPDenominator 
    "SimulateCommitSavings"; = ?; // dword_140FC3240
    "KernelPadSectionsOverride"; = ?; // dword_140FC3248
    "AllowUserHotPatchWithoutVbs"; = ?; // dword_140FC3250
    "SpecialPurposeMemoryStartPage"; = 0; // MmSpecialPurposeMemoryStartPage 
    "SpecialPurposeMemoryStartPageValueSize"; = 4294967295; // MmSpecialPurposeMemoryStartPageValueSize (0xFFFFFFFF) 
    "SpecialPurposeMemoryPages"; = 0; // MmSpecialPurposeMemoryPages 

"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Executive";
    "AdditionalCriticalWorkerThreads"; = 0; // ExpAdditionalCriticalWorkerThreads 
    "AdditionalDelayedWorkerThreads"; = 0; // ExpAdditionalDelayedWorkerThreads 
    "MaximumKernelWorkerThreads"; = 4096; // ExpMaximumKernelWorkerThreads 
    "WorkerThreadTimeoutInSeconds"; = 600; // ExpWorkerThreadTimeoutInSeconds (0x00000258) 
    "KernelWorkerTestFlags"; = 0; // ExpWorkerQueueTestFlags 
    "WorkerFactoryThreadCreationTimeout"; = 10; // ExpWorkerFactoryThreadCreationTimeoutInSeconds 
    "WorkerFactoryThreadIdleTimeout"; = 67; // ExpWorkerFactoryThreadIdleTimeoutInSeconds 
    "ForceEnableMutantAutoboost"; = 0; // ExpForceEnableMutantAutoboost 
    "MaxTimeSeparationBeforeCorrect"; = 60; // ExpMaxTimeSeperationBeforeCorrect 

"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power";
    "SleepStudyDisabled"; = 1; // PopSleepStudyDisabled 
    "SleepStudyDeviceAccountingLevel"; = 0; // PopSleepStudyDeviceAccountingLevel 
    "WatchdogSleepTimeout"; = 300; // PopWatchdogSleepTimeout (0x0000012C) 
    "WatchdogResumeTimeout"; = 120; // PopWatchdogResumeTimeout (0x00000078) 
    "IdleScanInterval"; = 30; // PopIdleScanInterval 
    "FlushPolicy"; = 0; // PopFlushPolicy 
    "Win32CalloutWatchdogBugcheckEnabled"; = 0; // PopWin32CalloutWatchdogBugcheckEnabled 
    "SkipTickOverride"; = 1; // PopSkipTickPolicy 

// Miscellaneous values

"HKLM\SYSTEM\CurrentControlSet\Control\LSA";
    "AuditBaseDirectories"; = 0; // ObpAuditBaseDirectories 
    "AuditBaseObjects"; = 0; // ObpAuditBaseObjects 

"HKLM\SYSTEM\CurrentControlSet\Control\LSA\audit";
    "ProcessAccessesToAudit"; = 0; // SepProcessAccessesToAudit 

"HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation";
    "ActiveTimeBias"; = ?; // dword_140FCE974
    "Bias"; = -60; // ExpAltTimeZoneBias (0xFFFFFFC4) 
    "RealTimeIsUniversal"; = 0; // ExpRealTimeIsUniversal 

"HKLM\SYSTEM\CurrentControlSet\Control\I/O System";
    "LargeIrpStackLocations"; = 15; // IopLargeIrpStackLocations 
    "MediumIrpStackLocations"; = 2; // IopMediumIrpStackLocations 
    "IoBlockLegacyFsFilters"; = 0; // IopBlockLegacyFsFilters 
    "IoCaseInsensitive"; = 1; // IopCaseInsensitive 
    "RequireDeviceAccessCheck"; = 1; // IopRequireDeviceAccessCheck 
    "IoFailZeroAccessCreate"; = 1; // IopFailZeroAccessCreate 
    "IoEnableSessionZeroAccessCheck"; = 0; // IopSessionZeroAccessCheckEnabled 
    "IoAllowLoadCrashDumpDriver"; = 0; // IopAllowLoadCrashDumpDriver 
    "IoIrpCompletionTimeoutInSeconds"; = 300; // IopIrpCompletionTimeoutInSeconds (0x0000012C) 
    "IoKeepAliveTimeMs"; = 5000; // IopKeepAliveTimeMs (0x00001388) 
    "DisableDiskCounters"; = 0; // PsDisableDiskCounters 

"HKLM\SYSTEM\CurrentControlSet\Control\Configuration Manager";
    "FastBoot"; = 1; // CmFastBoot 
    "SelfHealingEnabled"; = 1; // CmSelfHeal 
    "RegistryLazyFlushInterval"; = 60; // CmpLazyFlushIntervalInSeconds 
    "RegistryLazyReconcileInterval"; = 3600; // CmpLazyReconcileIntervalInSeconds (0x00000E10) 
    "RegistryLazyLocalizeInterval"; = 60; // CmpLazyLocalizeIntervalInSeconds 
    "RegistryLazyFlushBootDelay"; = 60; // CmpEnableLazyFlushBootDelayInterval 
    "RegistryLogFileSizeCap"; = 0; // CmpLogFileSizeCap 
    "RegistryReorganizationLimit"; = 1048576; // CmpReorganizeLimit (0x00100000) 
    "RegistryReorganizationLimitDays"; = 7; // CmpReorganizeDelayDays 
    "RegistryFlushGlobalFlags"; = 0; // CmpGlobalFlushControlFlags 
    "VirtualizationEnabled"; = 1; // CmVEEnabled 
    "DelayCloseSize"; = 2048; // CmpDelayedCloseSize (0x00000800) 
    "FreezeThawTimeoutInSeconds"; = 60; // CmFreezeThawTimeoutInSeconds 
    "SystemHiveLimitSize"; = 1610612736; // CmSystemHiveLimitSize (0x60000000) 
    "EnablePeriodicBackup"; = 0; // CmpDoIdleProcessing 
    "VolatileBoot"; = 0; // CmpVolatileBoot 
    "BugcheckRecoveryEnabled"; = 0; // CmBugcheckRecoveryEnabled 
    "Enabled"; = 0; // CmpLKGEnabled 
    "CallbackMemoryFromPool"; = 0; // CmpAllocateCallbackMemoryFromPool 
    "CallbackMemoryFromPerProcLookaside"; = 1; // CmpAllocateCallbackMemoryFromPerProcLookaside 

"HKLM\SYSTEM\CurrentControlSet\Control\StateSeparation\Policy";
    "Enabled"; = 0; // CmStateSeparationEnabled 
    "DevelopmentMode"; = 0; // CmStateSeparationDevMode 
    "AllHivesVolatile"; = 0; // CmStateSeparationAllHivesVolatile 

"HKLM\SYSTEM\CurrentControlSet\Control\ValidationRunlevels";
    "Global"; = 89274380; // CmGlobalValidationRunlevel (0x008824AC) 

```

## Power Values

See [kernel-symbols](https://github.com/5Noxi/wpr-reg-records/blob/main/power-symbols.txt) for reference.

```c
"HKLM\SYSTEM\CurrentControlSet\Control\Power";
    "MaximumFrequencyOverride"; = 0; // PpmFrequencyOverride
    "HibernateEnabledDefault"; = 0; // PopHiberEnabledDefaultReg
    "HiberbootEnabled"; = 0; // PopHiberbootEnabledReg
    "HibernateChecksummingEnabled"; = 1; // PopHiberChecksummingEnabledReg
    "EnableMinimalHiberFile"; = 0; // PopEnableMinimalHiberFile
    "ForceMinimalHiberFile"; = 0; // PopForceMinimalHiberFile
    "ThermalPollingMode"; = 0; // PopThermalPollingMode
    "ThermalTelemetryVerbosity"; = 1; // PopThermalTelemetryVerbosity
    "SmartUserPresenceGracePeriod"; = 1800; // PopSmartUserPresenceGracePeriod
    "SmartUserPresenceWakeOffset"; = 300; // PopSmartUserPresenceWakeOffset
    "SmartUserPresenceCheckTimeout"; = 10800; // PopSmartUserPresenceCheckTimeout
    "SmartUserPresenceAction"; = 0; // PopSmartUserPresenceAction
    "DozeDeferralMaxSeconds"; = 259200; // PopDozeDeferralMaxSeconds
    "DozeDeferralChecksToIgnore"; = 0; // PopDozeDeferralChecksToIgnore
    "Win32kCalloutWatchdogTimeoutSeconds"; = 30; // PopWin32kCalloutWatchdogTimeoutSeconds
    "PdcIdlePhaseDefaultWatchdogTimeoutSeconds"; = 30; // PopPdcIdlePhaseDefaultWatchdogTimeoutSeconds
    "FxAccountingTelemetryDisabled"; = 0; // PopDiagFxAccountingTelemetryDisabled
    "PowerActionTransitioningWatchdogTimeoutDefault"; = 600; // PopPowerActionTransitioningWatchdogTimeoutDefault
    "PowerActionResumeWatchdogTimeoutDefault"; = 300; // PopPowerActionResumingWatchdogTimeoutDefault
    "MSDisabled"; = 0; // PopModernStandbyDisabled
    "PlatformRoleOverride"; = 4294967295; // PopPlatformRoleOverride
    "PlatformAoAcOverride"; = 4294967295; // PopPlatformAoAcOverride
    "PerfBoostAtGuaranteed"; = 0; // PpmPerfBoostAtGuaranteed
    "IpiLastClockOwnerDisable"; = 0; // PpmIpiLastClockOwnerDisable
    "MfBufferingThreshold"; = 0; // PpmMfBufferingThreshold
    "MfOverridesDisabled"; = 1; // PpmMfOverridesDisabled
    "LatencyToleranceDefault"; = 100000; // PpmLatencyToleranceLimit
    "LatencyToleranceVSyncEnabled"; = 13001; // dword_140FC3424 dd 32C9
    "LatencyToleranceFSVP"; = 20000; // dword_140FC3428 dd 4E20
    "LatencyToleranceIdleResiliency"; = 1500000; // dword_140FC342C dd 16E360
    "LatencyToleranceSoftParked"; = 0; // PpmIdleSoftParkedLatencyLimit
    "LatencyToleranceParked"; = 0; // PpmIdleParkedLatencyLimit
    "PerfIdealAggressiveIncreasePolicyThreshold"; = 90; // PpmPerfIdealAggressiveIncreaseThreshold
    "PerfSingleStepSize"; = 5; // PpmPerfSingleStepSize
    "PerfCalculateActualUtilization"; = 1; // PpmPerfCalculateActualUtilization
    "PerfArtificialDomain"; = 4294967295; // PpmPerfArtificialDomainSetting
    "MultiparkGranularity"; = 8; // PpmParkMultiparkGranularity
    "Class1InitialUnparkCount"; = 64; // PpmParkInitialClass1UnParkCount
    "HighPerfDurationBoot"; = 90000; // PpmHighPerfDuration
    "HighPerfDurationCSExit"; = ?; // unk_140FC337C
    "HighPerfDurationSxExit"; = ?; // unk_140FC3380
    "IdleDurationExpirationTimeout"; = 4; // PpmIdleDurationExpirationTimeoutMs
    "DisableIdleStatesAtBoot"; = 0; // PpmIdleDisableStatesAtBoot
    "BootHeteroPolicyOverride"; = 0; // PpmPerfBootHeteroPolicyOverrideEnabled
    "HeteroMultiCoreClassesEnabled"; = 4294967295; // PpmHeteroMultiCoreClassesRegValue
    "HeteroMultiClassParkingEnabled"; = 4294967295; // PpmHeteroMultiClassParkingRegValue
    "AlwaysComputeQosHints"; = 0; // PpmPerfAlwaysComputeQosEnabled
    "ExperimentalClusterIdleMitigation"; = 0; // PpmIdleClusterIdleMitigation
    "EnforceDisconnectedStandby"; = 0; // PopEnforceDisconnectedStandby
    "EnforceAusterityMode"; = 0; // PopEnforceAusterityMode
    "UserBatteryDischargeEstimator"; = 0; // PopDisableBatteryDischargeEstimator
    "UserBatteryChargeEstimator"; = 0; // PopUserBatteryChargingEstimator
    "StandbyConnectivityGracePeriod"; = 0; // PopStandbyConnectivityGracePeriod
    "HeteroFavoredCoreRotationTimeoutMs"; = 30000; // PpmHeteroFavoredCoreRotationTimeoutMs
    "HeteroHgsPlusDisabled"; = 0; // PpmHeteroHgsThreadDisabled
    "PerfCheckTimerImplementation"; = 0; // PpmCheckTimerImplementation
    "HeteroHgsEePerfHintsIndependentEnabled"; = 0; // PpmHeteroHgsEePerfHintsIndependentEnabled
    "HeteroWpsWorkloadProminenceCutoff"; = 35; // PpmHeteroWpsWorkloadProminenceCutoff
    "EnablePowerButtonSuppression"; = 4294967295; // PopEnablePowerButtonSuppressionOverride
    "HeteroWpsContainmentEnumOverride"; = 0; // PpmHeteroWpsContainmentEnumOverride
    "ManualDimTimeout"; = ?; // PopAdaptiveManualDimTimeout
    "SuppressResumePrompt"; = 0; // PopSuppressResumePrompt
    "EventProcessorEnabled"; = 1; // PopEventProcessorEnabled
    "PdcOneWayEntry"; = 0; // PopPowerAggregatorOneWayEntry
    "RestrictedStandbyDozeTimeoutSeconds"; = 0; // PopPowerAggregatorRestrictedStandbyDozeTimeoutSeconds
    "PromoteHibernateToShutdown"; = 0; // PopPromoteHibernateToShutdown
    "CoalescingTimerInterval"; = 0; // PopCoalescingTimerInterval
    "CoalescingFlushInterval"; = 60; // PopCoalescingFlushInterval
    "EnableInputSuppression"; = 4294967295; // PopEnableInputSuppressionOverride
    "IgnoreLidStateForInputSuppression"; = 4294967295; // PopLidStateForInputSuppressionOverride
    "ActiveIdleTimeout"; = 1000; // PopFxActiveIdleTimeout
    "ActiveIdleThreshold"; = 5000000; // PopFxActiveIdleThreshold
    "ActiveIdleLevel"; = 1; // PopFxActiveIdleLevel
    "DirectedFxDefaultTimeout"; = 120; // PopFxDirectedFxDefaultTimeout
    "IdleStateTimeout"; = 500; // PopPepIdleStateTimeout
    "WatchdogWorkOrderTimeout"; = 300000; // PopFxWatchdogWorkOrderTimeout
    "DripsCallbackInterval"; = 35; // PopDripsCallbackInterval
    "DripsWatchdogDebounceInterval"; = 120; // PopDripsWatchdogDebounceInterval
    "DripsWatchdogTimeout"; = 300; // PopDripsWatchdogTimeout
    "DripsWatchdogAction"; = 198; // PopDripsWatchdogAction
    "DirectedDripsOverride"; = 4294967295; // PopDirectedDripsOverride
    "DirectedDripsTimeout"; = 300; // PopDirectedDripsTimeout
    "DirectedDripsDebounceInterval"; = 120; // PopDirectedDripsDebounceInterval
    "DirectedDripsAction"; = 3; // PopDirectedDripsAction
    "DirectedDripsWaitWakeTimeout"; = 5; // PopDirectedDripsWaitWakeTimeoutSeconds
    "DirectedDripsSurprisePowerOnTimeout"; = 5; // PopDirectedDripsSurprisePowerOnTimeoutSeconds
    "DirectedDripsDfxEnforcementPolicy"; = 1; // PopDirectedDripsDfxEnforcementPolicy
    "SleepstudyAccountingEnabled"; = 0; // SleepstudyHelperAccountingEnabled
    "SleepstudyLibraryBlockerLimit"; = 200; // SleepstudyHelperBlockerLibraryLimit
    "SleepstudyGlobalBlockerLimit"; = 3000; // SleepstudyHelperBlockerGlobalLimit
    "DisableVsyncLatencyUpdate"; = 0; // PpmDisableVsyncLatencyUpdate
    "ExitLatencyCheckEnabled"; = 1; // PpmExitLatencyCheckEnabled
    "EnergyEstimationEnabled"; = 1; // PopEnergyEstimationEnabled
    "CheckPowerSourceAfterRtcWakeTime"; = 30; // PopCheckPowerSourceAfterRtcWakeTime
    "DripsSwHwDivergenceThreshold"; = 270; // PopDripsSwHwDivergenceThreshold
    "DripsSwHwDivergenceEnableLiveDump"; = 0; // PopDripsSwHwDivergenceEnableLiveDump
    "IgnoreCsComplianceCheck"; = 0; // PopIgnoreCsComplianceCheck
    "PerfQueryOnDevicePowerChanges"; = 0; // PopFxPerfQueryOnDevicePowerChanges
    "EnforceConsoleLockScreenTimeout"; = 0; // PopEnforceConsoleLockScreenTimeout
    "DeepIoCoalescingEnabled"; = 0; // PopDeepIoCoalescingEnabled
    "AllowSystemRequiredPowerRequests"; = 1; // PopPowerRequestConvertSystemToExecution
    "AllowAudioToEnableExecutionRequiredPowerRequests"; = 1; // PopPowerRequestActiveAudioEnablesExecutionRequired
    "TtmEnabled"; = 0; // TtmpEnabled
    "ProximityEscapeMsec"; = 0; // TtmpProximityEscapeMsec
    "TimerRebaseThresholdOnDripsExit"; = 60; // PopTimerRebaseThresholdRegValue
    "FxRuntimeLogNumberEntries"; = 64; // PopFxRuntimeLogNumberEntries
    "CheckpointSystemSleep"; = 0; // PopCheckpointSystemSleepEnabledReg
    "IdleProcessorsRequireQosManagement"; = 0; // PpmPerfQosManageIdleProcessors
    "CheckpointSystemSleepSimulateFlags"; = 0; // PopCheckpointSystemSleepSimulateFlags
    "SkipHibernateMemoryMapValidation"; = 4294967295; // PopEnableHibernateMemoryMapValidationOverride
    "PoFxSystemIrpWaitForReportDevicePowered"; = 0; // PopPoFxSystemIrpWaitForReportDevicePoweredReg
    "DisableDisplayBurstOnPowerSourceChange"; = 0; // PopDisableDisplayBurstOnPowerSourceChange
    "DisableInboxPepGeneratedConstraints"; = 4294967295; // PopDisableInboxPepGeneratedConstraintsOverride
    "HibernateBootOptimizationEnabled"; = 0; // PopHiberBootOptimizationEnabledReg
    "HiberFileTypeDefault"; = 4294967295; // PopHiberFileTypeDefaultReg

"HKLM\SYSTEM\CurrentControlSet\Control\Power\ForceHibernateDisabled";
    "Policy"; = 0; // PopHiberForceDisabledReg
    "GuardedHost"; = ?; // unk_140FC5234

"HKLM\SYSTEM\CurrentControlSet\Control\Power\HiberFileBucket";
    "Percent1GBFull"; = ?; // unk_140FC3670
    "Percent1GBReduced"; = ?; // unk_140FC366C
    "Percent2GBFull"; = ?; // unk_140FC3688
    "Percent2GBReduced"; = ?; // unk_140FC3684
    "Percent4GBFull"; = ?; // unk_140FC36A0
    "Percent4GBReduced"; = ?; // unk_140FC369C
    "Percent8GBFull"; = ?; // unk_140FC36B8
    "Percent8GBReduced"; = ?; // unk_140FC36B4
    "Percent16GBFull"; = ?; // unk_140FC36D0
    "Percent16GBReduced"; = ?; // unk_140FC36CC
    "Percent32GBFull"; = ?; // unk_140FC36E8
    "Percent32GBReduced"; = ?; // unk_140FC36E4
    "PercentUnlimitedFull"; = ?; // unk_140FC3700
    "PercentUnlimitedReduced"; = ?; // unk_140FC36FC

"HKLM\SYSTEM\CurrentControlSet\Control\Power\ModernSleep";
    "EnabledActions"; = 0; // PopAggressiveStandbyActionsRegValue
    "EnableDsNetRefresh"; = 0; // PopEnableDsNetRefresh

"HKLM\SYSTEM\CurrentControlSet\Control\Power\PowerThrottling";
    "PowerThrottlingOff"; = 1; // PpmPerfQosGroupPolicyDisable
```
