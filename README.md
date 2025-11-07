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

    "ForceDirectFlip"; v66 = 0
    "DisableOverlays"; v67 = 0
    "EnableOfferReclaimOnDriver"; v37 = 1
    "LeanMemoryLimit"; v123 = 16
    "LeanMemoryLimit"; v122 = 1395864371
    "ForceEnableDxgMms2"; v39 = 0
    "ContextNoPatchMode"; v38 = 0
    "Force32BitFences"; v68 = 0
    "InitialPagingQueueFenceValue"; v45 = 7000
    "ForceInitPagingProcessVaSpace"; v40 = 0
    "DisableGdiContextGpuVa"; v41 = 0
    "DisablePagingContextGpuVa"; v42 = 0
    "DisableMonitoredFenceGpuVa"; v43 = 0
    "ForceExplicitResidencyNotification"; v44 = 0
    "DriverManagesResidencyOverride"; v46 = 1
    "GdiPhysicalAdapterIndex"; v74 = 0
    "ForceReplicateGdiContent"; v47 = 0
    "EnableTimedCalls"; v49 = 0
    "CreateGdiPrimaryOnSlaveGpu"; v48 = 0
    "ForceSurpriseRemovalSupport"; v75 = 0
    "EnableDecodeMPO"; v69 = 1
    "DisableBadDriverCheckForHwProtection"; v70 = 0
    "ForceSecondaryMPOSupport"; v97 = 0
    "ForceSecondaryIFlipSupport"; v72 = 0
    "EnablePanelFitterSupport"; v100 = 0
    "EnableMultiPlaneOverlay3DDIs"; v73 = 0
    "DisableSecondaryIFlipSupport"; v71 = 0
    "EnableWDDM23Synchronization"; v50 = 0
    "IoMmuFlags"; v51 = 0
    "DisableMultiSourceMPOCheck"; v76 = 0
    "DriverStoreCopyMode"; v33 = 1
    "ForceVariableRefresh"; v52 = 0
    "DeadlockTimeout"; v53 = 30000
    "DeadlockPulse"; v54 = 5000
    "DeadlockPulseTolerance"; v55 = 500
    "DisableIndependentVidPnVSync"; v56 = 0
    "NumVirtualFunctions"; v65 = 0
    "CrtcPhaseFrames"; v57 = 2
    "EnableFbrValidation"; v58 = 1
    "DisableBoostedVSyncVirtualization"; v59 = 0
    "EnableBasicRenderGpuPv"; v60 = 0
    "KnownProcessBoostMode"; v61 = 1
    "SmallQuantumMode"; v62 = 1
    "HighPriorityCompletionMode"; v63 = 1
    "GpuPriorityChangeMode"; v64 = 1

    "EnableRuntimePowerManagement"; v178 = 1;
    "DisableDevicePowerRequired"; v179 = 0;
    "DefaultLatencyToleranceOther"; v175 = -1;
    "DefaultExpectedResidency"; v176 = 2000;
    "DefaultLatencyToleranceIdle0"; v184 = 80;
    "DefaultLatencyToleranceIdle1"; v185 = 15000;
    "DefaultLatencyToleranceNoContext"; v186 = 35000;
    "DefaultLatencyToleranceIdle0MonitorOff"; v188 = 2000;
    "DefaultLatencyToleranceIdle1MonitorOff"; v189 = 50000;
    "DefaultLatencyToleranceNoContextMonitorOff"; v190 = 100000;
    "DefaultLatencyToleranceTimerPeriod"; v183 = 200;
    "DefaultIdleThresholdIdle0"; v187 = 200;
    "DefaultIdleThresholdIdle0MonitorOff"; v222 = 100;
    "MonitorLatencyTolerance"; v208 = 300000;
    "MonitorRefreshLatencyTolerance"; v207 = 17000;
    "DefaultPowerNotRequiredTimeout"; v209 = 25000;
    "DefaultActiveIdleThreshold"; v191 = 2000;
    "ulow"; v170 = 300;
    "uhigh"; v169 = 700;
    "uglitch"; v168 = 900;
    "uideal"; v167 = 500;
    "lowdebounce"; v182 = 3;
    "EnablePODebounce"; v180 = 0;
    "DisablePStateManagement"; v181 = 0;
    "DefaultD3TransitionLatencyActivelyUsed"; v192 = 80;
    "DefaultD3TransitionLatencyIdleShortTime"; v194 = 80000;
    "DefaultD3TransitionLatencyIdleLongTime"; v196 = 140000;
    "DefaultD3TransitionLatencyIdleVeryLongTime"; v198 = 200000;
    "DefaultD3TransitionLatencyIdleNoContext"; v199 = 250000;
    "DefaultD3TransitionLatencyIdleMonitorOff"; v200 = 250000;
    "DefaultD3TransitionIdleShortTimeThreshold"; v193 = 10000;
    "DefaultD3TransitionIdleLongTimeThreshold"; v195 = 60000;
    "DefaultD3TransitionIdleVeryLongTimeThreshold"; v197 = 60000;
    "DefaultLatencyToleranceMemory"; v201 = 15000;
    "DefaultLatencyToleranceMemoryNoContext"; v202 = 30000;
    "DefaultMemoryRefreshLatencyToleranceActivelyUsed"; v203 = 80;
    "DefaultMemoryRefreshLatencyToleranceIdleShortTime"; v204 = 15000;
    "DefaultMemoryRefreshLatencyToleranceNoContext"; v205 = 30000;
    "DefaultMemoryRefreshLatencyToleranceMonitorOff"; v206 = 80000;

    "TerminationListSizeLimit"; v62 = 67108864;
    "ValidateWDDMCaps"; v63 = 0;
    "WDDM2LockManagement"; v61 = 1;
    "MaximumAdapterCount"; v60 = 32;
    "InvestigationDebugParameter"; v65 = 0;
    "EnableIgnoreWin32ProcessStatus"; v66 = 0;
    "EnableHMDTestMode"; v67 = 0;
    "PreserveFirmwareMode"; v68 = 0;
    "PreventFullscreenWireFormatChange"; v69 = 0;
    "EnableFuzzing"; v64 = 0;
    "InternalDiagnosticsBufferSize"; v55 = 65536;
    "InternalDiagnosticsBufferMultiplier"; v57 = 2;
    "ExternalDiagnosticsBufferSize"; v56 = 16384;
    "ExternalDiagnosticsBufferMultiplier"; v59 = 1;
    "DiagnosticsBufferExpansionTime"; v58 = 300;
    "RapidHpdTimeoutInMilliseconds"; v70 = 0;
    "RapidHpdMaxChainInMilliseconds"; v71 = 0;
    "ForceUsb4MonitorSupport"; g_bDbgForceUsb4MonitorSupport = 0;
    "Usb4MonitorTargetId"; g_DbgUsb4MonitorTargetId = 0;
    "Usb4MonitorDpcdUSB4_Driver_ID"; g_DbgUsb4MonitorDpcdUSB4_Driver_ID = 0;
    "Usb4MonitorDpcdDP_IN_Adapter_Number"; g_DbgUsb4MonitorDpcdDP_IN_Adapter_Number = 0;
    "Usb4MonitorPowerOnDelayInSeconds"; g_DbgUsb4MonitorPowerOnDelayInSeconds = 0;
    "TreatUsb4MonitorAsNormal"; g_bDbgTreatUsb4MonitorAsNormal = 0;
    "AllowAdvancedEtwLogging"; v72 = 0;
    "NodeUsageTelemetryTimerInterval"; v73 = v73; // ?

    "GpuVirtualizationFlags"; v50 = (g_VgpuReplaceWarp ? 0x8 : 0x0); // bit0: CreatePVGpu=0, bit2: ForceSvm=0, bit3: ReplaceWarp=default from g_VgpuReplaceWarp ?
    "DisableVaBackedVm"; g_VgpuDisableVaBackedVm = 0;
    "VirtualGpuOnly"; g_VirtualGpuOnly = 0;
    "LimitNumberOfVfs"; g_LimitNumberOfVfs = 0;
    "DisableVersionMismatchCheck"; v52 = 0;

    "MiracastDefaultRtspPort"; dword_1C0153F64 = 7236;
    "PlatformSupportMiracast"; v26 = 0;
    "SuspendAdapterTimerPeriod"; v27 = 500000;
    "SupportMultipleIntegratedDisplays"; v28 = 0;
    "HwSchMode"; v29 = 0;
    "HwSchOverrideBlockList"; v31 = 1;
    "HwSchTreatExperimentalAsStable"; v30 = 0;
    "ForceBddFallbackOnly"; v35 = 0;

    "RapidHPDTime"; v16 = 1000;
    "RapidHPDThresholdCount"; *(_DWORD*)((char*)this + 544) = 5;
    "EnableExperimentalRefreshRates"; v22 = 0;

    "TdrLevel"; v13 = 3;
    "TdrDelay"; v8 = 2;
    "TdrDodPresentDelay"; v9 = 2;
    "TdrDodVSyncDelay"; v10 = 2;
    "TdrDdiDelay"; v11 = 5;
    "TdrLimitCount"; v14 = 5;
    "TdrLimitTime"; v15 = 60;
    "TdrDebugMode"; v12 = 2;

    "TdrTestMode"; v14 = 0

    "UnsupportedMonitorModesAllowed"; v5 = 0;

    "ForceEnableDWMClone"; v82 = 0

    "MultiMonSupport"; v39 = 1;

    "HybridInternalPanelOverrideEnable"; v13 = 0

    "OutputDuplicationSessionApplicationLimit"; v14 = 4

    "IsInternalRelease"; v44 = 0

    "DRTTestEnable"; v14 = 0; // 1484026436 = Enabled ?

    "EnableAcmSupportDeveloperPreview"; v7 = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Power";
    "UseSelfRefreshVRAMInS3"; v166 = 1;

"<PnPDeviceKey>\\DxgkSettings";
    "UseSelfRefreshVRAMInS3"; v166 = 1;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\BasicDisplay";
    "BasicDisplayUserNotified"; v2 = 0;

    "EnableBasicDisplayFallback"; v32 = -1;
    "DisableBasicDisplayFallback"; v33 = -1;
    "ForcePreserveBootDisplay"; v34 = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Smm";
    "ForceEnableIommu"; v3 = 0;
    "EnablePageTracking"; v8 = 0;
    "LogicalAddressMode"; v4 = 0;
    "PreferHighLogicalAddresses"; v10 = 0;
    "DebugMode"; v11 = 0;
    "IdentityMappedPassthrough"; v7 = 0;
    "ForceDmaRemapping"; v9 = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\DMM";
    "BadMonitorModeDiag"; v17 = 2;
    "AssertOnDdiViolation"; g_DmmAssertOnDdiViolation = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\DMM";
    "ModeListCaching"; v81 = 1;
    "SetTimingsFlags"; *((_DWORD*)this + 130) = 0;
    "ShortLinkTrainingTimeout"; *((_DWORD*)this + 131) = 200;
    "LongLinkTrainingTimeout"; *((_DWORD*)this + 132) = 1000;
    "HPDFilterLimit"; *((_DWORD*)this + 133) = 20000000;
    "EnableVirtualRefreshRateOnExternalMonitor"; *((_DWORD*)this + 134) = 0;

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Validation";
    "Level"; v7 = 0
    "FailEscapeDDI"; v8 = 0
    "FailRenderDDI"; v9 = 0
    "FailReserveGPUVA"; v10 = 0
    "ReportVirtualMachine"; v11 = 0

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\MonitorDataStore\\MONITOR-ID"
    "HDREnabled"; v2 = 0;
    "AdvancedColorEnabled"; v3 = 0;
    "EnableIntegratedPanelHdrByDefault"; v4 = 0;
    "MicrosoftApprovedAcmSupport"; v5 = 0;
    "EnableIntegratedPanelAcmByDefault"; v6 = 0;
    "AutoColorManagementEnabled"; v8 = 0;

"<AdapterPnpKey>";
    "EnableVirtualTopologySupport"; v84 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : EnableVirtualTopologySupport
    "NeedToSuspendVidSchBeforeSetGammaRamp"; v83 = (AdapterBuild < 8704 ? 1 : 0)
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : NeedToSuspendVidSchBeforeSetGammaRamp
    // \Registry\Machine\SYSTEM\ControlSet001\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}\0000 : NeedToSuspendVidSchBeforeSetGammaRamp

    "DisableNonPOSTDevice"; v40 = 0;
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicDisplay : DisableNonPOSTDevice
    // \Registry\Machine\SYSTEM\ControlSet001\Services\BasicRender : DisableNonPOSTDevice

    "Device PnP";
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
```

---

## Kernel Values

See [Kernel-Symbols](https://github.com/5Noxi/wpr-reg-records/blob/main/kernel-values.txt) for reference.

```c
"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Kernel";
    "BoostingPeriodMultiplier"; = 3; // KiNormalPriorityBoostingPeriodMultiplier
    "DefaultDynamicHeteroCpuPolicy"; = 3; // (policy enum only)
    // Behavior of Dynamic hetero policy All (0) (all available) Large (1) LargeOrIdle (2) Small (3) SmallOrIdle (4) Dynamic (5) (use priority and other metrics to decide) BiasedSmall (6) (use priority and other metrics, but prefer small) BiasedLarge (7).
    "DefaultHeteroCpuPolicy"; = 5; // KiDefaultHeteroCpuPolicy
    "DynamicHeteroCpuPolicyExpectedRuntime"; = 5200; // KiDynamicHeteroCpuPolicyExpectedRuntime
    "DynamicHeteroCpuPolicyImportant"; = 2; // (LargeOrIdle)
    // Policy for a dynamic thread that is deemed important.
    "DynamicHeteroCpuPolicyImportantPriority"; = 8; // KiDynamicHeteroCpuPolicyImportantPriority
    // Priority above which threads are considered important if prioritybased dynamic policy is chosen.
    "DynamicHeteroCpuPolicyImportantShort"; = 3; // (Small)
    // Policy for dynamic thread that is deemed important but run a short amount of time.
    "DynamicHeteroCpuPolicyMask"; = 7; //  (foreground status = 1, priority = 2, expected run time = 4)
    // Determine what is considered in assessing whether a thread is important.
    "ReadyTimeTicks"; = 6; // KiNormalPriorityBoostReadyTimeTicks
    "ReservedCpuSets"; = 0; // KiReservedCpuSets
    "ScanLatencyTicks"; = 7; // KiNormalPriorityBoostScanLatencyTicks
    "SeTokenSingletonAttributesConfig"; = 3; // SepTokenSingletonAttributesConfig
    "ThreadReadyCount"; = 1; // KiNormalPriorityBoostMaximumThreadReadyCount
    "DriveRemappingMitigation"; = 1; // ObpDriveRemappingMitigation
    "DPCTimeout"; = 20000; // KeDpcTimeoutMs
    "DpcSoftTimeout"; = 20000; // KeDpcSoftTimeoutMs
    "DpcCumulativeSoftTimeout"; = 120000; // KeDpcCumulativeSoftTimeoutMs
    "DpcWatchdogPeriod"; = 120000; // KeDpcWatchdogPeriodMs
    "DpcWatchdogProfileOffset"; = 10000; // KeDpcWatchdogProfileOffsetMs
    "DpcWatchdogProfileBufferSizeBytes"; = 266240; // KeDpcWatchdogProfileBufferSizeBytes
    "DpcWatchdogProfileSingleDpcThreshold"; = 18333; // KeDpcWatchdogProfileSingleDpcThresholdMs
    "DpcWatchdogProfileCumulativeDpcThreshold"; = 110000; // KeDpcWatchdogProfileCumulativeDpcThresholdMs
    "AmdTprLowerInterruptDelayConfig"; = 0; // KiAmdTprLowerInterruptDelayConfig
    "VerifierDpcScalingFactor"; = 1; // KeVerifierDpcScalingFactor
    "ThreadDpcEnable"; = 1; // KeThreadDpcEnable
    "DpcQueueDepth"; = 4; // KiMaximumDpcQueueDepth
    "MinimumDpcRate"; = 3; // KiMinimumDpcRate
    "MaximumSharedReadyQueueSize"; = 260; // KiMaximumSharedReadyQueueSize
    "MaximumCooperativeIdleSearchWidth"; = 16; // KiMaximumCooperativeIdleSearchWidth
    "AdjustDpcThreshold"; = 20; // KiAdjustDpcThreshold
    "PassiveWatchdogTimeout"; = 300; // KiPassiveWatchdogTimeout
    "SchedulerAssistThreadFlagOverride"; = 0; // KiSchedulerAssistThreadFlagOverride
    "SerializeTimerExpiration"; = 1; // KiSerializeTimerExpiration
    // This behavior is controlled by the kernel variable KiSerializeTimerExpiration, which is initialized based on a registry setting whose value is different between a server and client installation. By modifying or creating the value SerializeTimerExpiration under HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\kernel other than 0 or 1, serialization can be disabled, enabling timers to be distributed among processors. Deleting the value, or keeping it as 0, allows the kernel to make the decision based on Modern Standby availability, and setting it to 1 permanently enables serialization even on non-Modern Standby systems.
    "GlobalTimerResolutionRequests"; = 0; // KiGlobalTimerResolutionRequests
    "DisableLowQosTimerResolution"; = 1; // KeDisableLowQosTimerResolution
    "EnablePerCpuClockTickScheduling"; = 0; // KiEnableClockTimerPerCpuTickScheduling
    "EnableTickAccumulationFromAccountingPeriods"; = 0; // KiEnableTickAccumulationFromAccountingPeriods
    "DisableTsx"; = 0; // KiDisableTsx
    "VpThreadSystemWorkPriority"; = 30; // KiVpThreadSystemWorkPriority
    "PerfIsoEnabled"; = 0; // KiPerfIsoEnabled
    "CacheIsoBitmap"; = 0; // KiCacheIsoBitmap
    "IdealDpcRate"; = 20; // KiIdealDpcRate
    "CacheErrataOverride"; = 0; // KiTLBCOverride
    "MaxDynamicTickDuration"; = 8; // KiMaxDynamicTickDurationSize
    "DisableExceptionChainValidation"; = 2; // PspSehValidationPolicy
    "MitigationOptions"; = 0; // PspSystemMitigationOptions
    "MitigationAuditOptions"; = 0; // PspSystemMitigationAuditOptions
    "DisableControlFlowGuardExportSuppression"; = 0; // PspDisableControlFlowGuardExportSuppression
    "ForceIdleGracePeriod"; = 5; // KiForceIdleGracePeriodInSec
    "HotpatchTestMode"; = 0; // KeHotpatchTestMode
    "ObTracePoolTags"; = 0; // ObpTracePoolTagsBuffer / ObpTracePoolTagsLength
    "ObTraceProcessName"; = 0; // ObpTraceProcessNameBuffer / ObpTraceProcessNameLength
    "ObTracePermanent"; = 0; // ObpTracePermanent
    "PoCleanShutdownFlags"; = 0; // PopShutdownCleanly
    "ObUnsecureGlobalNames"; = 6619246; // ObpUnsecureGlobalNamesBuffer / ObpUnsecureGlobalNamesLength
    "TimerCheckFlags"; = 1; // KeTimerCheckFlags
    "SplitLargeCaches"; = 0; // KiSplitLargeCaches
    "AlwaysTrackIoBoosting"; = 0; // PspAlwaysTrackIoBoosting
    "ObCaseInsensitive"; = 1; // ObpCaseInsensitive
    "BugCheckUnexpectedInterrupts"; = 0; // KiBugCheckUnexpectedInterrupts
    "DebuggerIsStallOwner"; = 0; // KiDebuggerIsStallOwner
    "DebugPollInterval"; = 2000; // KiDebugPollInterval
    "PowerOffFrozenProcessors"; = 1; // KiPowerOffFrozenProcessors
    "DisableLightWeightSuspend"; = 0; // KiDisableLightWeightSuspend
    "ObObjectSecurityInheritance"; = 0; // ObpObjectSecurityInheritance
    "SeAllowAllApplicationAceRemoval"; = 0; // SepAllowAllApplicationAceRemoval
    "HyperStartDisabled"; = 0; // HvlVpStartDisabled
    "SeAllowSessionImpersonationCapability"; = 0; // SepAllowSessionImpersonationCap
    "InterruptSteeringFlags"; = 0; // KiInterruptSteeringFlags
    "SeCompatFlags"; = 0; // SeCompatFlags
    "SeTokenLeakDiag"; = 0; // SeTokenLeakTracking
    "SeTokenDoesNotTrackSessionObject"; = 0; // SeTokenDoesNotTrackSessionObject
    "HeteroFavoredCoreFallback"; = 0; // PpmHeteroFavoredCoreFallback
    "VirtualHeteroHysteresis"; = 4294967295; // PpmPerfQosTransitionHysteresisOverride
    "WpsSimulationOverride"; = 0; // PpmWpsSimulationOverride / PpmWpsSimulationOverrideSize
    "RebalanceMinPriority"; = 1; // KiRebalanceMinPriority
    "SeLpacEnableWatsonReporting"; = 0; // SeLpacEnableWatsonReporting
    "SeLpacEnableWatsonThrottling"; = 1; // SeLpacEnableWatsonThrottling
    "EnableWerUserReporting"; = 1; // DbgkEnableWerUserReporting
    "DeviceOwnerProtectionDowngradeAllowed"; = 0; // SeDeviceOwnerProtectionDowngradeAllowed
    "CacheAwareScheduling"; = 47; // KiCacheAwareScheduling
    "HeteroSchedulerOptions"; = 0; // KiHeteroSchedulerOptions
    "HeteroSchedulerOptionsMask"; = 0; // KiHeteroSchedulerOptionsMask
    "ForceForegroundBoostDecay"; = 0; // KiSchedulerForegroundBoostDecayPolicy
    "ForceBugcheckForDpcWatchdog"; = 0; // KiForceBugcheckForDpcWatchdog
    "LongDpcRuntimeThreshold"; = 100; // KiLongDpcRuntimeThreshold
    "LongDpcQueueThreshold"; = 3; // KiLongDpcQueueThreshold
    "HgsPlusFeedbackUpdateThresholdRuntime"; = 20; // dword_140FC33B4
    "HgsPlusFeedbackUpdateThresholdNetRuntime"; = 20; // dword_140FC33C0
    "HgsPlusInvalidFeedbackLimit"; = 50; // dword_140FC33D0
    "HgsPlusInvalidFeedbackDefaultClass"; = 0; // dword_140FC33D4
    "HgsPlusInvalidFeedbackDefaultClassSet"; = 0; // dword_140FC33D8
    "HgsPlusLowerPerfClassFeedbackThreshold"; = 4; // dword_140FC33DC
    "HgsPlusHigherPerfClassFeedbackThreshold"; = 1; // dword_140FC33E0
    "HgsPlusMinimumScoreDifferenceForSwap"; = 25; // dword_140FC33E8
    "HgsPlusThreadCreationDefaultClass"; = 0; // dword_140FC33E4
    "DisablePointerParameterAlignmentValidation"; = 0; // KiDisablePointerParameterAlignmentValidation
    "XStateContextLookasidePerProcMaxDepth"; = 0; // KiXStateContextLookasidePerProcMaxDepth
    "IdealNodeRandomized"; = 1; // PspIdealNodeRandomized
    "ForceParkingRequested"; = ?; // (no symbol)

"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Kernel\\RNG";
    "RNGAuxiliarySeed"; = ; // ExpRNGAuxiliarySeed = 742978275?
```
