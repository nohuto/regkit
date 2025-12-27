# Tracing Program Registry Activity

Requirements are [process monitor](https://learn.microsoft.com/en-us/sysinternals/downloads/procmon) and the program you want to trace. I'll trace the activity of the game `The Last Of Us Part II`.

Start with configuring the filter (`CTRL` + `L`):  
`Process Name` `contains` `tlou`  
`Operation` `is` `RegSetValue` - Registry writes  
`Operation` `is` `RegQueryValue` - Registry reads  
`Operation` `is` `RegDeleteValue`  

You can disable the other activity monitoring:

![](https://github.com/nohuto/win-registry/blob/main/images/monitoring.png)

If you want to trace all `Reg...` activity, which would use:
```c
RegCloseKey
RegCreateKey
RegDeleteKey
RegDeleteValue
RegEnumKey
RegEnumValue
RegFlushKey
RegLoadKey
RegOpenKey
RegQueryKey
RegQueryKeySecurity
RegQueryMultipleValueKey
RegQueryValue
RegSetInfoKey
RegSetKeySecurity
RegSetValue
RegUnloadKey
```
`Operation` `begins with` `Reg`

Now you'll see all kind of registry activity, what you want to look for is the path the program uses, if there's any. TLOU II uses:
`Path` `contains` `\Software\Naughty Dog\The Last of Us Part II`

Save the content using `.csv`, afterwards you can remove specific parts of the file, if you want to.

![](https://github.com/nohuto/win-registry/blob/main/images/pmsave.png)

`CTRL` + `H` (Regular expression):  
`"Reg\w*"` - Removes the operations  
`"process.exe"` - Replace `process`  

The result should look similar to: [TLOU2.txt](https://github.com/nohuto/win-registry/blob/main/TLOU2.txt)/[StartAllBack.txt](https://github.com/nohuto/win-registry/blob/main/StartAllBack.txt)/[Lighshot](https://github.com/nohuto/win-registry/blob/main/Lighshot.txt). This can be used to trace specific program/path registry activity, there's obviously much more you can do with procmon.

Operations which can be used:
```
CloseFile
CreateFile
CreateFileMapping
DeviceIoControl
FileSystemControl
FlushBuffersFile
Load Image
LockFile
NotifyChangeDirectory
Process Create
Process Exit
Process Profiling
Process Start
QueryAllInformationFile
QueryAttributeInformationVolume
QueryAttributeTagFile
QueryBasicInformationFile
QueryDeviceInformationVolume
QueryDeviceRelations
QueryDirectory
QueryEAFile
QueryEaInformationFile
QueryFileInternalInformationFile
QueryFullSizeInformationVolume
QueryInformationVolume
QueryNameInformationFile
QueryNetworkOpenInformationFile
QueryNetworkPhysicalNameInformationFile
QueryNormalizedNameInformationFile
QueryObjectIdInformationVolume
QueryOpen
QueryPositionInformationFile
QueryRemoteProtocolInformation
QuerySecurityFile
QuerySizeInformationVolume
QueryStandardInformationFile
QueryStreamInformationFile
ReadFile
RegCloseKey
RegCreateKey
RegDeleteKey
RegDeleteValue
RegEnumKey
RegEnumValue
RegFlushKey
RegLoadKey
RegOpenKey
RegQueryKey
RegQueryKeySecurity
RegQueryMultipleValueKey
RegQueryValue
RegSetInfoKey
RegSetKeySecurity
RegSetValue
RegUnloadKey
SetAllocationInformationFile
SetBasicInformationFile
SetDispositionInformationEx
SetDispositionInformationFile
SetEndOfFileInformationFile
SetPositionInformationFile
SetRenameInformationFile
SetSecurityFile
SetStorageReservedIdInformation
TCP Accept
TCP Connect
TCP Disconnect
TCP Receive
TCP Reconnect
TCP Retransmit
TCP Send
TCP TCPCopy
Thread Create
Thread Exit
UDP Receive
UDP Send
UnlockFileSingle
WriteFile
```
