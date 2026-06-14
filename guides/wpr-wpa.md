# Boot Registry Activity

Short guide on how to create a boot registry activity trace and how to format it so regkit can use it. Requirement is the [Windows Performance Toolkit](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/windows-performance-recorder) which you can get from [ADK](https://go.microsoft.com/fwlink/?linkid=2337875), or install it via `winget install Microsoft.WindowsADK`, but this will install more than the Performance Toolkit.

## Starting a Boot Record

Open `WPRUI`, expand `Resource Analysis`, select `Registry I/O activity`. [Performance scenario](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/performance-scenarios) = `Boot`.

```
General: Records general performance while the computer is running.
Boot: Records performance while the computer is booting.
Fast Startup: Records performance during a fast startup.
Shutdown: Records performance while shutting the computer down.
RebootCycle: Records performance during the entire cycle while the computer is rebooting.
Standby/Resume: Records performance when the computer is placed on standby and then resumed.
Hibernate/Resume: Records performance when the computer is placed in hibernation and then resumed.
```

[Detail level](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/detail-level) = `Verbose`. The Light detail level is primarily used for timing recordings. The Verbose detail level provides the detailed information that you need for analysis. [Logging mode](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/logging-mode) = `File`.

```
File mode: Records logging data to a sequential file, it's typically used for short recordings where you can anticipate the events, and file size is mainly limited by available disk space (very large files may be hard to analyze in WPA).

Memory mode: Records logging data to circular buffers in memory, it's typically used for events that can occur at any time, and logging duration is limited by buffer size and profile detail because older events are overwritten (this mode is the default and is recommended to limit file size).
```

Number of iterations = `1`.

![](https://github.com/nohuto/regkit/blob/main/guide/images/WPRUI.png?raw=true)

## Analyzing the Event Trace Log (ETL)

Open the saved `.etl` in WPA, expand `Other`, add the `Registry` graph to the analysis view.

Filter the operations to `QueryValue` by either right clicking on the operation then selecting `Filter To Selection`, or open the filter window and add `Operation` - `equals` - `QueryValue`. If you want a specific key, add a `Entire Key (Base+Remainder) Tree` - `contains` (if recursive) / `equals` (if strict) - `<key>` (note that the key is different to the standart hive keys, e.g. `\Registry\Machine` instead of `HKLM`, or `ControlSet00X` instead of it's link `CurrentControlSet`, see [regkit](https://github.com/nohuto/regkit) documentation).

Move the `Entire Key (Base+Remainder)` column to the far left so it doesn't export the same queried values but from different processes. Press `CTRL+A` to select the entire data table, right click on any row in the `Entire Key (Base+Remainder)` column, `Copy Other` - `Copy Column Selection`.

![](https://github.com/nohuto/regkit/blob/main/guide/images/WPA.png?raw=true)

Create a new `.txt` file anywhere, paste the content into it (preferably use notepad++ here for performance reasons), `Edit` - `Line Operations` - `Sort Lines Lexicographically Ascending` & `Remove Empty Lines`. You can now use the `.txt` via RegKit.