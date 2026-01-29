# RegKit

RegKit is a native Windows Registry editor written in C++ using the Win32 API and common controls (comctl32) for performance reasons. It exposes both registry views, the standard hives and the REGISTRY root.

## Table of Content

- [Differences to Default RegEdit](https://github.com/nohuto/regkit#differences-to-default-regedit)
- [Standard hives & REGISTRY Comparison](https://github.com/nohuto/regkit#standard-hives--registry-comparison)
- [Registry fundamentals](https://github.com/nohuto/regkit#registry-fundamentals)
  - [Keys, values, and naming](https://github.com/nohuto/regkit#keys-values-and-naming)
  - [Registry value types](https://github.com/nohuto/regkit#registry-value-types)
  - [Root keys and logical structure](https://github.com/nohuto/regkit#root-keys-and-logical-structure)
  - [Hives and on-disk files](https://github.com/nohuto/regkit#hives-and-on-disk-files)
- [REGISTRY only Keys](https://github.com/nohuto/regkit#registry-only-keys)
- [Icons Meaning](https://github.com/nohuto/regkit#icons-meaning)
  - [Symlink Icon](https://github.com/nohuto/regkit#symlink-icon)
  - [Database Icon](https://github.com/nohuto/regkit#database-icon)
- [Trace Menu](https://github.com/nohuto/regkit#trace-menu)
- [Credits/References](https://github.com/nohuto/regkit#creditsreferences)


## Differences to Default RegEdit

RegKit adds functionality that standard RegEdit doesn't support/expose:

- A real REGISTRY root view in addition to standard hives
- Symbolic link detection (`SymbolicLinkValue` value with the link target)
- Hive backed key detection using hivelist key
- Open Hive File (opens the backing hive file)
- Trace presets (23H2/24H2/25H2 - see below), used for "Read on boot" column
- Extra hives toggle, exposes additional predefined keys that RegEdit typically doesn't show, such as `HKEY_PERFORMANCE_DATA` (live performance counter data produced on demand, not stored in a hive file) and related keys like `HKEY_PERFORMANCE_TEXT`/`HKEY_PERFORMANCE_NLSTEXT` for e.g. counter name strings (read more [here](https://learn.microsoft.com/en-us/windows/win32/perfctrs/using-the-registry-functions-to-consume-counter-data))
- Theme modes (System/Light/Dark)
- Custom font support
- Favorites import/export
- Comment column for values with import/export support
- Loading/unloading hives
- Local/remote/offline registry
- Undo/redo, copy/paste (entire keys), replace, performant 'Find'
- Tab control
- History view
- Option to save/forget previous key tree state
- Research menu (redirections to [win-registry](https://github.com/nohuto/win-registry))
- Miscellaneous common functionalities

## Standard hives & REGISTRY Comparison

RegEdit shows five common hives: `HKEY_LOCAL_MACHINE`, `HKEY_USERS`, `HKEY_CURRENT_USER`, `HKEY_CLASSES_ROOT`, and `HKEY_CURRENT_CONFIG`. Internally, all registry keys are rooted at a single object named `\REGISTRY` in the Object Manager namespace. Native APIs (NtOpenKey / ZwOpenKey) can access paths under `\REGISTRY` directly. The registry actually exposes nine root keys (including performance and local-settings roots) but most tools only surface the common five.

## Registry fundamentals

### Keys, values, and naming

The registry is a database that looks a lot like a filesystem, keys are like directories, values are like files, and a key can contain both subkeys and values. Values are typed, have a name, and live under a key. Each key also has one unnamed value, displayed as `(Default)`.

### Registry value types

Most values are `REG_DWORD`, `REG_BINARY`, or `REG_SZ`, but the registry supports 12 value types.

| Type | Description |
| --- | --- |
| `REG_NONE` | No value type |
| `REG_SZ` | Fixed-length Unicode string |
| `REG_MULTI_SZ` | Array of Unicode NULL-terminated strings |
| `REG_EXPAND_SZ` | Variable-length Unicode string with embedded environment variables |
| `REG_BINARY` | Arbitrary-length binary data |
| `REG_DWORD` | 32-bit number |
| `REG_QWORD` | 64-bit number |
| `REG_DWORD_BIG_ENDIAN` | 32-bit number, high byte first |
| `REG_LINK` | Unicode symbolic link |
| `REG_RESOURCE_LIST` | Hardware resource description |
| `REG_FULL_RESOURCE_DESCRIPTOR` | Hardware resource description |
| `REG_RESOURCE_REQUIREMENTS_LIST` | Resource requirements |

### Root keys and logical structure

There are nine root keys, their names start with `HKEY` as they represent handles (H) to keys (KEY), some are links or merged views.

| Root key | Abbreviation | Description | Link |
| --- | --- | --- | --- |
| `HKEY_CURRENT_USER` | `HKCU` | Per-user preferences (current logged-on user) | `HKEY_USERS\<SID>` (SID of current logged-on user) |
| `HKEY_CURRENT_USER_LOCAL_SETTINGS` | `HKCULS` | Per-user settings local to the machine | `HKCU\Software\Classes\Local Settings` |
| `HKEY_USERS` | `HKU` | All loaded user profiles (including `.DEFAULT` for the system account) | - |
| `HKEY_CLASSES_ROOT` | `HKCR` | "Stores file association and Component Object Model (COM) object registration information" | Merged view of `HKLM\SOFTWARE\Classes` and `HKEY_USERS\<SID>\SOFTWARE\Classes` |
| `HKEY_LOCAL_MACHINE` | `HKLM` | Machine-wide configuration (BCD, COMPONENTS, HARDWARE, SAM, SECURITY, SOFTWARE, SYSTEM) | - |
| `HKEY_CURRENT_CONFIG` | `HKCC` | Stores some information about the current hardware profile (deprecated, "Hardware profiles are no longer supported in Windows, but the key still exists to support legacy applications that might depend on its presence." | `HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current` (legacy, Yosifovich shows `Hardware\Profiles\Current`, but that's a typo in his blog) |
| `HKEY_PERFORMANCE_DATA` | `HKPD` | Live performance counter data, available only via APIs | - |
| `HKEY_PERFORMANCE_TEXT` | `HKPT` | Performance counter names/descriptions in US English | - |
| `HKEY_PERFORMANCE_NLSTEXT` | `HKPNT` | Performance counter names/descriptions in the OS language | - |

Notes:
- `HKEY_CURRENT_USER` maps to the logged-on user hive (`Ntuser.dat`) and is created per-user at logon.
- `HKEY_CLASSES_ROOT` also contains UAC VirtualStore data, it is not a simple link.
- `HKEY_PERFORMANCE_*` keys aren't stored in hive files and are not visible in Regedit. They are provided by Perflib through registry APIs like `RegQueryValueEx`.
- SYSTEM = `S-1-5-18`, LocalService = `S-1-5-19`, NetworkService = `S-1-5-20`

### Hives and on-disk files

On disk, the registry is a set of hive files, not a single monolithic file. The Configuration Manager records loaded hive paths under `HKLM\SYSTEM\CurrentControlSet\Control\Hivelist` as they are mounted. The mapping below is from Windows Internals (some hives are volatile or virtualized):

| Hive registry path | Hive file path |
| --- | --- |
| `HKEY_LOCAL_MACHINE\BCD00000000` | `\EFI\Microsoft\Boot\BCD` |
| `HKEY_LOCAL_MACHINE\COMPONENTS` | `%SystemRoot%\System32\Config\Components` |
| `HKEY_LOCAL_MACHINE\SYSTEM` | `%SystemRoot%\System32\Config\System` |
| `HKEY_LOCAL_MACHINE\SAM` | `%SystemRoot%\System32\Config\Sam` |
| `HKEY_LOCAL_MACHINE\SECURITY` | `%SystemRoot%\System32\Config\Security` |
| `HKEY_LOCAL_MACHINE\SOFTWARE` | `%SystemRoot%\System32\Config\Software` |
| `HKEY_LOCAL_MACHINE\HARDWARE` | Volatile hive (memory only) |
| `HKEY_LOCAL_MACHINE\WindowsAppLockerCache` | `%SystemRoot%\System32\AppLocker\AppCache.dat` |
| `HKEY_LOCAL_MACHINE\ELAM` | `%SystemRoot%\System32\Config\Elam` |
| `HKEY_USERS\<SID of LocalService>` | `%SystemRoot%\ServiceProfiles\LocalService\Ntuser.dat` |
| `HKEY_USERS\<SID of NetworkService>` | `%SystemRoot%\ServiceProfiles\NetworkService\Ntuser.dat` |
| `HKEY_USERS\<SID of username>` | `\Users\<username>\Ntuser.dat` |
| `HKEY_USERS\<SID>_Classes` | `\Users\<username>\AppData\Local\Microsoft\Windows\Usrclass.dat` |
| `HKEY_USERS\.DEFAULT` | `%SystemRoot%\System32\Config\Default` |
| Virtualized `HKLM\SOFTWARE` | `\ProgramData\Packages\<PackageFullName>\<UserSid>\SystemAppData\Helium\Cache\<RandomName>.dat` |
| Virtualized `HKCU` | `\ProgramData\Packages\<PackageFullName>\<UserSid>\SystemAppData\Helium\User.dat` |
| Virtualized `HKLM\SOFTWARE\Classes` | `\ProgramData\Packages\<PackageFullName>\<UserSid>\SystemAppData\Helium\UserClasses.dat` |

Volatile hives (like `HKLM\HARDWARE`) are created at boot and never written to disk, virtualized hives are mounted on demand for packaged apps.

## REGISTRY only Keys

Keys that exist in the real REGISTRY view but are not reachable from standard hives:

- `\REGISTRY\A` (private keys used by some processes, including UWP apps)
- `\REGISTRY\WC` (Windows Containers / silos)

## Icons Meaning

### Symlink Icon

Symbolic link keys let the Configuration Manager redirect lookups to another key. They are created by passing `REG_CREATE_LINK` to `RegCreateKey` / `RegCreateKeyEx`. Internally, the link is stored as a `REG_LINK` value named `SymbolicLinkValue` that holds the target path. This value is nomrmally not visible in regedit.

RegKit marks keys as symbolic links when the registry reports a link target (done by checking for a symbolic link target during key enumeration).

Examples:
- `HKLM\SYSTEM\CurrentControlSet` -> `HKLM\SYSTEM\ControlSet00x`
- `HKEY_CURRENT_USER` -> `HKEY_USERS\<CurrentUserSID>`
- `HKEY_CURRENT_CONFIG` -> `HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current`

### Database Icon

RegKit marks keys that map to hive files listed under HKLM\SYSTEM\CurrentControlSet\Control\Hivelist (see
[A true hive is stored in a file.](https://scorpiosoftware.net/2022/04/15/mysteries-of-the-registry/)).

These hive-backed keys can be opened directly via "Open Hive File" (View menu or context menu). See [Hives and on-disk files](https://github.com/nohuto/regkit#hives-and-on-disk-files) for hive file paths.

## Trace Menu

There are three trace files which are quite similar, 23H2/24H2/25H2. I've done all of them on new installations. This will load trace files that contain registry paths in the kernel namespace, for example:

- `\\Registry\\Machine\\...`
- `\\Registry\\User\\<SID>\\...` (I've replaced my SID with a <CURRENT_USER_SID> placeholder)

It also normalizes those paths into standard hive paths (HKLM, HKU, HKCU), you can either use them for pure informational purposes or modify them. Note that WPR doesn't pass the type/data so you'll have to find that out on your own. Several ones are documented on my own in the [win-registry](https://github.com/nohuto/win-registry) repository (see 'Research' menu).

It's recommended that you create your own trace, as the templates are based on my system and IDs such as those for the disk won't be correct for your system. Follow the [wpr-wpa.md](https://github.com/nohuto/win-registry/blob/main/guide/wpr-wpa.md) guide to create a trace which regkit can use.

## Credits/References

[Mysteries-of-the-registry](https://scorpiosoftware.net/2022/04/15/mysteries-of-the-registry/) & [Windows-Internals-E7-P2](https://github.com/nohuto/windows-books/releases/download/7th-Edition/Windows-Internals-E7-P2.pdf) were used for better understanding of the Registry and the documentation, it's recommended to read through these if you want more detailed infomation, as this repository isn't intended to be a complete documeantation of the registry, and therefore only contains a summary of certain topics. [Registry-finder](https://registry-finder.com/) was used for UI inspiration/ideas and [TotalRegistry](https://github.com/zodiacon/TotalRegistry) for feature inspiration.