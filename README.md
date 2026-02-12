# RegKit

RegKit is a native Windows Registry editor written in C++ using the Win32 API and common controls for performance reasons. Based on its features, customization options, and the fact that it's FOSS, it's the best alternative to regedit.

## Table of Content

- [Differences to Default RegEdit](https://github.com/nohuto/regkit#differences-to-default-regedit)
- [Theme Presets](https://github.com/nohuto/regkit#theme-presets)
- [Icon Sets](https://github.com/nohuto/regkit#icon-sets)
  - [Icon Set Previews](https://github.com/nohuto/regkit#icon-set-previews)
- [Icons Meaning](https://github.com/nohuto/regkit#icons-meaning)
  - [Symlink Icon](https://github.com/nohuto/regkit#symlink-icon)
  - [Database Icon](https://github.com/nohuto/regkit#database-icon)
  - [Simulated Key Icon](https://github.com/nohuto/regkit#simulated-key-icon)
- [Trace Menu](https://github.com/nohuto/regkit#trace-menu)
- [Default Menu](https://github.com/nohuto/regkit#default-menu)
- [Rights and Elevation](https://github.com/nohuto/regkit#rights-and-elevation)
- [Registry Fundamentals](https://github.com/nohuto/regkit#registry-fundamentals)
  - [Standard hives & REGISTRY Comparison](https://github.com/nohuto/regkit#standard-hives--registry-comparison)
  - [REGISTRY only Keys](https://github.com/nohuto/regkit#registry-only-keys)
  - [Keys, values, and naming](https://github.com/nohuto/regkit#keys-values-and-naming)
  - [Registry value types](https://github.com/nohuto/regkit#registry-value-types)
  - [Root keys and logical structure](https://github.com/nohuto/regkit#root-keys-and-logical-structure)
  - [Hives and on-disk files](https://github.com/nohuto/regkit#hives-and-on-disk-files)
- [Credits/References](https://github.com/nohuto/regkit#creditsreferences)


## Differences to Default RegEdit

RegKit adds functionality that standard regedit doesn't support/expose:

- A real REGISTRY root view in addition to standard hives
- [Theme modes](https://github.com/nohuto/regkit#theme-presets) (System/Light/Dark) and custom theme presets (edit colors, import/export `.rktheme`)
- Custom font support
- Custom [icon support](https://github.com/nohuto/regkit#icon-sets) (has 4 sets installed by default)
- Symbolic link detection (`SymbolicLinkValue` value with the link target)
- Hive backed key detection using hivelist key & open Hive File (opens the backing hive file)
- [Trace presets](https://github.com/nohuto/regkit#trace-menu) (23H2/24H2/25H2 - see below), used for "Read on boot" column
- Default presets, this shows default data from new installations
- Extra hives toggle, exposes additional predefined keys that RegEdit typically doesn't show, such as `HKEY_PERFORMANCE_DATA` (live performance counter data produced on demand, not stored in a hive file) and related keys like `HKEY_PERFORMANCE_TEXT`/`HKEY_PERFORMANCE_NLSTEXT` for e.g. counter name strings (read more [here](https://learn.microsoft.com/en-us/windows/win32/perfctrs/using-the-registry-functions-to-consume-counter-data))
- Run with [SYSTEM/TI rights](https://github.com/nohuto/regkit#rights-and-elevation)
- Favorites import/export
- Comment column for values with import/export support
- Loading/unloading hives
- Local/remote/offline registry
- Undo/redo, copy/paste (entire keys), replace, performant 'Find'
- Find can search Standard Hives, the real REGISTRY root, and Trace values independently
- Address bar accepts multiple registry path formats (abbreviated HK*, full root, regedit address bar, .reg header, PowerShell drive/provider, escaped)
- Copy Key Path As menu for the same formats (to copy/paste into the address bar)
- Copy Value Name / Copy Value Data from value context menus
- Tab control
- Tab session restore (Save Tabs / Clear Tabs on Exit), including cached Find results
- Filter bar (value list filter)
- History view
- Option to save/forget previous key tree state
- Simulated keys toggle (from traces)
- Compare Registries (compare two registry sources or .reg files and see differences)
- .reg / hive file/folder drag and drop support
- Research menu (redirections to [win-registry](https://github.com/nohuto/win-registry))
- Miscellaneous common functionalities

## Theme Presets

RegKit includes built in presets and a theme editor to customize colors (backgrounds, text, selection, borders, focus). Presets can be saved, exported, and imported as `.rktheme` files to share themes across machines. Examples:

`Ayu Dark`:

![](https://github.com/nohuto/regkit/blob/main/assets/images/ayu-dark.png?raw=true)

`Catppuccin Latte`:

![](https://github.com/nohuto/regkit/blob/main/assets/images/catppuccin-latte.png?raw=true)

`Everforest Dark`:

![](https://github.com/nohuto/regkit/blob/main/assets/images/everforest-dark.png?raw=true)

`Kanagawa Dragon`:

![](https://github.com/nohuto/regkit/blob/main/assets/images/kanagawa-dragon.png?raw=true)

I haven't spent much time setting them up properly, some may not be perfect yet. You're able to edit each of these via the menu.

## Icon Sets

RegKit comes with multiple icon sets and supports user provided icons. Switch sets from `Options > Icons`.

Built-in sets:
- Lucide (default)
- Tabler
- Fluent UI
- Material Symbols

You can set your own ico set via `%LOCALAPPDATA%\Noverse\RegKit\icons`. If `icons\dark` and `icons\light` exist, regkit uses them for dark/light modes, if not it will use the root `icons` folder for both modes.

Required filenames: `back.ico`, `binary.ico`, `copy.ico`, `database.ico`, `delete.ico`, `export.ico`, `folder.ico`, `folder-sim.ico`, `forward.ico`, `local-registry.ico`, `offline-registry.ico`, `paste.ico`, `redo.ico`, `refresh.ico`, `remote-registry.ico`, `replace.ico`, `search.ico`, `symlink.ico`, `text.ico`, `undo.ico`, `up.ico`.

### Icon Set Previews

| Icon | Lucide | Tabler | Fluent UI | Material Symbols |
| --- | --- | --- | --- | --- |
| `back` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/back.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/back.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/back.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/back.ico?raw=true" width="16" height="16"> |
| `binary` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/binary.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/binary.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/binary.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/binary.ico?raw=true" width="16" height="16"> |
| `copy` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/copy.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/copy.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/copy.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/copy.ico?raw=true" width="16" height="16"> |
| `database` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/database.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/database.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/database.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/database.ico?raw=true" width="16" height="16"> |
| `delete` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/delete.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/delete.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/delete.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/delete.ico?raw=true" width="16" height="16"> |
| `export` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/export.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/export.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/export.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/export.ico?raw=true" width="16" height="16"> |
| `folder` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/folder.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/folder.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/folder.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/folder.ico?raw=true" width="16" height="16"> |
| `folder-sim` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/folder-sim.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/folder-sim.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/folder-sim.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/folder-sim.ico?raw=true" width="16" height="16"> |
| `forward` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/forward.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/forward.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/forward.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/forward.ico?raw=true" width="16" height="16"> |
| `local-registry` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/local-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/local-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/local-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/local-registry.ico?raw=true" width="16" height="16"> |
| `offline-registry` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/offline-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/offline-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/offline-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/offline-registry.ico?raw=true" width="16" height="16"> |
| `paste` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/paste.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/paste.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/paste.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/paste.ico?raw=true" width="16" height="16"> |
| `redo` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/redo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/redo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/redo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/redo.ico?raw=true" width="16" height="16"> |
| `refresh` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/refresh.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/refresh.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/refresh.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/refresh.ico?raw=true" width="16" height="16"> |
| `remote-registry` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/remote-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/remote-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/remote-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/remote-registry.ico?raw=true" width="16" height="16"> |
| `replace` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/replace.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/replace.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/replace.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/replace.ico?raw=true" width="16" height="16"> |
| `search` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/search.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/search.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/search.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/search.ico?raw=true" width="16" height="16"> |
| `symlink` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/symlink.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/symlink.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/symlink.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/symlink.ico?raw=true" width="16" height="16"> |
| `text` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/text.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/text.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/text.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/text.ico?raw=true" width="16" height="16"> |
| `undo` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/undo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/undo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/undo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/undo.ico?raw=true" width="16" height="16"> |
| `up` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/up.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/tabler/light/up.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/fluentui/light/up.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/up.ico?raw=true" width="16" height="16"> |


## Icons Meaning

### Symlink Icon <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/symlink.ico?raw=true" width="16" height="16">

A key created with `REG_OPTION_CREATE_LINK` is a registry symbolic link key, symbolic link keys let the Configuration Manager redirect lookups to another key. They're created by passing `REG_CREATE_LINK` to `RegCreateKey` / `RegCreateKeyEx`. Internally, the link is stored as a `REG_LINK` value named `SymbolicLinkValue` that holds the target path. This value is usually not visible in regedit.

RegKit displays keys as symbolic links when the registry reports a link target (done by checking for a symbolic link target during key enumeration).

Examples:
- `HKLM\SYSTEM\CurrentControlSet` -> `HKLM\SYSTEM\ControlSet00x`
- `HKEY_CURRENT_USER` -> `HKEY_USERS\<CurrentUserSID>`
- `HKEY_CURRENT_CONFIG` -> `HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current`

### Database Icon <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/database.ico?raw=true" width="16" height="16">

RegKit marks keys that map to hive files listed under HKLM\SYSTEM\CurrentControlSet\Control\Hivelist (see "[A true hive is stored in a file.](https://scorpiosoftware.net/2022/04/15/mysteries-of-the-registry/)").

These hive-backed keys can be opened directly via "*Open Hive File*" (view menu or context menu). See [Hives and on-disk files](https://github.com/nohuto/regkit#hives-and-on-disk-files) for hive file paths.

### Simulated Key Icon <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/folder-sim.ico?raw=true" width="16" height="16">

Keys displayed as simulated are virtual entries created from trace files when a key exists in a trace but not in the actual hive view. They're displayed with the *folder-sim* icon so you can differ them from real keys. Creating or modifying a value in a simulated key will create the key path on demand.

## Trace Menu

There are three trace files which are quite similar, 23H2/24H2/25H2. I've done all of them on new installations. Trace loading supports multiple active traces at once and shows "Read on boot" as `Yes (TraceName, ...)`.

The trace key menu shows the kernel paths as they appear in the trace (for example `REGISTRY\\MACHINE\\...`), but trace data is also shown in the standard hives. Registry symbolic links (the `SymbolicLinkValue` targets) are resolved so trace values appear under linked keys (including `CurrentControlSet` and other link keys), and kernel-only roots like `REGISTRY\\A` or `REGISTRY\\WC` remain available. It can also [simulate missing keys](https://github.com/nohuto/regkit#simulated-key-icon) for trace-only data (optional "Simulated Keys" view toggle). You can either use traces for informational purposes or modify them (simulated keys are created on demand).

Note that WPR doesn't pass the type/data so you'll have to find that out on your own. Several ones are documented on my own in the [win-registry](https://github.com/nohuto/win-registry) repository (see 'Research' menu).

It's recommended that you create your own trace, as the templates are based on my system and IDs such as those for the disk won't be correct for your system. Follow the [wpr-wpa.md](https://github.com/nohuto/win-registry/blob/main/guide/wpr-wpa.md) guide to create a trace which regkit can use.

> [!WARNING]
> Loading traces affects startup time and memory consumption. Therefore, it's recommended to either load only one trace or none at all if you don't use them frequently (loading a trace takes only a few seconds, so it's better to load it when needed than to keep it active all the time). Parsing them doesn't affect the UI/display time, as a different thread is used for it.

## Default Menu

Default presets are `.reg` exports that fill the value list's `Default` column with data from new installations. If a value is included in the registry but not in the defaults list, it'll be displayed as `(Missing)`.

> [!WARNING]
> Loading your entire *Computer* export in here isn't recommended as it'll take a long time to parse the file. Therefore split top level keys into smaller parts.

The current built in list isn't complete, I'll expand it over time.

## Rights and Elevation

RegKit can relaunch itself under different security contexts because many registry areas are protected by ACLs and/or owned by TI (TrustedInstaller). Some keys are owned by TI, and only that SID has write permissions (SYSTEM may be read-only). If a key is readable but writes fail with access denied, check the owner and ACLs. If the owner is TI, use the TI mode, if it is SYSTEM, use SYSTEM. Use the Options menu to restart with higher rights or to make the app always relaunch with them on startup.

> [!CAUTION]
> These levels can bypass protections, use them only when you understand the impact.

- Restart as Admin: uses UAC elevation for a standard elevated token
- Restart as SYSTEM: uses an elevated process to duplicate a SYSTEM token, then creates a new RegKit process in the active session
- Restart as TI: uses SYSTEM to start/query the TI service, duplicates its token, then launches RegKit with that token

SYSTEM rights are for example needed for reading keys such as `HKLM\SAM\SAM`, `HKLM\SECURITY\Policy`, TI rights are for example needed to write in keys like `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing`.

## Registry Fundamentals

### Standard hives & REGISTRY Comparison

RegEdit shows five common hives `HKEY_LOCAL_MACHINE`, `HKEY_USERS`, `HKEY_CURRENT_USER`, `HKEY_CLASSES_ROOT`, and `HKEY_CURRENT_CONFIG`. Internally, all registry keys are rooted at a single object named `\REGISTRY` in the Object Manager namespace. Native APIs (NtOpenKey/ZwOpenKey) can access paths under `\REGISTRY` directly. The registry actually exposes nine root keys (including performance and local settings roots) but most tools only show the common five.

You can query the REGISTRY key using WinDbg `!reg query \REGISTRY`.

### REGISTRY only Keys

Keys that exist in the real REGISTRY view but are not reachable from standard hives:

- `\REGISTRY\A` - private keys used by some processes, including UWP apps
- `\REGISTRY\WC` - Windows Containers / silos, used by modern registry virtualization and differencing hives

### Keys, values, and naming

The registry is a database that looks a lot like a filesystem, keys are like directories, values are like files, and a key can contain both subkeys and values. Values are typed, have a name, and live under a key. Each key also has one unnamed value, displayed as `(Default)`.

### Registry value types

Most values are `REG_DWORD`, `REG_BINARY`, or `REG_SZ`, but the registry supports 12 value types.

Some values are stored with extra flag bits in the upper 16 bits (e.g. `0x20000`, `0x40000`). These aren't new base types, the actual base type is `type & 0xFFFF`, and regkit displays them as `REG_* (0xXXXX)` (for example `0x20001` is `REG_SZ` with a flag, `0x20004` is `REG_DWORD`, and `0x40007` is `REG_MULTI_SZ`). These flagged types are included in the Find > Data Types filter. `RegQueryValueEx` returns a `DWORD` type, and in multiple cases the high 16 bits were non-zero while the low 16 bits matched a documented `REG_*` constant. Masking with `0xFFFF` consistently produced a known base type, and the returned data layout matched that base type (e.g., UTF-16 multi-strings for `REG_MULTI_SZ`, 32-bit integers for `REG_DWORD`). Note that this behavior was determined based on observed values and isn't validated by official Microsoft documentation, it's just a personal assumption.

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
| `HKEY_CURRENT_CONFIG` | `HKCC` | Stores some information about the current hardware profile (deprecated, "Hardware profiles are no longer supported in Windows, but the key still exists to support legacy applications that might depend on its presence.") | `HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current` (legacy, Yosifovich shows `Hardware\Profiles\Current`, but that's a typo in his blog) |
| `HKEY_PERFORMANCE_DATA` | `HKPD` | Live performance counter data, available only via APIs | - |
| `HKEY_PERFORMANCE_TEXT` | `HKPT` | Performance counter names/descriptions in US English | - |
| `HKEY_PERFORMANCE_NLSTEXT` | `HKPNT` | Performance counter names/descriptions in the OS language | - |

Notes:
- `HKEY_CURRENT_USER` maps to the logged-on user hive (`Ntuser.dat`) and is created per-user at logon.
- `HKEY_CLASSES_ROOT` also contains UAC VirtualStore data, it isn't a simple link.
- `HKEY_PERFORMANCE_*` keys aren't stored in hive files and aren't visible in Regedit. They are provided by Perflib through registry APIs like `RegQueryValueEx`.
- SYSTEM = `S-1-5-18`, LocalService = `S-1-5-19`, NetworkService = `S-1-5-20`

Low level view of the REGISTRY ([*](https://projectzero.google/2024/10/the-windows-registry-adventure-4-hives.html)):

![](https://github.com/nohuto/regkit/blob/main/assets/images/REGISTRYview.png?raw=true)

### Hives and on-disk files

On disk, the registry is a set of hive files, not a single file. The Configuration Manager records loaded hive paths under `HKLM\SYSTEM\CurrentControlSet\Control\Hivelist` (WinDbg cmd to get HiceAddr etc. = `!reg hivelist`) as they are mounted. Each hive is a PRIMARY file plus `.LOG<1/2>` (also possible to only be a `.LOG` if REG_HIVE_SINGLE_LOG) used during flushing/crash recovery. The mapping below is from Windows Internals (some hives are volatile or virtualized):

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

Volatile hives (like `HKLM\HARDWARE`) are created at boot and never written to disk (in paged pool, lost after reboot), virtualized hives are mounted on demand for packaged apps. Hives are loaded at boot or explicitly via `NtLoadKey` / `RegLoadKey` (SeRestorePrivilege required).

## Credits/References

[Mysteries-of-the-registry](https://scorpiosoftware.net/2022/04/15/mysteries-of-the-registry/), [Windows-Internals-E7-P2](https://github.com/nohuto/windows-books/releases/download/7th-Edition/Windows-Internals-E7-P2.pdf), [NTRegistryImplimentation](https://empyreal96.github.io/nt-info-depot/Windows-Kernel-Internals/NTRegistryImplimentation.pdf), [The Windows Registry Adventure](https://projectzero.google/2024/10/the-windows-registry-adventure-4-hives.html) were used for better understanding of the Registry and the documentation. It's recommended to read through these if you want more detailed infomation, as this repository isn't intended to be a complete documeantation of the registry, and therefore only contains a summary of certain topics. [Registry-finder](https://registry-finder.com/) was used for UI inspiration/ideas and [TotalRegistry](https://github.com/zodiacon/TotalRegistry) for feature inspiration. [Tabler icons](https://tabler.io/icons), [Lucide](https://lucide.dev/) for the icons.

I recommend reading through [Mysteries-of-the-registry](https://scorpiosoftware.net/2022/04/15/mysteries-of-the-registry/), [Windows-Internals-E7-P2](https://github.com/nohuto/windows-books/releases/download/7th-Edition/Windows-Internals-E7-P2.pdf), [NTRegistryImplimentation](https://empyreal96.github.io/nt-info-depot/Windows-Kernel-Internals/NTRegistryImplimentation.pdf), [The Windows Registry Adventure](https://projectzero.google/2024/10/the-windows-registry-adventure-4-hives.html) if you want a better understanding about it, since this repo is only a "short summary" and is missing on details.