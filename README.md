# RegKit

RegKit is a native Windows Registry editor written in C++ using the Win32 API and common controls for performance reasons. Based on its features, customization options, and the fact that it's FOSS, it's the best alternative to regedit. It currently supports Windows Vista through Windows 11 (32bit & 64bit versions).

Note that native RegEdit can't run alongside RegKit, as RegKit uses `RegEdit_RegEdit` window class (required for jump support), which causes Regedit to see this window as an existing instance and exits instead of opening another one.

## Differences to Native RegEdit

RegKit adds functionality that standard regedit doesn't support:

- A real REGISTRY root view in addition to the standard root keys
- [Theme modes](https://noverse.dev/docs/regkit/overview/#theme-presets) (System/Light/Dark) and custom theme presets (edit colors, import/export `.rktheme`)
- Custom font support
- Custom [icon support](https://noverse.dev/docs/regkit/overview/#icon-sets) (has 4 sets installed by default)
- Symbolic link detection (`SymbolicLinkValue` value with the link target)
- Hive backed key detection using hivelist key & open Hive File (opens the backing hive file)
- [Trace presets](https://noverse.dev/docs/regkit/overview/#trace-menu) (23H2/24H2/25H2 - see below), used for "Read on boot" column
- Default presets, this shows default data from new installations
- Extra root keys toggle, exposes additional predefined keys that RegEdit typically doesn't show, such as `HKEY_PERFORMANCE_DATA` (live performance counter data produced on demand, not stored in a hive file) and related keys like `HKEY_PERFORMANCE_TEXT`/`HKEY_PERFORMANCE_NLSTEXT` for e.g. counter name strings (read more [here](https://learn.microsoft.com/en-us/windows/win32/perfctrs/using-the-registry-functions-to-consume-counter-data))
- Run with [SYSTEM/TI rights](https://noverse.dev/docs/regkit/overview/#rights-and-elevation)
- Favorites import/export
- Comment column for values with import/export support
- Loading/unloading hives
- Local/remote/offline registry
- Undo/redo, copy/paste (entire keys), replace, performant 'Find'
- Find can search Root Keys, the real REGISTRY root, and Trace values independently
- Address bar accepts multiple registry path formats (abbreviated HK*, full root, regedit address bar, `.reg` header, PowerShell drive/provider, escaped)
- Copy Key Path As menu for the same formats (to copy/paste into the address bar)
- Copy Value Name / Copy Value Data from value context menus
- Tab control
- Tab session restore (Save Tabs / Clear Tabs on Exit), including cached Find results
- Filter bar (value list filter)
- History view
- Option to save/forget previous key tree state
- Simulated keys toggle (from traces)
- Compare Registries (compare two registry sources or `.reg` files and see differences)
- `.reg` / hive file/folder drag and drop support
- Read only mode
- Miscellaneous common functionalities

## Keyboard Shortcuts

### Files & Registries

| Shortcut | Action |
| --- | --- |
| `Ctrl+N` | Open local registry tab |
| `Ctrl+R` | Connect to remote registry |
| `Ctrl+O` | Open offline registry |
| `Ctrl+Shift+O` | Open `.reg` file |
| `Ctrl+S` | Save current editable tab |
| `Ctrl+E` | Export |

### Windows/Tabs

| Shortcut | Action |
| --- | --- |
| `Ctrl+Shift+N` | Open new RegKit window |
| `Ctrl+W` | Close current tab |
| `Ctrl+Tab` / `Ctrl+Shift+Tab` | Next/previous tab |
| `Ctrl+1` - `Ctrl+9` | Select tab by position |
| `Alt+F4` | Close RegKit |

### Navigation

| Shortcut | Action |
| --- | --- |
| `Alt+Left` / `Alt+Right` | Navigate back/forward |
| `Alt+Up` | Navigate to parent key |
| `Backspace` | Navigate back (outside text fields) |
| `Ctrl+G` | Go to a registry path |
| `Ctrl+L` / `Alt+D` | Focus address bar |
| `Ctrl+Shift+V` | Paste into address bar and navigate |
| `Ctrl+K` | Show and focus filter bar |
| `F6` / `Shift+F6` | Cycle forward/backward through the visible panes |
| `Tab` | Switch between key tree & value list |
| `F5` | Refresh |

### Editing

| Shortcut | Action |
| --- | --- |
| `F7` | New key |
| `F2` | Rename selected key or value |
| `Delete` | Delete selected registry items |
| `Ctrl+C` / `Ctrl+V` | Copy / paste registry items |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| `Ctrl+A` | Select all in the focused list or text field |
| `Ctrl+Shift+C` | Copy current key name |
| `Alt+Enter` | Permissions for the selected key |

### View

| Shortcut | Action |
| --- | --- |
| `Ctrl+Shift+H` | Show or hide History pane |
| `F10` | Activate menu bar |
| `Alt+F` / `Alt+E` / `Alt+V`... | Open matching menu (F = File, A = Favorites) |
| `Shift+F10` / `Menu` | Open context menu for focused item |
| `F1` | Open help |

### Find & Replace

| Shortcut | Action |
| --- | --- |
| `Ctrl+F` | Find |
| `Ctrl+H` | Replace |
| `Enter` | Start search / replace |
| `Escape` | Close window |
| `Ctrl+Shift+G` | Open selected result in a new tab |
| `Enter` (results list) | Open selected result |

### Filter Bar

| Shortcut | Action |
| --- | --- |
| `Escape` | Clear filter |
| `Enter` / `Down` | Move to value list |

### Value Editor

| Shortcut | Action |
| --- | --- |
| `Enter` | Confirm |
| `Ctrl+Enter` | Insert a line break in multi line fields |
| `Escape` | Cancel |

## Command Line

### regedit

| Command | Notes |
| --- | --- |
| `regkit file.reg` | Import `.reg` file (asks for confirmation) |
| `regkit /s file.reg` | Import without confirmation |
| `regkit /e file.reg <key>` | Export key to a `.reg` file |
| `regkit /a file.reg <key>` | Same as `/e` (for compatibility) |
| `regkit /c` `/m` `/l:file` `/r:file` | Accepted and ignored |

### reg

Using `reg` here is optional, means both `regkit reg query` & `regkit query` work.

| Command | Notes |
| --- | --- |
| `add <key> [/v name \| /ve] [/t type] [/s sep] [/d data] [/f]` | `/f` overwrites an existing value |
| `delete <key> [/v name \| /ve \| /va] [/f]` | Without `/v` the whole key tree is removed |
| `query <key> [/v name \| /ve] [/s]` | `/s` recurses into subkeys |
| `copy <src> <dst> [/s] [/f]` | `/s` copies subkeys too |
| `export <key> <file.reg> [/y]` | `/y` overwrites an existing file |
| `import <file.reg>` | |
| `save <key> <file.hiv> [/y]` | Needs the backup privilege |
| `restore <key> <file.hiv>` | Needs the restore & backup privileges |
| `load <key> <file.hiv>` / `unload <key>` | Mounts/releases a hive file |
| `compare <key1> <key2> [/s]` | Exit code `0` = identical, `2` = different |
| `/reg:32` `/reg:64` | Selects the 32/64 bit registry view |

### regkit (additions)

| Command | Notes |
| --- | --- |
| `regkit <key>` | Open the window at that key |
| `regkit --goto <key>` | The same, in explicit form |
| `regkit --restart-system` | Relaunch under the SYSTEM account |
| `regkit --restart-ti` | Relaunch under TrustedInstaller |
| `regkit --help` | Print usage text |

## Theme Presets

It includes built in presets and a theme editor to customize colors, presets can also be saved, exported, and imported as `.rktheme` files.

### Examples

#### Default Dark

![](https://github.com/nohuto/regkit/blob/main/assets/images/default-dark.png?raw=true)

#### Default Light

![](https://github.com/nohuto/regkit/blob/main/assets/images/default-light.png?raw=true)

##### W7 Light

![](https://github.com/nohuto/regkit/blob/main/assets/images/default-light-w7.png?raw=true)

#### Gruvbox Dark

![](https://github.com/nohuto/regkit/blob/main/assets/images/gruvbox-dark.png?raw=true)

#### Kanagawa Wave

![](https://github.com/nohuto/regkit/blob/main/assets/images/kanagawa-wave.png?raw=true)

## Icon Sets

RegKit comes with multiple icon sets and supports loading your own icons, you can switch them via `Options > Icons`.

Built in sets:

- Phosphor + RegEdit (default)
- Phosphor
- Lucide
- Material Symbols

You can set your own ico set via `%LOCALAPPDATA%\Noverse\RegKit\icons` (use naming of icons listed below). If `icons\dark` and `icons\light` exist, regkit uses them for dark/light modes, if not it will use the root `icons` folder for both modes.

### Previews

| Icon | Phosphor | Lucide | Material Symbols |
| --- | --- | --- | --- |
| `back` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/back.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/back.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/back.ico?raw=true" width="16" height="16"> |
| `binary` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/binary.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/binary.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/binary.ico?raw=true" width="16" height="16"> |
| `copy` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/copy.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/copy.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/copy.ico?raw=true" width="16" height="16"> |
| `database` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/database.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/database.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/database.ico?raw=true" width="16" height="16"> |
| `delete` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/delete.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/delete.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/delete.ico?raw=true" width="16" height="16"> |
| `export` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/export.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/export.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/export.ico?raw=true" width="16" height="16"> |
| `folder` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/folder.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/folder.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/folder.ico?raw=true" width="16" height="16"> |
| `folder-sim` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/folder-sim.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/folder-sim.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/folder-sim.ico?raw=true" width="16" height="16"> |
| `forward` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/forward.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/forward.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/forward.ico?raw=true" width="16" height="16"> |
| `local-registry` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/local-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/local-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/local-registry.ico?raw=true" width="16" height="16"> |
| `offline-registry` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/offline-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/offline-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/offline-registry.ico?raw=true" width="16" height="16"> |
| `paste` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/paste.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/paste.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/paste.ico?raw=true" width="16" height="16"> |
| `redo` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/redo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/redo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/redo.ico?raw=true" width="16" height="16"> |
| `refresh` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/refresh.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/refresh.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/refresh.ico?raw=true" width="16" height="16"> |
| `remote-registry` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/remote-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/remote-registry.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/remote-registry.ico?raw=true" width="16" height="16"> |
| `replace` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/replace.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/replace.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/replace.ico?raw=true" width="16" height="16"> |
| `search` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/search.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/search.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/search.ico?raw=true" width="16" height="16"> |
| `symlink` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/symlink.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/symlink.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/symlink.ico?raw=true" width="16" height="16"> |
| `text` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/text.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/text.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/text.ico?raw=true" width="16" height="16"> |
| `undo` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/undo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/undo.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/undo.ico?raw=true" width="16" height="16"> |
| `up` | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/phosphor/light/up.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/up.ico?raw=true" width="16" height="16"> | <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/materialsymbols/light/up.ico?raw=true" width="16" height="16"> |

## Icons Meaning

### Symlink Icon <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/symlink.ico?raw=true" width="16" height="16">

A key created with `REG_OPTION_CREATE_LINK` is a registry symbolic link key, which let the Configuration Manager redirect lookups to another key. Internally, the link is saved as a `REG_LINK` value named `SymbolicLinkValue` that holds the path.

RegKit displays keys as symbolic links when the registry reports a link (done by checking for a symbolic link during key enumeration), the value is usually not visible in regedit.

Examples:
- `HKLM\SYSTEM\CurrentControlSet` -> `HKLM\SYSTEM\ControlSet00x`
- `HKEY_CURRENT_CONFIG` -> `HKLM\SYSTEM\CurrentControlSet\Hardware Profiles\Current`
- `HKU\S-1-5-18` -> `HKU\.DEFAULT`
- `HKLM\SOFTWARE\Wow6432Node\Classes` -> `HKLM\SOFTWARE\Classes\Wow6432Node`

### Database Icon <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/database.ico?raw=true" width="16" height="16">

Used to mark keys that map to hive files listed under `HKLM\SYSTEM\CurrentControlSet\Control\Hivelist` (see "[A true hive is stored in a file.](https://scorpiosoftware.net/2022/04/15/mysteries-of-the-registry/)"). These (hive backed) keys can be opened directly via '*Open Hive File*' (menu). See [Hives and on-disk files](https://noverse.dev/docs/regkit/registry-internals/registry-fundamentals/#hives-and-on-disk-files) for hive file paths.

### Simulated Key Icon <img src="https://github.com/nohuto/regkit/blob/main/assets/icons/lucide/light/folder-sim.ico?raw=true" width="16" height="16">

Keys displayed as simulated are virtual entries created from trace files when a key exists in a trace but not in the actual hive view. They're displayed with the *folder-sim* icon so you can differ them from real keys. Creating or modifying a value in a simulated key will create the key path on demand.

## Trace Menu

There are three trace files which are quite similar, `23H2`/`24H2`/`25H2`. I've done all of them on new installations. Trace loading supports multiple active traces at once and shows "Read on boot" as `Yes (<traceName>, ...)`.

The trace key menu shows the kernel paths as they appear in the trace (for example `REGISTRY\\MACHINE\\...`), but trace data is also shown under the root keys. Registry symbolic links (the `SymbolicLinkValue` targets) are also resolved so trace values show up under linked keys (including `CurrentControlSet` and other link keys). It can also [simulate missing keys](https://noverse.dev/docs/regkit/overview/#simulated-key-icon) for trace only data (optional "Simulated Keys" view toggle), you can either use traces for informational purposes or modify them.

Note that WPR doesn't pass the type/data so you'll have to find that out on your own.

It's recommended that you create your own trace, as the templates are based on my system and IDs such as those for the disk won't be correct for your system. Follow the [Boot Registry Activity](https://noverse.dev/docs/regkit/guides/wpr-wpa/) guide to create a trace which regkit can use.

Loading traces affects startup time and memory consumption, therefore, it's recommended to either load only one trace or none at all if you don't use them frequently (loading a trace takes only a few seconds, so it's better to load it when needed than to keep it active all the time).

## Default Menu

Default presets are `.reg` exports that fill the value list's `Default` column with data from installation images. If a value is included in the registry but not in the loaded defaults, it'll be displayed as `(Missing)`. Currently all (beside two 25H2 exports were directly read from the hive files of images)

`HKLM-SYSTEM-IMAGE.reg` for example is from `Windows\System32\config\SYSTEM`, `HKCU-DEFAULT-IMAGE.reg` from `Users\Default\NTUSER.DAT`, which is the template used when a user profile is created. 

These are the exact builds for each file:

| Release | Edition | Architecture | Build |
| --- | --- | --- | --- |
| [Windows Vista RTM](https://github.com/nohuto/regkit/tree/main/assets/defaults/WVista%20Business%20x64%20-%206.0.6000.16386) | Business | x64 | `6.0.6000.16386` |
| [Windows 7 RTM](https://github.com/nohuto/regkit/tree/main/assets/defaults/W7%20Professional%20x64%20-%206.1.7600.16385) | Pro | x64 | `6.1.7600.16385` |
| [Windows 8](https://github.com/nohuto/regkit/tree/main/assets/defaults/W8%20Pro%20x64%20-%206.2.9200.16384) | Pro | x64 | `6.2.9200.16384` |
| [Windows 8.1](https://github.com/nohuto/regkit/tree/main/assets/defaults/W8.1%20Pro%20x64%20-%206.3.9600.16384) | Pro | x64 | `6.3.9600.16384` |
| [Windows 10 21H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W10%2021H2%20Home%20x64%20-%2010.0.19044.3086) | Home | x64 | `10.0.19044.3086` |
| [Windows 10 22H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W10%2022H2%20Home%20x64%20-%2010.0.19045.6456) | Home | x64 | `10.0.19045.6456` |
| [Windows 11 21H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W11%2021H2%20Home%20x64%20-%2010.0.22000.978) | Home | x64 | `10.0.22000.978` |
| [Windows 11 22H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W11%2022H2%20Home%20x64%20-%2010.0.22621.963) | Home | x64 | `10.0.22621.963` |
| [Windows 11 23H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W11%2023H2%20Home%20x64%20-%2010.0.22631.6060) | Home | x64 | `10.0.22631.6060` |
| [Windows 11 24H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W11%2024H2%20Home%20x64%20-%2010.0.26100.9168) | Home | x64 | `10.0.26100.9168` |
| [Windows 11 25H2](https://github.com/nohuto/regkit/tree/main/assets/defaults/W11%2025H2%20Home%20x64%20-%2010.0.26200.8037) | Home | x64 | `10.0.26200.8037` |
| [Windows 11 26H1](https://github.com/nohuto/regkit/tree/main/assets/defaults/W11%2026H1%20Home%20x64%20-%2010.0.28000.2704) | Home | x64 | `10.0.28000.2704` |

## Rights and Elevation

RegKit can relaunch itself under different security contexts as many registry areas are protected by ACLs and/or owned by TI (TrustedInstaller). Some keys are owned by TI, and only that SID has write permissions (SYSTEM may be read only). If a key is readable but writes fail with access denied, check the owner and ACLs, if the owner is TI, use the TI mode, if it is SYSTEM, use SYSTEM. Use the '*Options*' menu to restart with higher rights or to make the app always relaunch with them on startup.

These levels can bypass protections, use them only when you understand the possible impact.

- Restart as Admin: uses UAC elevation for a standard elevated token
- Restart as SYSTEM: uses an elevated process to duplicate a SYSTEM token, then creates a new RegKit process in the active session
- Restart as TI: uses SYSTEM to start/query the TI service, duplicates its token, then launches RegKit with that token

SYSTEM rights are for example needed for reading keys such as `HKLM\SAM\SAM`, `HKLM\SECURITY\Policy`, TI rights are for example needed to write in keys like `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing`.
