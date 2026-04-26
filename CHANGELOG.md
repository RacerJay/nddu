# Changelog

All notable changes to this project will be documented in this file.

## ![Version](https://img.shields.io/badge/Version-v1.3.5-gold) 04-26-2026

- `Fixed:` **Path traversal hardening**: `sanitize_folder_name()` now rejects `.` and `..` in addition to the existing invalid-character blacklist, returning `"Default"` instead. Prevents a client name like `..` from resolving to a path outside the output folder.
- `Fixed:` **Hostname filename sanitization**: Device output filenames now run the hostname returned from `show running-config | include hostname` through `sanitize_folder_name()` before joining with the output folder. Prevents a compromised or misconfigured network device from writing output outside its run folder by returning a hostname containing path separators or traversal sequences.

## ![Version](https://img.shields.io/badge/Version-v1.3.4-blue) 04-24-2026

- `Added:` **Port Map editor**: An "Edit…" button in the Compare Runs dialog opens the port map CSV in the built-in file editor. If no file is selected, the sample file (`input/sample_port_map.csv`) opens by default. Save As defaults to CSV filter; "Set as Default" is hidden for port map files.
- `Added:` **Progress bar color on completion**: The output progress bar turns green on success, red on failure, and amber if the run was cancelled. Colors are consistent with the Go/Stop button palette.
- `Fixed:` **Missing interfaces in Excel report**: Interfaces that only appeared in `show interfaces description` (e.g., Loopback, FastEthernet, Vlan SVIs) were missing from the Interfaces sheet. Added `desc_map` to the merged interface set so all three command sources contribute.
- `Fixed:` **Truncated descriptions**: The Interfaces sheet now prefers the full description from `show interfaces description` over the truncated version from `show interface status`.
- `Fixed:` **Duplicate interface entries**: Added `Vl` → `Vlan` and `Fa`/`FastEthernet` → `FastEthernet` to the interface name normalizer so `Vl1` and `Vlan1` (and `Fa0/1` and `FastEthernet0/1`) merge into a single row.
- `Changed:` **Unified button/progress bar color palette**: Go button and progress bar now share the same green (`#5cb85c`), red (`#d9534f`), and amber (`#f0ad4e`) colors for visual consistency.

## ![Version](https://img.shields.io/badge/Version-v1.3.3-blue) 04-13-2026

- `Changed:` **Diff - Summary per-cell highlighting**: Previously the entire row was highlighted yellow when any field differed between Run A and Run B. Now only the specific changed column pairs (Hostname, IP, Platform, Version, Uptime) are highlighted, and the Device Type cell is highlighted on its own when changed. Makes it easier to see exactly which field differs.
- `Added:` **"No changes detected" status line** on every diff sheet that contains only its header when no changes were found. Applies to: Diff - Summary, Interfaces, Neighbors, VLANs, Routing, STP, MAC Summary, MAC Addresses, and Templates.

## ![Version](https://img.shields.io/badge/Version-v1.3.2-blue) 04-07-2026

- `Added:` **Port Map for Device Diff**: Optional CSV file mapping old interface names to new interface names for switch replacement scenarios (e.g., 24-port stack → 48-port switch where Gi1/0/1 becomes Gi1/0/25). Browse for the CSV in the Compare Runs dialog. Mappings are applied to all diff sheets that reference interfaces: Interfaces, Neighbors, VLANs, MAC Addresses, and Templates. Abbreviated names accepted (Gi1/0/1 instead of GigabitEthernet1/0/1). Sample file at `input/sample_port_map.csv`. Port map usage shown in the Diff - Run Info sheet.
- `Fixed:` **Diff - VLANs empty sheet**: The raw parser for `show vlan brief` returns a dict-of-dicts keyed by VLAN ID, not a list. The diff extractor now handles both formats.
- `Fixed:` **Diff - MAC Summary empty sheet**: Same dict-of-dicts issue as VLANs.
- `Fixed:` **Diff - Templates wrap text**: Interface columns in the Templates diff sheet now wrap text for long interface lists.

## ![Version](https://img.shields.io/badge/Version-v1.3.1-blue) 04-06-2026

- `Removed:` **CLI mode**: The command-line interface (`-cli`, `-d`, `-c`, `-ks`, `-ku`, `-xl`, `--components`, `--verbose`, `--combined`, `--force-redetect`, `--no-txt` flags) has been removed. nddu is now GUI-only. The last version with CLI support is v1.3.0 (archived). Removed imports: `argparse`, `getpass`. Removed functions: `check_cli_updates()`, `parse_args()`, `run_cli()`.
- `Changed:` **Help dialog** updated to describe GUI features (Excel Report, Client Manager, Device Diff, Device Type Cache, Keyring Credentials) instead of CLI flags and usage patterns.

## ![Version](https://img.shields.io/badge/Version-v1.3.0-blue) 04-04-2026

- `Added:` **Device Diff** (`Compare Runs…` button in Client Manager): Select any two prior runs for the same client and generate an Excel diff report comparing them. Two-pass device matching — primary by IP address, secondary by hostname — handles cases where a device's IP has changed between runs. The diff report contains 10 sheets: Run Info (side-by-side run metadata), Summary (per-device status with version/platform/uptime changes), Interfaces (added/removed/changed with descriptions, IPs, VLANs, speed, duplex), Neighbors (CDP changes), VLANs (add/remove/rename with interface assignments), Routing (route count changes per protocol), STP (mode, root bridge, blocking/forwarding), MAC Summary (per-VLAN count changes), MAC Addresses (individual MAC add/remove/move), and Templates (interface assignment changes). Color-coded rows: green = added, red = removed, yellow = changed. Sheets with no changes show only the header row.
- `Fixed:` **Excel table name validation**: Table names containing hyphens or special characters (e.g., from sheet titles like "Diff - Summary") caused Excel to flag the workbook as needing repair on open. Table names are now sanitized to contain only alphanumeric characters and underscores.

## ![Version](https://img.shields.io/badge/Version-v1.2.3-blue) 04-03-2026

- `Added:` **Client / Job Manager** (`Manage Clients` button in Actions row): New dialog for managing client output folders. Features include New, Rename, Archive/Delete (single dialog with Archive, Archive & Delete, Delete, Cancel options), and Report (XLSX). Run history displays as a collapsible tree — each timestamp row expands to show output files. Double-click or use the dynamic Open Folder/Open File button to open items with the OS default app. Right-click context menu on runs and files. Delete Run button removes a single timestamped run folder. Device Type Cache section shows cached IP → device type entries with Remove Entry (single IP) and Clear All Cache options.
- `Added:` **`run_info.json`** written to each run folder at the start of every Go run, recording the device file, command file, and timestamp. Used by the Client Manager to display run metadata without scanning output files.
- `Changed:` **Client Manager layout**: Left panel lists clients with action buttons; right panel shows runs tree and device cache group. Window sized at 720×620.
- `Changed:` **Archive/Delete dialog**: Custom `QDialog` replaces `QMessageBox` to guarantee left-to-right button order (Archive | Archive & Delete | Delete | Cancel) on all platforms.
- `Changed:` **Client Manager Report**: XLSX only (CSV removed). Includes Client, Run Timestamp, Device File, Command File, and Files columns. Files column lists all output files per run (newline-separated, wrap text). All cells vertically center-aligned.

## ![Version](https://img.shields.io/badge/Version-v1.2.2-blue) 04-03-2026

- `Changed:` **"Combined Output File" moved to Settings**: The checkbox has been removed from the main window Script Options and is now a persistent setting in the Settings dialog under the Output section. The value is saved to `nddu.ini` and applied on every run without needing to re-check it each session.
- `Changed:` **Settings stored in local `nddu.ini` file**: All application settings (output options, Excel components, remembered credentials, default input file paths) are now saved to `nddu.ini` next to `nddu.py` instead of the OS-native store (registry on Windows, plist on macOS). The file is portable with the application.
- `Changed:` **Input file paths always written to ini**: The active device and command file paths (including the hardcoded defaults on first launch) are always written to `nddu.ini` on startup. Reset restores them to the hardcoded defaults rather than removing the keys, so the ini always reflects the current state.
- `Changed:` **Credential sub-groups indented**: The Manual and Keyring credential field boxes are now visually indented under their respective radio buttons, making the hierarchy clearer.
- `Changed:` **Credential layout reordered**: The Keyring Credentials radio row now appears between the Manual fields and the Keyring fields (radio → fields → radio → fields), rather than both radio buttons appearing before both field groups.
- `Changed:` **"Remember" checkboxes gated to active credential type**: "Remember Username" is only enabled when Manual Credentials is selected; "Remember System Name" is only enabled when Keyring Credentials is selected. The inactive checkbox is greyed out.

## ![Version](https://img.shields.io/badge/Version-v1.2.1-blue) 04-03-2026

- `Fixed:` **Device type cache race condition**: When multiple devices were detected concurrently, concurrent threads would each load the cache, add their entry, then write — causing last-writer-wins data loss. Added a `threading.Lock` that serializes all cache reads and writes. The write path now re-reads the file inside the lock so no entries are overwritten.
- `Fixed:` **Device type fallback not cached**: When both Phase 1 and SSHDetect failed, the `cisco_ios` fallback was returned but never written to `device_cache.json`. Subsequent runs would re-attempt detection on every run. The fallback result is now saved to the cache like any other detected type.
- `Fixed:` **9300X crash — `show ip route summary` parser**: Lines like `Internal:  0 unicast next-hops` appearing after the route table header caused `invalid literal for int()`. Non-numeric data lines are now skipped. Added `try/except` guards around all three raw parsers (`show ip route summary`, `show spanning-tree summary`, `show mac address-table count`) so a parser failure produces a missing sheet section rather than crashing the entire device.
- `Fixed:` **Verbose password mask reveals password length**: Changed `'*' * len(password)` to the static string `[hidden]` for both password and enable password fields.
- `Changed:` **Report component panel layout**: Reorganized from three rows (label+buttons, 4 checkboxes, 4 checkboxes) to a compact horizontal layout with the label/All/None buttons on the left and checkboxes in two side-by-side columns of 4, reducing vertical height by ~40%.
- `Added:` **Client/Job name guard rails**: Maximum 64-character limit enforced at input. On focus-out, the name is sanitized in-place (invalid filesystem characters stripped). If the name was changed, a tooltip on the field shows what it was changed to.
- `Added:` **Device type cache verbose logging**: Cache reads log `Cache hit: {ip} -> '{type}'`; cache writes log `Cache saved:` or `Cache updated:` with the filename, making cache behavior visible in verbose mode.

## ![Version](https://img.shields.io/badge/Version-v1.2.0-blue) 03-29-2026

- `Changed:` **Replaced pyATS/Genie with ntc-templates (TextFSM) for structured parsing**: pyATS/Genie does not support Windows, preventing the Excel report feature from working on Windows machines. Replaced with `ntc-templates` — a cross-platform, community-maintained library of 500+ TextFSM templates for Cisco IOS, IOS-XE, NX-OS, and other vendors. This eliminates the pyATS dependency entirely and ensures identical structured parsing behavior on Windows, macOS, and Linux.
- `Fixed:` **Excel report now works on Windows**: pyATS could not be installed via pip on Windows (`ERROR: No matching distribution found for pyats`). All Genie-parsed fields (Summary, Interfaces, Neighbors, Routing, STP, MAC Addresses) were empty on Windows. With the ntc-templates migration, `pip install -r requirements.txt` succeeds on all platforms and all Excel sheets populate correctly.
- `Fixed:` **macOS GIL contention eliminated**: Genie's pure-Python regex parsing held the GIL during `send_command()`, blocking concurrent SSH reads on macOS and inflating runtimes to 5-6 minutes. TextFSM templates are lightweight and do not exhibit this issue. The deferred Phase 1/Phase 2 architecture (SSH collection then parsing) is retained for optimal parallelism.
- `Added:` **Raw parsers for commands without ntc-templates coverage**: `show spanning-tree summary`, `show ip route summary`, `show template`, `show port-profile`, `show mac address-table count`, and `show vlan brief` all use custom raw parsers. These are cross-platform and produce consistent output regardless of parser library availability.
- `Changed:` **Interfaces sheet** — removed "Method" column (not available in TextFSM output); "Vlan" column now shows `vlan_id` from `show interface status`.
- `Changed:` **Routing sheet** — simplified column layout; VRF column removed (raw parser currently reports single VRF).
- `Changed:` **All Excel sheet generators** rewritten to consume the TextFSM flat list-of-dicts format instead of Genie's nested dict format.
- `Changed:` `requirements.txt` — removed `pyats[library]>=23.0` and `genie>=23.0`; added `ntc-templates>=6.0.0`.
- `Added:` **Report Component Selection**: When "Excel Report" is checked, a component selection panel appears with checkboxes for each report section (Interfaces, Neighbors, VLANs, Routing, STP, MAC Addresses, MAC Summary, Templates). Only selected components' commands are run during inventory collection and only their sheets are included in the Excel report. Summary and Run Info are always included. "Select All" / "Deselect All" buttons for quick toggling. CLI support via `--components` flag (comma-separated list).
- `Added:` **Device Type Caching per Client/Job**: Auto-detected device types are now cached in `output/{Client}/device_cache.json`. Subsequent runs for the same client skip auto-detection for cached devices, significantly reducing run times. A "Force Re-detect" checkbox (visible when Excel Report is enabled) ignores the cache and re-detects all devices. CLI support via `--force-redetect` flag.
- `Fixed:` **Excel Report — empty component selection blocked**: If "Excel Report" is checked but all components are deselected, clicking Go now shows an error in the log and aborts. Previously, an empty selection silently fell back to generating a full report.

## ![Version](https://img.shields.io/badge/Version-v1.1.9-blue) 03-15-2026

- `Fixed:` **Auto-detect speed — replaced SSHDetect with fast show-version detection**: Netmiko's `SSHDetect` iterates 40+ device types (~7s each command check) to identify a device. When a device matches at priority <99 (e.g., Catalyst 4500 IOS-XE chassis switches whose `show version` says "Cisco IOS Software" but not "Cisco IOS XE Software"), SSHDetect cannot early-exit and checks **every** remaining type — taking **108-115 seconds** per device. Replaced with a two-phase approach: Phase 1 connects as generic `cisco_ios`, runs `show version` once, and pattern-matches for IOS-XE/NX-OS/XR/ASA/IOS (~5-8s). Phase 2 falls back to SSHDetect only for non-Cisco devices. Results: Catalyst 4500 chassis switches drop from **115s to ~8s** detection; NX-OS devices (which failed SSHDetect entirely on Mac, causing a 314s total processing time with fallback+reconnect) now detect correctly in ~8s with no reconnect needed.
- `Fixed:` **NX-OS detection failure on macOS**: SSHDetect's channel read would fail entirely on NX-OS devices on macOS ARM, falling back to `cisco_ios` and requiring a post-connection reconnect as `cisco_nxos` (adding ~180s). The new fast detection connects successfully as `cisco_ios`, detects NX-OS from `show version`, and returns `cisco_nxos` directly — eliminating the reconnect.

## ![Version](https://img.shields.io/badge/Version-v1.1.8-blue) 03-15-2026

- `Fixed:` **macOS ARM hang — Genie `show vlan brief` catastrophic regex backtracking**: Root cause identified and fixed. Genie's `show vlan brief` parser contains regex patterns that trigger catastrophic backtracking on devices with long port-list lines (>100 chars). A single device took **192.7 seconds** to parse on Mac M1 Pro vs <1 second on Windows. Replaced with `_parse_vlan_brief_raw()` — a simple line-splitting parser that produces both the Genie IOS-XE schema (`vlan` key for Excel) and a clean schema (`vlans` key for VLAN poll). Added `_SKIP_GENIE` set to bypass Genie for known-problematic commands. Mac runtime dropped from ~6 minutes to ~2:28, matching Windows (~2:25).
- `Added:` **Per-command Genie parse timing**: Verbose log now reports `[TIMING] Genie parse '{command}' took {N}s` for any parser exceeding 1 second, plus timing milestones for connection, user commands, inventory commands, post-processing, and total elapsed time.
- `Added:` **Platform-aware `show vlan` / `show vlan brief` skip**: On IOS/IOS-XE, `show vlan` (full) is skipped — it outputs every port membership per VLAN which can be thousands of lines on large chassis switches; `show vlan brief` covers the same data. On NX-OS, `show vlan brief` is skipped since `show vlan` is the canonical command.
- `Changed:` **Reverted v1.1.6/v1.1.7 GIL workarounds**: Removed deferred Genie parsing (Phase 1/2 architecture), `max_workers` reduction (20 → 8), macOS QOS thread promotion (pthread/ctypes), per-VLAN MAC poll deferral to post-phase, log drain timer increase (150ms → 500ms), and `time.sleep(0)` GIL yields — all workarounds that masked the Genie regex issue. Code is now ~130 lines cleaner with the same performance.
- `Fixed:` **Log messages**: Unicode `→` arrows replaced with ASCII `->` for consistent encoding across terminals and log files.

## ![Version](https://img.shields.io/badge/Version-v1.1.7-blue) 03-15-2026

- `Fixed:` **macOS ARM hang — Genie `show vlan brief` catastrophic regex backtracking**: Identified the root cause of the macOS hanging issue: Genie's `show vlan brief` parser contains regex patterns that trigger catastrophic backtracking on devices with long port-list lines (>100 chars). A single device took **192.7 seconds** to parse on Mac M1 Pro vs <1 second on Windows. Total script time went from ~6 minutes (Mac) to ~2:28 (matching Windows at ~2:25). Replaced with `_parse_vlan_brief_raw()` — a simple line-splitting parser that produces both the Genie IOS-XE schema (`vlan` key for Excel) and a clean schema (`vlans` key for VLAN poll). Added `_SKIP_GENIE` set to bypass Genie for known-problematic commands.
- `Added:` **Per-command Genie parse timing**: Verbose log now reports `[TIMING] Genie parse '{command}' took {N}s` for any parser exceeding 1 second, enabling rapid identification of future slow parsers.
- `Changed:` **GIL workarounds from v1.1.6**: Deferred Genie parsing (Phase 1/2 architecture), `max_workers` reduction (20 → 8), macOS QOS thread promotion (pthread/ctypes), per-VLAN MAC poll deferral to post-phase, and log drain timer increase (150ms → 500ms) — all added as workarounds for the Genie hang. These masked the symptom but added ~130 lines of complexity. Reverted in v1.1.8 after root cause was found.

## ![Version](https://img.shields.io/badge/Version-v1.1.6-blue) 03-14-2026

- `Added:` **Excel — MAC Addresses sheet**: New sheet collects every MAC address table entry across all devices. Sources: `show mac address-table` for MAC/VLAN/port/type; `show arp` (IOS-XE, Genie-parsed) and `show ip arp` (NX-OS, raw-parsed) correlated by MAC address to populate the IP Address column. Handles both IOS-XE (filters Genie junk interface keys `ip`/`ipx`/`assigned`/`other`; uses `entry_type`) and NX-OS (uses `mac_type`). Columns: Hostname, IP, VLAN, MAC Address, Type, Port, IP Address.
- `Added:` **Excel — MAC Summary sheet**: Companion summary sheet from `show mac address-table count`. Shows one row per VLAN (Dynamic / Static / Total) plus a grand TOTAL row per device. Platforms that only return aggregate counts (chassis IOS, N77, N9K) are queried per-VLAN using `show mac address-table count vlan N`. Raw parsers handle all known format variants (standard IOS-XE, WS-C4510R+E chassis IOS, N77 NX-OS, N9K NX-OS) including multi-line Dynamic/Static fields and Local+Remote accumulation. Both `show ip arp` and `show mac address-table count` added to inventory command set.
- `Added:` **Excel — Summary sheet**: PID column (from `show inventory`) added before the Serial column; Serial column renamed to "Serial Number". For switch stacks, all member PIDs and serial numbers are listed in order (newline-separated, wrap text) so each member is visible without expanding the cell.
- `Fixed:` **Device type detection — IOS-XE misdetected as `cisco_ios`**: SSHDetect scores `cisco_ios` for WS-C4510R+E and WS-C4506-E platforms running IOS-XE 3.x (no NX-OS patterns in output). A post-connection `show version` check was already in place for NX-OS correction; extended to also detect `IOS-XE`/`IOS XE` in the output and silently correct `auto_device_type` to `cisco_xe` without tearing down or re-establishing the connection. Prevents these devices from being parsed with the wrong Genie OS and producing empty structured data.
- `Fixed:` **VLANs — empty list on `cisco_ios` (chassis IOS-XE)**: When Genie fails to parse `show vlan brief` for the `ios` OS (returns empty dict), a raw-text fallback now scans the output for lines matching `^\s*(\d+)\s+\S` to extract VLAN IDs. Used for per-VLAN MAC polling and any other VLAN-dependent logic.
- `Fixed:` **Routing — EIGRP/OSPF/BGP routes missing**: Genie IOS-XE nests multi-instance protocols one level deeper than flat protocols: `sources['eigrp'] = {'1': {subnets, networks}}` rather than a flat `sources['eigrp 1']`. The routing sheet builder now detects this nesting (by absence of `networks`/`subnets` at the top level) and iterates the instance sub-dict, emitting one row per instance.
- `Fixed:` **Routing — `internal` routes not appearing**: An explicit `continue` statement was skipping `internal` protocol entries. Removed; `internal` now appears as its own row consistent with all other protocol types.
- `Changed:` **Excel — all sheets**: All data cells are now vertically middle-aligned. Wrap Text is also enabled on the STP "Root Bridge For" column and the Templates "Interfaces" column.
- `Changed:` **Log messages**: Added `[N/M]` prefix to inventory-phase messages that were previously missing it, for consistent alignment with connection-phase messages.
- `Fixed:` **UI hang during multi-device runs**: Several compounding causes addressed:
  - Reduced `max_workers` from `min(50, cpu_count * 4)` to `min(20, cpu_count * 2)` to reduce GIL contention from concurrent paramiko crypto threads.
  - Replaced per-message Qt signal emissions from background threads with a `queue.SimpleQueue` + `QTimer` drain (150 ms interval). The drain collects all pending messages, builds a single HTML string, and calls `insertHtml()` once per cycle with `setUpdatesEnabled(False)` to suppress intermediate layout passes.
  - Removed the `textChanged` signal handler that was resetting horizontal scroll position after every insertion; scroll reset is now done once at the end of each drain cycle.
  - Added `time.sleep(0)` at the end of both command-execution loops and the per-VLAN MAC polling loop to explicitly yield the GIL between SSH round-trips.
  - Added Genie parser pre-warming: before the `ThreadPoolExecutor` starts, `try_genie_parse()` is called once for each device type in `_GENIE_OS_MAP` with an empty string. This forces all lazy parser module imports to complete in the main thread, preventing concurrent first-import contention (Python import lock serialisation) when multiple `cisco_xe` devices connect simultaneously.

## ![Version](https://img.shields.io/badge/Version-v1.1.5-blue) 03-14-2026

- `Added:` **GUI — Client/Job dropdown**: The Client/Job text field is now an editable dropdown (`QComboBox`) pre-populated with existing client names from the `output/` folder. Selecting from the list fills in an existing client; typing a new name creates a new client folder on run. Bare timestamp folders (no-client runs) are excluded from the list. The dropdown is refreshed from disk each time the application starts.
- `Added:` **Excel — Templates sheet**: New "Templates" sheet collects interface templates (IOS/IOS-XE `show template`) and port profiles (NX-OS `show port-profile`). Both commands added to the inventory command set. Raw parsers handle both formats since Genie has no support for either command. Columns: Template Name, Type (User/Built-in for IOS-XE; profile type for NX-OS), Status (N/A for IOS-XE; enabled/disabled for NX-OS), Usage (interface count), Interfaces (comma-separated bound/assigned interfaces).
- `Added:` **Status — Structured data collection progress**: Added a normal-log `info` message ("Collecting structured data for {device} (N inventory commands)...") shown as soon as inventory command collection begins, so there is visible feedback during the delay between "Connected" and "Done". The JSON-saved confirmation was also promoted from verbose-only to the normal log.
- `Fixed:` **Interfaces — VLAN SVI IP Address N/A (NX-OS)**: Genie's NX-OS `show ip interface brief` parser nests VLAN entries one level deeper than Ethernet/Loopback/Port-channel interfaces (`Vlan254 → {vlan_id: {254: {ip_address, interface_status}}}` vs flat `{ip_address, interface_status}`). The `ip_map` builder in `_sheet_interfaces` now flattens these nested VLAN entries before lookup, so IP addresses and protocol status populate correctly for all SVI interfaces. Existing JSON files are fixed without re-running.
- `Fixed:` **Routing — Instance column always showed `-`**: Genie embeds the instance ID directly in the `route_source` key string (`"ospf 1"`, `"bgp 65001"`) rather than nesting it. The previous code looked for a nested dict and always fell to the flat branch. The key is now split on the first space (IOS/IOS-XE) to separate the protocol name from the instance ID.
- `Fixed:` **Routing — Notes column was blank**: `_routing_notes()` was receiving the full key string (e.g., `"ospf 1"`) which never matched the `"ospf"` comparison. Now receives only the base protocol name after the split.
- `Fixed:` **Routing — NX-OS per-protocol breakdown missing**: NX-OS Genie uses `best_paths → {proto: count}` (not `route_source`). Added handling for this schema: one row per protocol with its path count in the Routes column; instance split from hyphen-notation keys (`eigrp-1` → Protocol=`eigrp`, Instance=`1`). TOTAL row still shows `routes / paths` from `total_routes` / `total_paths`.
- `Changed:` **Routing sheet** — rebuilt with per-protocol rows and protocol-specific detail.
  - Columns: **Hostname, IP, VRF, Protocol, Instance, Routes, Notes**.
  - IOS / IOS-XE: all active protocols listed individually. Flat protocols (connected, static, RIP) show as single rows; instance-based protocols (OSPF, BGP, EIGRP, IS-IS) show one row per process/AS number with the instance ID in its own column. Zero-count flat protocols (application, etc.) are hidden.
  - **Notes column**: OSPF — Intra/Inter/Ext1/Ext2/NSSA breakdown; BGP — External/Internal/Local breakdown; IS-IS — L1/L2 breakdown.
  - **TOTAL row** appended per VRF from `total_route_source`.
  - NX-OS: single TOTAL row with `routes / paths` count (per-protocol breakdown not available from `show ip route summary` on NX-OS).
  - `Fixed:` Previous code only read flat protocols and silently produced wrong counts for nested protocols (OSPF, BGP, EIGRP returned `N/A` because `proto_data` was a nested dict, not a flat one).

## ![Version](https://img.shields.io/badge/Version-v1.1.4-blue) 03-13-2026

- `Fixed:` **Interfaces — "eth 25G" prefix in Description (NX-OS)**: Some NX-OS platforms emit real `Type` and `Speed` columns in `show interface description` (e.g. `type=eth`, `speed=25G`, `description=RSVD FORTIGATE`). Added `_NXOS_INTF_TYPE_KEYWORDS` set; when `type` matches a known NX-OS interface type keyword, only the `description` field is used. The join-all-three fallback is preserved for platforms/versions that omit the Type/Speed columns and Genie misidentifies description words.
- `Fixed:` **Interfaces — VLAN/Loopback IP Address N/A (NX-OS)**: `_normalize_intf_name()` used case-sensitive prefix matching, causing lowercase names from `show ip interface brief` (`loopback0`, `vlan10`) to fail normalization and never merge with their mixed-case counterparts from `show interface status` (`Loopback0`, `Vlan10`). Normalization is now case-insensitive throughout.
- `Fixed:` **Interfaces — duplicate rows (NX-OS)**: NX-OS `show ip interface brief` returns abbreviated names (`Eth1/11`, `Lo0`, `Po2`) while `show interface status` returns canonical full names (`Ethernet1/11`, `Loopback0`, `Port-channel2`). Added `_normalize_intf_name()` which expands abbreviations before merging; duplicates no longer appear and data from both sources is combined into a single row.
- `Fixed:` **Interfaces — missing descriptions on NX-OS routed Ethernet ports**: Routed ports that appear only in `show ip interface brief` (not in `show interface status`) now receive descriptions from `show interface description`, with keys also normalized before lookup.
- `Changed:` **Interfaces sheet** — unified both platforms onto a single code path using `show interface status` as the primary interface inventory source. Covers physical ports, SVIs, port-channels, and (via `show ip interface brief`) Loopbacks and Tunnels that don't appear in `show interface status`.
- `Removed:` **Interfaces sheet** — `OK?` column removed.
- `Added:` **Interfaces sheet** — `Vlan`, `Duplex`, `Speed`, and `Type` columns (SFP/transceiver type) from `show interface status`, consistent across IOS, IOS-XE, and NX-OS.
- `Changed:` **Interfaces sheet** — Description source priority: full text from `show interfaces description` (IOS-XE) first; `name` field from `show interface status` as fallback (handles truncation gracefully).
- `Added:` `show interface status` added to `EXCEL_COMMANDS` for all platforms (previously NX-OS only).
- `Fixed:` **Device type detection — post-connection validation**: When `SSHDetect` returns `cisco_ios` with a zero `cisco_nxos` score (making the `potential_matches` correction ineffective), a `show version` is now sent immediately after connecting. If the output contains `NX-OS` or `Nexus`/`Cisco`, the connection is torn down and re-established with `cisco_nxos`. Covers VPC pairs and other cases where the SSH fingerprint produces no NX-OS score at all.
- `Fixed:` **NX-OS SVI/L3 IP addresses**: Genie's `show ip interface brief` parser silently drops all non-`unnumbered` entries on NX-OS. Raw command output is now retained during both the user and inventory command loops. After all commands complete, a regex fallback parser fills any missing interfaces into `structured['show ip interface brief']`, merging with any entries Genie did capture.
- `Fixed:` **Device type detection** — After `SSHDetect` returns `cisco_ios`, `potential_matches` is now checked for any non-zero `cisco_nxos` score and corrected automatically. NX-OS patterns (`NX-OS`, `Nexus`) cannot appear in real IOS output, making any positive score a reliable correction signal. No additional connection required.
- `Fixed:` **Interfaces (NX-OS)** — Genie's `show ip interface brief` parser only captures `unnumbered` entries on NX-OS; all routed and SVI interfaces are silently dropped. NX-OS devices now use `show interface status` as the primary interface source, which correctly parses all interface types (Ethernet, Vlan, port-channel). IP addresses are cross-referenced from `show ip interface brief` where available.
- `Fixed:` **Interfaces (IOS-XE descriptions)** — Genie's NX-OS `show interface description` parser misidentifies multi-word descriptions as `type`/`speed`/`description` columns. All three fields are now joined to reconstruct the full description string.
- `Fixed:` **STP — VLANs N/A on NX-OS** — `total_statistics` is Optional in the NX-OS STP schema and absent on some output variants. When missing, VLAN count is derived by counting per-VLAN entries in the `mode` dict instead.
- `Added:` `show interface status` to `EXCEL_COMMANDS` for NX-OS interface coverage.
- `Fixed:` **Summary** — NX-OS `show version` uses `platform → {hardware, software, kernel_uptime}` schema; model, serial, version, and uptime now populate correctly for NX-OS devices.
- `Fixed:` **Interfaces** — NX-OS uses `show interface description` (no 's'); both command names now tried for the description lookup. NX-OS combined `interface_status` field used as fallback when separate `status`/`protocol` fields are absent. Devices with no parsed interface data now appear with N/A rather than being silently dropped.
- `Fixed:` **Neighbors** — NX-OS places neighbor IPs in `interface_addresses`; IOS-XE uses `entry_addresses`. Both keys now tried in order.
- `Fixed:` **VLANs** — NX-OS uses `show vlan` (not `show vlan brief`) with a different schema (`vlans → {id} → {name, state, interfaces}`). Both commands added to `EXCEL_COMMANDS`; sheet auto-selects the correct schema.
- `Fixed:` **Routing** — NX-OS `show ip route summary` schema has no `route_source` per-protocol breakdown; falls back to `total_routes / total_paths` summary row.
- `Fixed:` **STP** — `total_statistics` is Optional in NX-OS schema; missing stats now show N/A rather than causing parse errors.
- `Added:` **STP** — "Root Bridge For" column populated from `root_bridge_for` field (present on both IOS-XE and NX-OS when the device is root for any VLANs).
- `Added:` `show vlan` and `show interface description` to `EXCEL_COMMANDS` for NX-OS support.

## ![Version](https://img.shields.io/badge/Version-v1.1.3-blue) 03-13-2026

- `Added:` Excel workbook (`nddu_report.xlsx`) generated automatically in the run folder when Structured Output is enabled; built from per-device JSON files after all devices complete.
- `Added:` Sheets: **Run Info** (client, timestamp, version, device count), **Summary** (hostname, IP, platform, IOS version, serial, uptime, stack members), **Interfaces** (`show ip interface brief`), **Neighbors** (`show cdp neighbors detail`), **VLANs** (`show vlan brief`), **Routing** (`show ip route summary`), **STP** (`show spanning-tree summary`).
- `Added:` `openpyxl>=3.1.0` to `requirements.txt`.
- `Added:` `_get()` safe nested-dict accessor used throughout report generation — missing or unsupported data on any platform falls back to `N/A` without raising exceptions.
- `Fixed:` VLANs sheet — Genie schema for `show vlan brief` uses `vlan` (not `vlans`) as the top-level key, with entries keyed as `vlan1`/`vlan10` and fields `vlan_name`/`vlan_status`/`vlan_port`.
- `Added:` Interfaces sheet — Description column populated from `show interfaces description`; `show interfaces description` added to `EXCEL_COMMANDS`.
- `Changed:` All Excel sheets use native Excel Tables (`TableStyleMedium9`) with alternating row stripes, auto-filter dropdowns, and sortable columns.

## ![Version](https://img.shields.io/badge/Version-v1.1.2-blue) 03-13-2026

- `Changed:` "Device Type Auto-detection" checkbox replaced by "Structured Output" — enabling it activates device type auto-detection, Genie parsing, inventory commands, and per-device JSON output as a single option.
- `Added:` `EXCEL_COMMANDS` — built-in inventory command set that runs alongside `Commands.txt` when Structured Output is enabled; commands already in the user's file are deduplicated automatically; unsupported commands on a given platform are silently skipped.
- `Added:` `try_genie_parse()` — parses already-captured raw command output locally using pyATS/Genie without re-sending the command to the device; returns `None` on any failure so raw `.txt` output is never affected.
- `Added:` Per-device `.json` file written alongside each `.txt` file when Genie-parsed data is available; contains `meta` (host, hostname, device type, timestamp, version) and `structured` (command → parsed dict).
- `Changed:` CLI flag renamed from `-a`/`--autodetect` to `-s`/`--structured`.
- `Added:` `pyats[library]>=23.0` to `requirements.txt` (required for Structured Output).
- `Fixed:` `try_genie_parse()` — `get_parser()` returns a `(class, kwargs)` tuple in this version of Genie; the class was previously called as a tuple causing a silent `TypeError` and no JSON output.
- `Changed:` Per-device JSON files are now written to a `json/` subfolder inside the run folder, separating them from `.txt` output files.

## ![Version](https://img.shields.io/badge/Version-v1.1.1-blue) 03-13-2026

- `Added:` Client / Job field in the GUI (Script Options) and `--client` CLI argument — when provided, output is organized under `output/{Client}/` instead of directly under `output/`.
- `Added:` `sanitize_folder_name()` helper — strips filesystem-invalid characters from the client name before constructing the path, with a `Default` fallback for blank results.
- `Fixed:` Duplicate `output_folder` construction in `on_go()` — folder is now created once with the correct client-aware path before credentials are read.
- `Changed:` "Open Output Folder" button navigates to the client subfolder when a client name is set and that folder exists.

## ![Version](https://img.shields.io/badge/Version-v1.1.0-blue) 03-12-2026

- `Fixed:` Thread-safety gap — `active_connections` and `output_files` are now guarded by a `threading.Lock` across all concurrent `process_device()` calls and the `cancel()` method.
- `Fixed:` `detect_device_type()` — removed a dead outer time-guard (elapsed was always ~0s), fixed unreachable `if best_match not in ALLOWED_TYPES` branch, and wired `AUTO_DETECT_TIMEOUT` into the SSHDetect connection parameters where the timeout actually takes effect.
- `Fixed:` CLI mode — a `Worker` was being instantiated with empty credentials and immediately discarded before the credential prompt; the redundant construction is removed.
- `Changed:` Log format — per-device lines now use a zero-padded `[N/M]` prefix for aligned columns; the disconnection, command count, and output file path are collapsed into a single `Done` completion line.
- `Added:` Retry on timeout — timed-out devices are retried once with extended connection and read timeouts (`CONN_TIMEOUT_RETRY`, `READ_TIMEOUT_RETRY`) before being skipped; authentication failures still fail immediately.
- `Changed:` `max_workers` formula updated from `min(32, cpu_count + 8)` to `min(50, cpu_count * 4)` to better utilize available threads for I/O-bound SSH workloads.
- `Fixed:` Verbose mode — credentials are now masked (`***`) in log output; password length is shown but value is never logged in plaintext.
- `Added:` New Script Option for Netmiko device type auto-detection (Cisco IOS, IOS-XE, IOS-XR, NX-OS, ASA, WLC).
- `Added:` "New version" notification prompts.
- `Changed:` Input files now ignore everything after the first `#` sign found in a given line, it no longer has to be just the first character.

## ![Version](https://img.shields.io/badge/Version-v1.0.1-green) 08-27-2025

- `Fixed:` Hostnames were being truncated to 16 characters

## ![Version](https://img.shields.io/badge/Version-v1.0.0-green) 04-17-2025 - Initial Release

- `Added:` Added comprehensive docstrings, added type hints throughout the code for improved organization and enhanced maintainability.
- `Changed:` Keyring Tools v2.0-rc4 (04-06-2025)
- `Added:` Dark mode skin.
- `Changed:` Cleaned up constants and variables (PEP 8).
- `Changed:` New unified logging facility.
- `Changed:` Go button becomes Stop button to cancel the running script.
- `Changed:` GUI improvements.
- `Changed:` Valid commands cannot be longer than 256 characters.
- `Changed:` Input/Output files and paths are now defined in configuration variables.
- `Fixed:` Error handling when there are no valid input device IPs or commands.
- `Fixed:` Enable state is now retained upon script completion.
- `Added:` Script options to have verbose output, and a combined output file.

## ![Version](https://img.shields.io/badge/Version-v1.0.0_rc4-gold) 03-21-2025

- `Added:` Improved input file lint checking.
- `Fixed:` GUI - Improved error handling, and general functional improvements.
- `Added:` A credential validation process is now performed to mitigate account lockout.
- `Changed:` -cli, --cli-defaults behavior.

## ![Version](https://img.shields.io/badge/Version-v1.0.0_rc3-gold) 03-17-2025 - 🍀 Lucky Leprechaun Edition

- `Changed:` Fixed Netmiko timeout issue.
- `Changed:` Increased maximum number of threads to execute concurrently.

## ![Version](https://img.shields.io/badge/Version-v1.0.0_rc2-gold) 03-16-2025

- v1.0.0 Release Candidate 2

## ![Version](https://img.shields.io/badge/Version-v1.0.0_rc1-gold) 03-15-2025

- v1.0.0 Release Candidate 1

## **MultiSwitchDoc**: `v2.2 (04-22-2023)` [🌎 Earth Day Edition]

- nddu's CLI-only predecessor script
