# GroupMon

VB.NET Windows service that watches Active Directory group membership and emails alerts when members are added or removed. Versioned trees from 0.9 through 2.0 include the GroupMonSvc worker, MSI setup projects, and companion Test / GroupStatus apps.

**Source last updated:** 2012-07-07 · **Language:** VB.NET · **Target:** .NET Framework · **Output:** Windows service + tools

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| GroupMonSvc | VB.NET | Windows service | AD group membership watcher |
| Test / GroupStatus | VB.NET | WinExe | Companion test/status apps |

## How to open

Open the versioned `.sln` under the desired tree (e.g. `2.0`) in Visual Studio 2005 or later.

## Requirements

- Visual Studio 2005 to 2010

## Attribution and provenance

Working copy from my Historical Dev folder `GroupMon-VB`.
No third-party source attribution markers were identified during review.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
