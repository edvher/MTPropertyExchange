# MT_PropertyExchange

Toolkit used by the SAP DMS/PLM integration to read and write document
properties of Microsoft Office files and to run Office macros. SAP (or any
other client) uses it in two ways:

* **COM**: `CreateObject("MT_PropertyExchange.Globals")`
  (CLSID `{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}`) — the primary way the SAP
  frontend calls the toolkit.
* **Command line**: `MT_PropertyExchange.exe -mode=... -file=...`
  (run `MT_PropertyExchange.exe -help` for the full reference).

## Why this branch exists — 64-bit support

The toolkit was historically built and registered **32-bit only**. The 64-bit
SAP Business Client cannot see 32-bit COM registrations (they live in the
`Wow6432Node` registry view) and cannot load the 32-bit native `dsofile.dll`
in-process, so `CreateObject("MT_PropertyExchange.Globals")` fails there.

This branch makes the toolkit bitness-complete:

* the native **dsofile** COM component builds for **x86 and x64** from the
  included sources (modern `dsofile.vcxproj`, VS 2022),
* the .NET assemblies stay **AnyCPU** and are registered in **both** registry
  views (32-bit *and* 64-bit `regasm`/`regsvr32`) by the new `install.bat`,
* everything is retargeted from .NET Framework 3.5 to **.NET Framework 4.8**
  (preinstalled on Windows 10/11 — no ".NET 3.5 optional feature" needed, and
  the CLR loaded into the calling 64-bit process is v4).

Result: legacy 32-bit callers (old SAP GUI, SEAL scripts) and the 64-bit SAP
Business Client work side by side on the same machine.

## Documentation

- **[docs/OVERVIEW.md](docs/OVERVIEW.md)** — business perspective: what the
  toolkit is for, who calls it (SAP frontend, command line, SEAL conversion
  server), what the data looks like, typical use cases, product history.
- **[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)** — technical reference:
  components, the full COM API with error codes, format handling
  (OPC vs. DSOFile), data formats, logging, the 32/64-bit registration model
  and a per-folder repository guide.

## Repository layout

| Path | Content |
| --- | --- |
| `MT_PropertyExchange.dll/` | VB.NET COM-visible class library `MT_PropertyExchangeDLL.dll` — the core. `FilePropIO2007.vb` handles `.docx`/`.xlsx` (pure .NET, `System.IO.Packaging`); `FilePropIO2003.vb` handles `.doc`/`.xls` via the native **dsofile** COM component; `Globals.vb` is the COM entry class. |
| `MT_PropertyExchange.exe/` | VB.NET console wrapper `MT_PropertyExchange.exe` around the same DLL. |
| `CommandLine/` | C# command-line argument parser used by the exe. |
| `Dsofile/` | C++ sources of Microsoft's *DSO OLE Document Properties Reader 2.1* (`dsofile.dll`, ProgID `DSOFile.OleDocumentProperties`). Builds x86 **and** x64 via `dsofile.sln` / `dsofile.vcxproj`. |
| `lib/` | Holds the generated `Interop.Dsofile.dll` (see `lib/README.md`). |
| `MTPropertyExchangeInstall/` | `install.bat` (installs + registers both bitnesses), `stage.bat` (collects build output into the installer layout), `test/` (smoke tests + sample `SAP.doc/.docx/.xls/.xlsx`). |
| `SEAL_Conversion/` | Historical SEAL/perl integration scripts (reference only). |

## Building

Prerequisites on the build machine:

* **Visual Studio 2022** with workloads *“.NET desktop development”* and
  *“Desktop development with C++”* (x86/x64 build tools + Windows 10/11 SDK).
* **Office PIAs** for the Word/Excel interop references (installed with
  MS Office; alternatively the *Office 2010 PIA Redistributable*). Any Office
  version at runtime works — the interop is out-of-process COM.
* NuGet access for the `log4net` package restore.

Steps (order matters — the C++ component provides the type library interop):

1. Build the native component, both platforms:
   `Dsofile\dsofile.sln` → **Release|x86** and **Release|x64**
   (outputs `Dsofile\bin\Win32\Release\dsofile.dll` and `Dsofile\bin\x64\Release\dsofile.dll`).
2. Build the managed solution:
   `MsoPropertyTransferUtils3.5.sln` → **Release|Any CPU**.
   On first build, `lib\Interop.Dsofile.dll` is generated automatically from
   the checked-in `Dsofile\Lib\dsofile.tlb` (signed with `dsofile_key.snk`);
   see `lib/README.md` if that step needs to be run manually.
3. Stage the installer payload:
   run `MTPropertyExchangeInstall\stage.bat`.
4. *(optional)* Build a **single self-extracting installer EXE**:
   run `make-installer.bat` → `Install\MTPropertyExchangeInstall-<timestamp>.exe`.
   One-time prerequisites (7-Zip + the `7zSD.sfx` module) are described in
   `7-zip\README.md`.

Or from a VS 2022 *Developer Command Prompt*:

```bat
msbuild Dsofile\dsofile.sln /p:Configuration=Release /p:Platform=x86
msbuild Dsofile\dsofile.sln /p:Configuration=Release /p:Platform=x64
msbuild /restore MsoPropertyTransferUtils3.5.sln "/p:Configuration=Release;Platform=Any CPU"
MTPropertyExchangeInstall\stage.bat
make-installer.bat
```

## Installing

**Single file:** copy `Install\MTPropertyExchangeInstall-<timestamp>.exe` to the
target machine and double-click it — it extracts itself and starts the
installer (UAC elevation is requested automatically).

**Folder:** alternatively copy the whole `MTPropertyExchangeInstall` folder
(after `stage.bat`) to the target machine and run `install.bat` as
administrator.

Either way the installer writes a full step-by-step log to
`%TEMP%\MT_PropertyExchange_install.log`, ends with a **green** console on
success or a **red** console (plus the full log on screen) on failure, and
waits for ENTER before closing. It

1. unregisters any previous installation (old single-bitness setups included),
2. copies the toolkit to `%ProgramFiles%\Siemens\MT_PropertyExchange`,
3. registers `dsofile.dll` **x64** with the 64-bit `regsvr32` and **x86** with
   the 32-bit (`SysWOW64`) `regsvr32`,
4. registers `MT_PropertyExchangeDLL.dll` with **both** the 64-bit and 32-bit
   .NET 4.x `regasm` (`/codebase /tlb`),
5. runs a COM activation smoke test with the 64-bit **and** the 32-bit script
   host.

`install.bat /u` uninstalls, `/register` / `/unregister` re-register only.

> **Note:** if the old `MT_PropertyExchangeDLL_x64_hack.reg` (DllSurrogate
> workaround) was ever applied on a machine, those `AppID`/`DllSurrogate`
> registry entries for CLSID `{41A13AC0-...}` should be removed — the clean
> dual registration replaces that hack.

## Testing

* `MTPropertyExchangeInstall\test\test-com.bat` — COM activation in both
  bitnesses (also run automatically by `install.bat`).
* `MTPropertyExchangeInstall\test\runTests.bat` — full read/write/export
  round-trips against the bundled `SAP.doc`, `SAP.docx`, `SAP.xls`, `SAP.xlsx`.
* Final acceptance: a property read/write from the 64-bit SAP Business Client.

Logs go to `%TMP%\PropertyExchange.log`
(configured in `MT_PropertyExchangeLog.config`, log4net).

## Known deviations from the historical build

* `Interactivity.vb` (a WinForms dialog shown when `MT_PropertyExchange.exe` was
  started without arguments) is not present in the repository; the exe now
  prints the command-line help instead. SAP integration is unaffected — it
  always calls the COM DLL or the exe with arguments.
* The solution-level `lib\` binaries (log4net, prebuilt interop) were never
  committed; log4net now comes from NuGet (2.0.15, patched successor of the
  original 1.2.11) and the interop is generated from the checked-in type
  library at build time.
* The old post-build steps that pushed installers to internal
  Siemens/Primetals network shares and built a 7-zip self-extractor were
  removed; `stage.bat` + `install.bat` replace them.
* `AssemblyVersion` stays `1.0.0.0` on purpose (COM registration stability);
  `AssemblyFileVersion` is bumped to `1.1.0.0`, and
  `MT_PropertyExchange.Globals.Version()` now returns `2026-07-20 12:00:00`
  so a successful 64-bit rollout is easy to verify via `test-com.bat`.
