# MT_PropertyExchange — Overview (business perspective)

*Everything in this document is deduced from the code, scripts, test data and
install tooling in this repository; the original developer documentation was
lost. Statements that are inference rather than hard fact are marked as such.*

## What the toolkit does

MT_PropertyExchange transfers **document metadata between SAP and Microsoft
Office files**. It reads and writes the *document properties* (built-in ones
like Title/Subject/Author and any number of custom ones) of Word and Excel
files, and it can trigger Office macros that re-draw title blocks from those
properties.

In business terms: when an engineering document is managed in SAP DMS, the SAP
attributes (document number, revision, status, project, plant, …) must appear
*inside* the Word/Excel file — in its property fields and, via macros, in the
visible title block of the document. This toolkit is the bridge that makes
that happen automatically.

## Who calls it

Three integration styles are visible in the code:

1. **SAP frontend (interactive use).** The SAP GUI / SAP Business Client
   creates the COM object `MT_PropertyExchange.Globals` on the user's PC and
   calls methods like `ReadProperty`, `WriteProperty`,
   `ImportPropertiesFromText`, `GetAllCusProperties`, `RunMacro`. This is the
   path that broke when SAP switched to the **64-bit** Business Client — the
   reason for the current 64-bit port.
2. **Command line / scripts.** `MT_PropertyExchange.exe` wraps the same
   functionality for batch use (`-mode=r/w/d/i/e/g/n/m`).
3. **SEAL Systems conversion server (batch/server use).** The perl scripts in
   `SEAL_Conversion/` are working units for the SEAL DPF ("Digital Process
   Factory" / dpf4convert) server: during document conversion (e.g. to PDF),
   the server stamps the current SAP attributes into the Office file first.
   The scripts' history shows this replaced an older third-party tool
   ("MAINTEC" / `MAINTECSAP2Word.exe`) that could only handle the old
   Office 2003 formats; this toolkit was its successor supporting all Office
   formats.

## What the data looks like

The bundled test file `MTPropertyExchangeInstall/test/PropertyUpdateFile.txt`
shows a typical payload coming from SAP:

```
[PROPERTIES]
DocumentNo=T.TEST.ET31.3107/ALP001
RevisionNo=01
Status=Draft
DocumentTypeText=Flow chart
DocType=ALP
ProjectCode=T.TEST
PlantCode=ET31
Title1=Test project
Title2=Training Plant
Title3=EasyDMS Area
...
```

These are classic SAP DMS / *SAP Easy Document Management* attributes for
plant-engineering documents (project code, plant code, document type,
revision, status, multi-line title block). The toolkit writes each line as a
document property of the target file.

The Excel macro path searches for company PLM add-ins
(`PT_Addin_01_PLM.xlam`, `MT_Addin_01_PLM.xlam`, `plmUpdateServer.xla`) in
Office library folders — these add-ins (not part of this repository) contain
the macros that refresh title blocks/sheets from the updated properties.

## Typical use cases

| # | Use case | Toolkit feature used |
|---|----------|----------------------|
| 1 | SAP check-in/check-out: push current SAP attributes into the document | `ImportPropertiesFromText` / `WriteProperty` |
| 2 | Show/verify document attributes in SAP from the file | `ReadProperty`, `GetAllCusProperties`, `GetAllProperties` |
| 3 | Refresh the visible title block after attribute changes | `RunMacro` (with the PLM Office add-ins) |
| 4 | Batch stamping during server-side conversion (SEAL DPF) | `MT_PropertyExchange.exe -mode=i -txt=...` |
| 5 | Housekeeping: clear or delete properties | `DeleteProperty`, import of `empty_*.xml` / `delete_all.xml` |

## Product lineage (deduced from names, paths and copyrights)

Siemens VAI Metals Technologies → Siemens Metals Technologies ("Siemens MT")
→ Primetals Technologies. The code carries all three generations
(`vaiCompany=Siemens VAI...` test data, `Siemens MT` assembly info from 2012,
`\\cadlnz.primetals.com` in the old build scripts, add-in search paths for
both `Siemens\MT\Office` and `Primetals\Office`). The last functional change
before this repository was archived dates to ~2013
(`Globals.Version()` returned `2013-02-21`).

The core file-property engine for legacy formats is Microsoft's own sample
component **DSOFile** ("DSO OLE Document Properties Reader 2.1", included in
source form in `Dsofile/`), which Siemens shipped and registered with the
toolkit.

## Why the 64-bit port (2026)

The toolkit's COM registration was 32-bit only and `dsofile.dll` existed only
as a 32-bit build. The 64-bit SAP Business Client therefore could not create
`MT_PropertyExchange.Globals` at all. The port (see the main `README.md` and
`docs/ARCHITECTURE.md`) builds dsofile for x64, keeps the .NET parts AnyCPU,
and registers everything in **both** registry views, so old 32-bit callers
and the new 64-bit SAP client work side by side on the same PC.
