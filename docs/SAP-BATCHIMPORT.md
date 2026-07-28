# ZBATIMP / ZPLM_BATCHIMPORT — how SAP uses the toolkit and what a document must contain

*Derived from the ABAP source of report `ZPLM_BATCHIMPORT` (transaction
`ZBATIMP`) and this repository's code. Line references refer to the 2026
source extract analysed during the 64-bit migration.*

## Invocation chain

1. **DLL presence check** — `FORM check_dlls`: `CREATE OBJECT
   'MT_PropertyExchange.Globals'` must succeed, then `CALL METHOD
   'GetVersion'` must succeed (the return value — `"V2"` — is not compared
   there; only `sy-subrc`). Failure shows *"Your PMT Office DLL needs to be
   updated"*.
2. **Property reading** — `FORM office_properties` calls function modules
   `Z_GETDOC_PROPERTIES` (read) and `Z_UPDATE_FILEPROPERTIES_FILE`
   (write-back, when *Insert/Change Properties* is ticked). These function
   modules wrap the COM DLL. Reading is per-property (the `ReadProperty`
   path), not via the XML round-trip.
3. **Validations** — `FORM check_scs_testrun` (+ helpers) validate the
   values read from the file against the transaction input and the
   customizing tables, then `check_scs_other_dokar` checks SCS uniqueness
   against DMS.

## Property → SAP field mapping (`FORM office_properties`)

Property names are upper-cased before matching, so the case in the document
does not matter.

| Document property | SAP work field | Used for |
|---|---|---|
| `Title` | `title` | DIS title |
| `Author` | `author` | |
| `Subject` | `subject` | |
| `Keywords` | `keywords` | |
| `AsBuilt` | `vaiasbuilt` | |
| `Mandator` | `vaidepartment` | |
| `vaiCompany` | `vaicompany` | |
| `Title5` | `vaidescription` / `title5` | description line |
| `DocumentNo` | `vaiscsnum` | **the SCS** — checked for uniqueness against `draw-zzscs` |
| `RevisionNo` | `vairevision` | revision |
| `Status` | `vaistatus` | status |
| `DocType` / `DocumentType` | `vaidoctype` / `zztype` | **document type** — must exist in table `zdoctype` |
| `SupplyGroup` | `zzsgcd` | must exist in `zsgccode` |
| `PlantCode` | `zzpla` | must exist in `zzscs_pp` for the project |
| `ProjectCode` | `file_proj` | compared against the project entered in the transaction |
| `CountNo` | `count` | must be exactly 3 digits |
| `Area`, `Equipment`, … | `zzarea`, `zzeqip`, … | further SCS segments, validated against customizing |

## The validations behind the common messages

| Message | Check (FORM) | Meaning / fix |
|---|---|---|
| `Project Code X is invalid` | `check_scs_testrun`: entered project must exist in `zzscs_procode` | choose a valid project |
| `Proj. Code <file> in Original is differnt to entered Proj. Code <entered>` | `check_scs_testrun`: file property `ProjectCode` ≠ transaction input. **Warning** normally, **error** on the PDB master. A blank `<file>` in the message means the document has no `ProjectCode` property at all. | fill `ProjectCode` in the document (or accept the warning) |
| `VAI Document type is missing` | `check_scs_testrun`: file property `DocType`/`DocumentType` is empty | fill `DocType` in the document |
| `Documenttype X is invalid` | value must exist in table `zdoctype` | use a valid type (e.g. from an existing document) |
| `Count Number ... is invalid` | `CountNo` must be exactly 3 digits | e.g. `001` |
| `The SCS <no> is already used in document number <doknr>` | `check_scs_other_dokar`: `draw` already contains a document with `zzscs` = the file's `DocumentNo` (same type/version/language constellation) | the SCS is consumed — usually a leftover from an earlier import test. Find it via CV04N (search by SCS), delete/archive it or use the next free number |

## Practical: preparing a test document

Fill the required properties with the toolkit itself before importing —
adapt names/values to the project:

```bat
cd "C:\Program Files\Siemens\MT_PropertyExchange"
MT_PropertyExchange.exe -mode=w -file="D:\test.docx" -prop=ProjectCode -val=T.FAB0
MT_PropertyExchange.exe -mode=w -file="D:\test.docx" -prop=DocType     -val=DOK
MT_PropertyExchange.exe -mode=w -file="D:\test.docx" -prop=DocumentNo  -val=T.FAB0.CB2.30/DOK001
MT_PropertyExchange.exe -mode=w -file="D:\test.docx" -prop=RevisionNo  -val=00
MT_PropertyExchange.exe -mode=w -file="D:\test.docx" -prop=Status      -val=Draft
:: verify what SAP will see:
MT_PropertyExchange.exe -mode=g -file="D:\test.docx"
```

Or import a complete set at once from a `[PROPERTIES]` text file /
`<DOCUMENT><PROPERTIES>` XML (see `docs/ARCHITECTURE.md` → data formats and
the samples in `MTPropertyExchangeInstall/test/`).
