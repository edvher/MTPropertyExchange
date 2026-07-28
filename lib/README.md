# lib

This folder holds `Interop.Dsofile.dll`, the .NET interop assembly for the
native `dsofile.dll` COM component.

It is **generated automatically on first build** of
`MT_PropertyExchangeDLL.dll.vbproj` (see the `GenerateDsofileInterop` target in
that project) and is deliberately not checked in.

To generate it manually, run from a *Visual Studio Developer Command Prompt* in
the repository root:

```
tlbimp Dsofile\Lib\dsofile.tlb /keyfile:Dsofile\dsofile_key.snk /namespace:dsofile /out:lib\Interop.Dsofile.dll
```

Notes:

* The `/keyfile` is required: `MT_PropertyExchangeDLL` is strong-named, so all
  referenced assemblies must be strong-named too. (The copy of
  `Interop.dsofile.dll` inside the `Dsofile` folder is **unsigned** and cannot
  be referenced directly.)
* The `/namespace:dsofile` matters: the VB code does `Imports dsofile`.
* The interop assembly is pure MSIL (AnyCPU) and works for both the x86 and
  the x64 `dsofile.dll` — it only carries the COM type definitions.
