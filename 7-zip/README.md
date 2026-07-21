# 7-zip — single-file installer tooling

`..\make-installer.bat` builds one self-contained
`MTPropertyExchangeInstall-<timestamp>.exe` out of the staged installer
folder, exactly like the historical build did
(`SFX module + config + 7z archive` concatenated).

Double-clicking the resulting EXE extracts the payload to a temporary folder
and starts `install.bat`, which installs, registers (32- **and** 64-bit),
runs the smoke test, writes a full log and shows a green (success) or red
(failure) console that closes on ENTER.

## Required tools (one-time setup on the build machine)

Two files are needed and are **not** checked in (third-party binaries):

1. **`7z.exe`** — install [7-Zip](https://www.7-zip.org/); the script finds it
   automatically in `%ProgramFiles%\7-Zip`. Alternatively copy `7z.exe` and
   `7z.dll` into this folder.
2. **`7zSD.sfx`** — the installer SFX module. Download the **“7-Zip Extra”**
   package from https://www.7-zip.org/download.html (file `7z<ver>-extra.7z`),
   extract `7zSD.sfx` from it and place it **in this folder**.

`MTPE_sfx_config.txt` is the SFX configuration (title, prompt, and the
command that runs after extraction). 7-Zip is free software (LGPL); the SFX
modules may be redistributed with your installer.
