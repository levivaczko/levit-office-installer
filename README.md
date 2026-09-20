# levIT Office LTSC installer

A single-file Windows installer for **Microsoft Office LTSC 2019 / 2021 / 2024
Professional Plus**. It wraps Microsoft's own Office Deployment Tool (ODT) in a
small console program that does the things a customer would otherwise have to
do by hand, and in the right order:

1. asks for administrator rights (relaunches itself elevated)
2. lets the user pick **which applications** to install (`[x]` / `[ ]`, arrow
   keys and space) and whether to install the **64-bit or 32-bit** edition
3. detects any Office already on the machine and lists it by name
4. asks for permission, then removes all of it, including preinstalled
   Microsoft 365 trials
5. installs the chosen Office LTSC edition from Microsoft's CDN

The installer is built for customers of [levit.hu](https://levit.hu), a
Hungarian software licence shop, so the on-screen text is Hungarian. The source
is published so anyone can see exactly what the executable does before running
it.

## Why the removal step exists

ODT's `<RemoveMSI />` only removes old MSI-based Office (2010/2013/2016 MSI). It
does **not** touch Click-to-Run installations: Microsoft 365, factory trials,
retail Home & Student. If those stay, Office LTSC installs next to them into the
same Click-to-Run container and the result is no longer a clean volume install.
A valid MAK key is then rejected with *"the product key is invalid"*.

Only `<Remove All="TRUE" />` fixes this, and it needs a separate ODT run. That is
why there are two configurations and two steps.

## How the configuration is built

`install_<edition>.xml` is a template. At run time the installer copies it,
sets `OfficeClientEdition` to `32` or `64` and adds one `<ExcludeApp ID="..." />`
per application the user unticked, then hands the result to `setup.exe
/configure`. Pressing Enter twice installs the full suite, 64-bit, which is
exactly what the template says.

The list of offered applications, and which edition contains each one (Office
LTSC 2024 no longer ships Publisher), is the `APPS` table at the top of
`office_installer.py`.

## Repository layout

| File | Purpose |
|---|---|
| `office_installer.py` | The installer. One source, three executables. |
| `build.py` | Builds all three with PyInstaller, optionally code-signs them. |
| `remove.xml` | Removes every Office on the machine, silently. |
| `install_2019.xml` | Template: ProPlus2019Volume, PerpetualVL2019, hu-hu |
| `install_2021.xml` | Template: ProPlus2021Volume + language pack, PerpetualVL2021, hu-hu |
| `install_2024.xml` | Template: ProPlus2024Volume, PerpetualVL2024, hu-hu |
| `diag.bat` | Support tool: dumps Office licence state to a text file, changes nothing. |
| `levit.ico` | Icon embedded in the executables. |

Not in the repository, by design: **`setup.exe`**, the Office Deployment Tool.
It is Microsoft's binary. Download it from the
[Microsoft Download Center](https://www.microsoft.com/en-us/download/details.aspx?id=49117),
run the self-extractor and copy only `setup.exe` next to `build.py`.

## Building

Windows only (PyInstaller cannot cross-compile).

Use a **32-bit Python**: the resulting executable then runs on both 32-bit and
64-bit Windows, and can install either Office edition from either.

```bat
py -3.12-32 -m pip install pyinstaller
py -3.12-32 build.py
```

Output in `dist/`:

```
Office_2019_Pro_Plus_LTSC.exe
Office_2021_Pro_Plus_LTSC.exe
Office_2024_Pro_Plus_LTSC.exe
```

`build.py` warns if it is started with a 64-bit Python.

## Code signing

If `SIGN_SUBJECT` is set (environment variable, or the constant at the top of
`build.py`) to the certificate's subject name, `build.py` signs all three
executables with the Windows SDK `signtool` using SHA-256 and a Certum RFC 3161
timestamp, then verifies them. A failed signature stops the build; nothing
unsigned leaves silently. With `SIGN_SUBJECT` empty the step is skipped.

```bat
set SIGN_SUBJECT=Open Source Developer, Vaczkó Levente
py -3.12-32 build.py
```

## Testing a build

On a VM with a factory Microsoft 365 trial:

1. run the executable, allow the removal
2. after the install, open Word and enter the MAK key
3. from an elevated command prompt:

```bat
cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /dstatus
```

Exactly one `LICENSE NAME` line should appear, containing `VL_MAK`, with
`LICENSE STATUS: ---LICENSED---`.

## Notes

- The installer downloads Office from Microsoft's CDN; an internet connection
  is required. The executable itself stays around 10 MB.
- A Microsoft 365 subscription installed on the machine is removed too. The
  installer says so and asks before touching anything.
- A licence key is not included and not needed to run the installer; it is
  entered in Office afterwards.
- Microsoft, Windows and Office are trademarks of Microsoft Corporation. This
  project is not affiliated with or endorsed by Microsoft.

## Licence

MIT, see [LICENSE](LICENSE).
