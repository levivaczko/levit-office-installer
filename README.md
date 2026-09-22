# levIT Office LTSC installer

A single-file Windows installer for **Microsoft Office 2016 / 2019 / 2021 / 2024
Professional Plus** volume (MAK) licences. It wraps Microsoft's own Office Deployment Tool (ODT) in a
small console program that does the things a customer would otherwise have to
do by hand, and in the right order:

1. asks for administrator rights (relaunches itself elevated)
2. lets the user pick **which applications** to install (`[x]` / `[ ]`, arrow
   keys and space) and whether to install the **64-bit or 32-bit** edition
3. detects any Office already on the machine and lists it by name
4. asks for permission, then removes all of it, including preinstalled
   Microsoft 365 trials
5. lists any product keys and KMS server setting left on the machine and asks
   separately whether to remove those too
6. installs the chosen Office edition from Microsoft's CDN
7. Office 2016 only: prepares the volume licence and offers to enter and
   activate the product key right away

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

## Why it asks about the product key separately

An ODT removal deletes the files, **not the licence**: the product key stays in
the Windows Software Protection Platform store and only `ospp.vbs /unpkey`
removes it. The visible consequence is that an earlier Office licence makes a
freshly installed Office report itself as activated, under the old name: the
Account pane shows `Activated product: Microsoft Office Professional Plus 2016`
on a 2024 LTSC installation. The 2016/2019/2021/2024 volume licences all belong
to the same "Office 16" family, which is why they are accepted across
installations. The customer is then never asked for the key they bought.

The same step surfaces a **KMS server override**
(`KMS machine registry override defined:` in `/dstatus`), which activator
scripts write into the registry. While it is there, Office re-activates against
that server every 180 days instead of using the purchased MAK key, and the
Account pane shows a `..._KMS_Client edition` licence. On a yes the installer
clears it with `/remhst` and `/cachst:FALSE`.

The installer reads all of this **before** the removal, while `ospp.vbs` still
exists, prints what it found, and only acts on an explicit yes: a retail key
must not be dropped without the owner knowing, because reactivating it needs
their key card.

## How the configuration is built

`install_<edition>.xml` is a template. At run time the installer copies it,
sets `OfficeClientEdition` to `32` or `64` and adds one `<ExcludeApp ID="..." />`
per application the user unticked, then hands the result to `setup.exe
/configure`. Pressing Enter twice installs the full suite, 64-bit, which is
exactly what the template says.

The list of offered applications, and which edition each one is offered for,
is the `APPS` table at the top of `office_installer.py`.

## Office 2016 is different

The Office Deployment Tool has no volume product ID for Office 2016 (the 2016
volume edition was still MSI-based and is only available from the Volume
Licensing portal). The 2016 installer therefore installs the **Retail**
Click-to-Run edition (`ProPlusRetail`) and then uses Office's own `ospp.vbs`
to install the volume licence files that ship inside every Click-to-Run
installation (`root\Licenses16\client-issuance-*`, `pkeyconfig-office`,
`ProPlusVL_MAK*`). After that Office accepts a MAK key. The installer asks
for the key at the end; leaving it empty skips activation, and the key can
be entered in Word instead.

## Repository layout

| File | Purpose |
|---|---|
| `office_installer.py` | The installer. One source, four executables. |
| `build.py` | Builds all four with PyInstaller, optionally code-signs them. |
| `remove.xml` | Removes every Office on the machine, silently. |
| `install_2016.xml` | Template: ProPlusRetail, Current channel, hu-hu (see above) |
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
Office_2016_Pro_Plus_LTSC.exe
Office_2019_Pro_Plus_LTSC.exe
Office_2021_Pro_Plus_LTSC.exe
Office_2024_Pro_Plus_LTSC.exe
```

`build.py` warns if it is started with a 64-bit Python.

## Code signing

If `SIGN_SUBJECT` is set (environment variable, or the constant at the top of
`build.py`) to the certificate's subject name, `build.py` signs all four
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
