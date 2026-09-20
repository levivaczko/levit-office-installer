"""Lefordítja mind a három levIT Office telepítőt egy-egy önálló exe-be.

Windowson kell futtatni, mert a PyInstaller nem tud más rendszerre fordítani.

Előfeltételek ugyanebben a mappában:
  setup.exe          a Microsoft Office Deployment Toolból kicsomagolva
  remove.xml
  install_2019.xml / install_2021.xml / install_2024.xml
  office_installer.py

Futtatás:
  pip install pyinstaller
  python build.py

Eredmény a dist/ mappában:
  Office_2019_Pro_Plus_LTSC.exe
  Office_2021_Pro_Plus_LTSC.exe
  Office_2024_Pro_Plus_LTSC.exe

Ezek mennek a webshop uploads mappájába, a régiek helyére. A neveket ne
változtasd: a Letöltések oldal és mind a hat aktiválási útmutató ezekre mutat.

Kódaláírás:
  Ha a SIGN_SUBJECT alább (vagy környezeti változóként) ki van töltve a
  tanúsítvány nevével, a kész exe-ket a Windows SDK signtool-ja aláírja,
  Certum időbélyeggel. Ehhez a SimplySign Desktopnak futnia kell, az adja a
  virtuális kártyát. Üres SIGN_SUBJECT esetén a lépés kimarad.
"""

import glob
import os
import re
import shutil
import struct
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
EDITIONS = ["2019", "2021", "2024"]

# A tanúsítvány "Kiadva" (Subject CN) neve, pontosan úgy, ahogy a Certum
# kiállította. Környezeti változóval felülírható.
SIGN_SUBJECT = os.environ.get("SIGN_SUBJECT", "")
SIGN_TIMESTAMP = os.environ.get("SIGN_TIMESTAMP", "http://time.certum.pl")


def need(path):
    if not os.path.exists(os.path.join(HERE, path)):
        sys.exit(
            f"Hianyzik: {path}\n"
            "A setup.exe a Microsoft Office Deployment Toolbol jon: toltsd le a\n"
            "Microsoft Download Centerbol, futtasd, es csomagold ki ide."
        )


def find_signtool():
    """A signtool.exe a PATH-on, vagy a Windows SDK-ból a legújabb verzió."""
    found = shutil.which("signtool")
    if found:
        return found
    pattern = r"C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe"
    hits = sorted(glob.glob(pattern))
    return hits[-1] if hits else None


def sign(paths):
    """Aláírja a kész exe-ket. Sikertelen aláírásnál hibával leáll, hogy ne
    kerüljön aláíratlan fájl a webshopba úgy, hogy azt hisszük, aláírt."""
    if not SIGN_SUBJECT:
        print("\nAláírás kihagyva: nincs SIGN_SUBJECT megadva, az exe-k aláíratlanok.")
        return
    tool = find_signtool()
    if not tool:
        sys.exit(
            "Nincs signtool.exe. Telepítsd a Windows SDK-t (elég a\n"
            "\"Windows SDK Signing Tools for Desktop Apps\" összetevő)."
        )
    for path in paths:
        print(f"\nAláírás: {os.path.basename(path)}")
        rc = subprocess.run([
            tool, "sign", "/fd", "SHA256", "/tr", SIGN_TIMESTAMP, "/td", "SHA256",
            "/n", SIGN_SUBJECT, path,
        ]).returncode
        if rc != 0:
            sys.exit("Az aláírás nem sikerült. Fut a SimplySign Desktop, és jó a SIGN_SUBJECT?")
        rc = subprocess.run([tool, "verify", "/pa", path], capture_output=True).returncode
        if rc != 0:
            sys.exit(f"Az aláírás ellenőrzése elbukott: {os.path.basename(path)}")


def main():
    need("setup.exe")
    need("remove.xml")
    need("office_installer.py")
    for ed in EDITIONS:
        need(f"install_{ed}.xml")

    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        sys.exit("Nincs PyInstaller. Telepitsd:  pip install pyinstaller")

    # Az exe olyan bites lesz, amilyen ez a Python. A 32 bites exe 64 bites
    # Windowson is fut, forditva nem, ezert 32 bites Pythonnal erdemes forditani.
    if struct.calcsize("P") * 8 == 64:
        print("FIGYELEM: 64 bites Python, az exe 32 bites Windowson nem fog elindulni.")
        print("          32 bites Pythonnal (py -3-32 build.py) mindketton fut.\n")

    src = open(os.path.join(HERE, "office_installer.py"), encoding="utf-8").read()
    os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)

    for ed in EDITIONS:
        print(f"\n=== Office {ed} ===")
        # Az EDITION konstans atirasa   igy egy forrasbol lesz harom exe.
        patched = re.sub(r'^EDITION = "\d{4}"', f'EDITION = "{ed}"', src, count=1,
                         flags=re.M)
        if f'EDITION = "{ed}"' not in patched:
            sys.exit("Nem talaltam az EDITION sort az office_installer.py-ban.")
        tmp = os.path.join(HERE, f"_build_{ed}.py")
        open(tmp, "w", encoding="utf-8").write(patched)

        # PyInstaller resolves a relative --add-data source against the SPEC
        # file's directory, not the cwd. The spec lives in _work/, so relative
        # names would be looked up there and fail. Always pass absolute paths.
        sep = ";" if os.name == "nt" else ":"
        def data(name):
            return f"{os.path.join(HERE, name)}{sep}."

        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile", "--console", "--clean", "--noconfirm",
            "--name", f"Office_{ed}_Pro_Plus_LTSC",
            "--distpath", os.path.join(HERE, "dist"),
            "--workpath", os.path.join(HERE, "_work"),
            "--specpath", os.path.join(HERE, "_work"),
            "--add-data", data("setup.exe"),
            "--add-data", data("remove.xml"),
            "--add-data", data(f"install_{ed}.xml"),
        ]
        icon = os.path.join(HERE, "levit.ico")
        if os.path.exists(icon):
            cmd += ["--icon", icon]
        cmd.append(tmp)   # tmp is already absolute (built from HERE)

        rc = subprocess.run(cmd, cwd=HERE).returncode
        os.remove(tmp)
        if rc != 0:
            sys.exit(f"A(z) {ed} forditasa hibaval allt le.")

    shutil.rmtree(os.path.join(HERE, "_work"), ignore_errors=True)
    sign([os.path.join(HERE, "dist", f"Office_{ed}_Pro_Plus_LTSC.exe") for ed in EDITIONS])
    print("\nKesz. A kesz fajlok a dist/ mappaban vannak.")
    print("Ezeket masold a webshop uploads mappajaba, a regiek helyere.")


if __name__ == "__main__":
    main()
