"""levIT Office LTSC telepítő.

Egyetlen exe-be fordítva (PyInstaller, onefile). Magában hordozza a Microsoft
Office Deployment Tool setup.exe-jét és a konfigurációkat, így a vevőnek csak
egy fájlt kell elindítania.

Menete:
  1. rendszergazda jog megszerzése (magától újraindul emelt joggal)
  2. a vevő kijelöli, mely alkalmazások kerüljenek fel (nyilak + szóköz),
     és hogy 32 vagy 64 bites Office legyen
  3. megnézi, van-e Office a gépen
  4. ha van, kiírja mit talált, és rákérdez, eltávolíthatja-e
  5. eltávolít mindent (Microsoft 365 próbaverziót is)
  6. felteszi a megvásárolt LTSC változatot a kijelölés szerint

Ha nincs mit eltávolítani, a 4-5. lépés kimarad. Enter-Enter a 2. lépésben
a teljes csomagot adja, 64 bitesen, ahogy a bundled XML írja.

Fordítás: lásd README.md. Az EDITION értéket a build script írja át.
"""

import ctypes
import os
import platform
import subprocess
import sys
import tempfile
import time

# A build script cseréli 2019 / 2021 / 2024 értékre.
EDITION = "2024"

APP_TITLE = f"levIT  -  Office {EDITION} Professional Plus LTSC"
SUPPORT = "support@levit.hu"
SHOP = "levit.hu"

# ANSI colours. os.system("") below turns on VT processing on Windows 10+;
# on anything older the codes would show as garbage, so they are dropped there.
_C = {
    'red':   '\033[38;5;167m',   # the levIT brand red
    'white': '\033[97m',
    'dim':   '\033[38;5;245m',
    'green': '\033[38;5;71m',
    'amber': '\033[38;5;179m',
    'off':   '\033[0m',
}


def c(name, text):
    return f"{_C.get(name, '')}{text}{_C['off']}" if _C else str(text)


W = 58   # a doboz belső szélessége


def _row(*parts):
    """Egy dobozsor. A parts (szín, szöveg) párokból áll; a kitöltést a LÁTHATÓ
    hosszból számoljuk, mert a színkódok nem foglalnak helyet a képernyőn."""
    visible = "".join(t for _, t in parts)
    body = "".join(c(col, t) for col, t in parts)
    pad = " " * max(0, W - len(visible))
    return f"  {c('dim', '│')}{body}{pad}{c('dim', '│')}"


def banner():
    bar = "─" * W
    print()
    print(f"  {c('dim', '┌' + bar + '┐')}")
    print(_row(('', '  ')))
    print(_row(('', '   '), ('white', 'lev'), ('red', 'IT'),
               ('', '   '), ('dim', '·'), ('', '   '),
               ('white', f'Office {EDITION} Professional Plus LTSC')))
    print(_row(('', '  ')))
    print(_row(('', '   '), ('dim', f'{SHOP}   ·   {SUPPORT}')))
    print(_row(('', '  ')))
    print(f"  {c('dim', '└' + bar + '┘')}")
    print()


# A Click-to-Run kulcs a megbízható jelzés arra, hogy van Office a gépen.
# A 32 bites ág is kell: 64 bites Windowson 32 bites Office ide kerül.
C2R_KEYS = [
    r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
    r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration",
]

# Emberi nevek a registryben talált termékazonosítókhoz. Ami nincs a listán,
# azt nyersen írjuk ki   jobb egy csúnya azonosító, mint egy néma lista.
PRODUCT_NAMES = {
    "O365HomePremRetail": "Microsoft 365 (otthoni)",
    "O365ProPlusRetail": "Microsoft 365 Apps for enterprise",
    "O365BusinessRetail": "Microsoft 365 Apps for business",
    "O365SmallBusPremRetail": "Microsoft 365 Business Premium",
    "HomeStudentRetail": "Office Home & Student",
    "HomeBusinessRetail": "Office Home & Business",
    "ProfessionalRetail": "Office Professional",
    "ProPlusRetail": "Office Professional Plus (retail)",
    "ProPlus2019Volume": "Office 2019 Professional Plus LTSC",
    "ProPlus2021Volume": "Office 2021 Professional Plus LTSC",
    "ProPlus2024Volume": "Office 2024 Professional Plus LTSC",
    "Standard2019Volume": "Office 2019 Standard",
    "Standard2021Volume": "Office 2021 Standard",
    "Standard2024Volume": "Office 2024 Standard",
    "VisioProRetail": "Visio Professional",
    "ProjectProRetail": "Project Professional",
}


# Az ODT ExcludeApp azonosítói és a vevőnek mutatott nevük. A harmadik elem
# mondja meg, mely kiadásokban van benne az alkalmazás: az LTSC 2024-ben már
# nincs Publisher, ott nem is kínáljuk fel. Alapból minden be van jelölve,
# így Enter-Enter ugyanazt adja, mint a régi telepítő.
APPS = [
    ("Word",       "Word",                    ("2019", "2021", "2024")),
    ("Excel",      "Excel",                   ("2019", "2021", "2024")),
    ("PowerPoint", "PowerPoint",              ("2019", "2021", "2024")),
    ("Outlook",    "Outlook",                 ("2019", "2021", "2024")),
    ("OneNote",    "OneNote",                 ("2019", "2021", "2024")),
    ("Access",     "Access",                  ("2019", "2021", "2024")),
    ("Publisher",  "Publisher",               ("2019", "2021")),
    ("Lync",       "Skype Vállalati verzió",  ("2019", "2021", "2024")),
    ("OneDrive",   "OneDrive",                ("2019", "2021", "2024")),
]


def _read_key():
    """Egy billentyű a konzolról. A nyilak két kódként jönnek (0x00 vagy 0xE0,
    majd a betű), ezeket névvé alakítjuk; minden más kisbetűsen jön vissza."""
    import msvcrt
    ch = msvcrt.getwch()
    if ch in ("\x00", "\xe0"):
        return {"H": "up", "P": "down", "K": "left", "M": "right"}.get(msvcrt.getwch(), "")
    if ch == "\r":
        return "enter"
    if ch == " ":
        return "space"
    if ch == "\x1b":
        return "esc"
    return ch.lower()


def choose(title, hint, labels, checked, single=False):
    """Konzolos jelölőlista. ↑↓ mozgat, szóköz jelöl, Enter továbblép. A
    single=True egyválasztós (rádió) lista, ott az Enter a kijelöltet választja.
    A blokkot helyben rajzoljuk újra ANSI kurzormozgatással, így nem görget.
    Ha nincs igazi konzol (pl. átirányított bemenet), az alapértékek mennek."""
    checked = list(checked)
    if not sys.stdin.isatty() or os.name != "nt":
        return checked
    n = len(labels)
    cur = 0
    total = n + 5  # cím, üres, sorok, üres, tipp, hibasor

    def draw(first, err=""):
        if not first:
            sys.stdout.write(f"\033[{total}A")
        out = [f"  {c('white', title)}", ""]
        for i, label in enumerate(labels):
            if single:
                mark = "(x)" if checked[i] else "( )"
            else:
                mark = "[x]" if checked[i] else "[ ]"
            mark = c("green", mark) if checked[i] else c("dim", mark)
            pointer = c("red", "›") if i == cur else " "
            out.append(f"   {pointer} {mark} {c('white', label) if i == cur else label}")
        out.append("")
        out.append(f"  {c('dim', hint)}")
        out.append(f"  {c('amber', err)}" if err else "")
        for line in out:
            sys.stdout.write("\033[2K" + line + "\n")
        sys.stdout.flush()

    sys.stdout.write("\033[?25l")   # kurzor el, hogy ne villogjon a listában
    try:
        draw(True)
        while True:
            key = _read_key()
            err = ""
            if key == "up":
                cur = (cur - 1) % n
            elif key == "down":
                cur = (cur + 1) % n
            elif single and key in ("space", "left", "right", "enter"):
                checked = [i == cur for i in range(n)]
                if key == "enter":
                    draw(False)
                    break
            elif key == "space":
                checked[cur] = not checked[cur]
            elif key == "enter":
                if any(checked):
                    break
                err = "Legalább egy alkalmazást jelölj be."
            draw(False, err)
    finally:
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()
    return checked


def is_64bit_windows():
    # 32 bites exe 64 bites Windowson: a machine() az igazi gépet mondja,
    # az ARCHITEW6432 pedig csak WOW64 alatt létezik.
    return platform.machine().endswith("64") or "PROCESSOR_ARCHITEW6432" in os.environ


def ask_options():
    """Alkalmazások és bitesség. Visszaad: (kihagyott ExcludeApp ID-k, 32|64)."""
    apps = [a for a in APPS if EDITION in a[2]]
    picked = choose(
        f"Mi kerüljön fel az Office {EDITION}-ből?",
        "↑↓ mozgás   ·   Szóköz jelöl / kihagy   ·   Enter tovább",
        [name for _, name, _ in apps],
        [True] * len(apps),
    )
    excluded = [app_id for (app_id, _, _), on in zip(apps, picked) if not on]

    if is_64bit_windows():
        sel = choose(
            "Melyik változat?",
            "↑↓ mozgás   ·   Enter választ",
            ["64 bites   (ajánlott)", "32 bites"],
            [True, False],
            single=True,
        )
        bits = "64" if sel[0] else "32"
    else:
        bits = "32"
        print(f"  {c('dim', 'Ez a Windows 32 bites, ezért a 32 bites Office kerül fel.')}")
        print()

    kept = [name for (_, name, _), on in zip(apps, picked) if on]
    print(f"  {c('dim', 'Telepítendő:')}  {', '.join(kept)}")
    if excluded:
        skipped = [name for (_, name, _), on in zip(apps, picked) if not on]
        print(f"  {c('dim', 'Kihagyva:')}     {', '.join(skipped)}")
    print(f"  {c('dim', 'Változat:')}     {bits} bites")
    print()
    return excluded, bits


def build_install_xml(excluded, bits):
    """A csomagolt install_<kiadás>.xml-ből készít egy példányt a vevő
    választásával: OfficeClientEdition, és ExcludeApp sorok a fő termék alá
    (a LanguagePack nem kap ilyet, annak nincsenek alkalmazásai). A kész fájl
    a temp mappába megy; az elérési útját adja vissza."""
    import xml.etree.ElementTree as ET
    src = os.path.join(resource_dir(), f"install_{EDITION}.xml")
    tree = ET.parse(src)
    add = tree.getroot().find("Add")
    add.set("OfficeClientEdition", bits)
    main_product = next((p for p in add.findall("Product") if p.get("ID") != "LanguagePack"), None)
    if main_product is not None:
        for app_id in excluded:
            ET.SubElement(main_product, "ExcludeApp", {"ID": app_id})
    try:
        ET.indent(tree, space="  ")
    except AttributeError:   # Python < 3.9
        pass
    out = os.path.join(tempfile.gettempdir(), f"levit_office_{EDITION}.xml")
    tree.write(out, encoding="utf-8")
    return out


def resource_dir():
    """A PyInstaller onefile futáskor egy ideiglenes mappába csomagol ki; a
    setup.exe és az XML-ek ott lesznek. Fejlesztéskor a szkript mappája."""
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def relaunch_as_admin():
    """Újraindítja magát emelt joggal. A setup.exe rendszergazda nélkül csendben
    elbukik, ezért ezt nem bízzuk a vevőre."""
    try:
        params = " ".join(f'"{a}"' for a in sys.argv[1:])
        rc = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, params, None, 1
        )
        return rc > 32
    except Exception:
        return False


def installed_office():
    """A gépen lévő Office termékek listája, emberi néven."""
    import winreg

    found = []
    for path in C2R_KEYS:
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | view
                ) as k:
                    ids, _ = winreg.QueryValueEx(k, "ProductReleaseIds")
            except OSError:
                continue
            for pid in str(ids).split(","):
                pid = pid.strip()
                if not pid:
                    continue
                name = PRODUCT_NAMES.get(pid, pid)
                if name not in found:
                    found.append(name)
    return found


def run_setup(xml_name, what):
    """Egy ODT futtatás. A setup.exe a Microsoft CDN-jéről tölt, ezért tarthat
    percekig; a visszatérési kód dönt, nem a futásidő."""
    base = resource_dir()
    setup = os.path.join(base, "setup.exe")
    xml = xml_name if os.path.isabs(xml_name) else os.path.join(base, xml_name)
    if not os.path.exists(setup) or not os.path.exists(xml):
        print(f"  Hiányzó fájl a csomagban: {os.path.basename(setup if not os.path.exists(setup) else xml)}")
        return False
    try:
        proc = subprocess.run([setup, "/configure", xml], cwd=base)
        return proc.returncode == 0
    except Exception as exc:
        print(f"  Nem sikerült elindítani a telepítőt ({what}): {exc}")
        return False


def bail(msg):
    print()
    print(f"  {msg}")
    print(f"  Ha elakadtál, írj nekünk: {SUPPORT}")
    print()
    input("  Nyomj Entert a bezáráshoz...")
    sys.exit(1)


def main():
    os.system("")   # Windows 10+: bekapcsolja az ANSI színek feldolgozását
    try:
        ctypes.windll.kernel32.SetConsoleTitleW(APP_TITLE)
    except Exception:
        pass

    if not is_admin():
        banner()
        print(f"  {c('dim', 'A telepítéshez rendszergazdai jog kell, ezért újraindulok.')}")
        print(f"  {c('dim', 'Kérlek engedélyezd a felugró ablakban.')}")
        print()
        time.sleep(2)
        if relaunch_as_admin():
            sys.exit(0)
        bail("Nem sikerült megszerezni a rendszergazdai jogot. Zárd be ezt az "
             "ablakot, majd jobb gombbal: „Futtatás rendszergazdaként”.")

    banner()
    excluded, bits = ask_options()
    install_xml = build_install_xml(excluded, bits)
    existing = installed_office()

    if existing:
        print(f"  {c('amber', 'A gépen már van Office.')} Ezt el kell távolítani, különben")
        print("  az új licenckulcs nem fog működni.")
        print()
        for name in existing:
            print(f"     {c('dim', '•')} {name}")
        print()
        print(f"  {c('amber', 'Figyelem:')} ha van előfizetésed Microsoft 365-re, azt is")
        print("  eltávolítja. Előfizetés esetén később újra tudod telepíteni")
        print("  a Microsoft fiókodból.")
        print()
        answer = input(f"  Eltávolítsam, és feltegyem az Office {EDITION}-t? (I/N): ")
        if answer.strip().lower() not in ("i", "igen", "y", "yes"):
            print()
            print(f"  {c('dim', 'Megszakítva. Nem történt változás a gépen.')}")
            print()
            input("  Nyomj Entert a bezáráshoz...")
            sys.exit(0)

        print()
        print("  Eltávolítás folyamatban, ez pár percig tarthat.")
        print(f"  {c('dim', 'Kérlek ne zárd be ezt az ablakot...')}")
        print()
        if not run_setup("remove.xml", "eltávolítás"):
            bail("Az eltávolítás hibával állt le. Indítsd újra a gépet, "
                 "majd próbáld meg megint.")
        print(f"  {c('green', '✓')} Eltávolítás kész.")
        print()
    else:
        print(f"  {c('green', '✓')} Nincs korábbi Office a gépen, indul a telepítés.")
        print()

    print(f"  Indul az Office {EDITION} telepítése.")
    print(f"  {c('dim', 'A telepítő saját ablakot nyit, kövesd ott a folyamatot...')}")
    print()
    if not run_setup(install_xml, "telepítés"):
        bail("A telepítés hibával állt le.")

    bar = "─" * W
    print()
    print(f"  {c('green', '┌' + bar + '┐')}")
    print(f"  {c('green', '│')}{c('green', '   ✓  Kész! Az Office ' + EDITION + ' feltelepült.')}"
          + " " * max(0, W - len(f"   ✓  Kész! Az Office {EDITION} feltelepült.")) + f"{c('green', '│')}")
    print(f"  {c('green', '└' + bar + '┘')}")
    print()
    print("  Következő lépés: nyisd meg a Wordöt, és add meg a levIT-től")
    print("  e-mailben kapott 25 karakteres licenckulcsot.")
    print()
    print(f"  {c('dim', 'Köszönjük, hogy a ')}{c('white', 'lev')}{c('red', 'IT')}"
          f"{c('dim', ' ügyfele vagy!')}")
    print(f"  {c('dim', SHOP + '   ·   ' + SUPPORT)}")
    print()
    input("  Nyomj Entert a bezáráshoz...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
    except Exception as exc:  # a vevo ne csak egy eltuno ablakot lasson
        print()
        print(f"  Váratlan hiba: {exc}")
        print(f"  Kérlek küldd el ezt nekünk: {SUPPORT}")
        print()
        input("  Nyomj Entert a bezáráshoz...")
        sys.exit(1)
