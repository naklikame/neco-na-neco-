# Nastroje pro soubory

Jednoducha aplikace pro Windows na praci se soubory: prevody, prejmenovani, komprese obrazku a prace s PDF.

## Jak stahnout a spustit

1. Nahore klikni na zelene tlacitko **Code**, pak na **Download ZIP**.
2. ZIP rozbal (prave tlacitko, **Extrahovat vse**). Nespoustej to primo ze ZIPu.
3. Dvakrat klikni na **`SPUSTIT_KONVERTOR.bat`**.
4. Kdyz Windows ukaze modre okno „Systém Windows ochránil váš počítač“, klikni na **Další informace** a pak **Přesto spustit**.

Vsechny soubory musi zustat spolu v jedne slozce.

## Co aplikace umi

| Zalozka | Co dela | Co je potreba |
|---|---|---|
| **Prevod** | Word (.docx) na PDF | **Microsoft Word** |
| | Obrazek (JPG/PNG) na PDF | **Microsoft Word** |
| | Obrazek na WebP | nic navic (`cwebp.exe` je pribaleny) |
| **Prejmenovani** | Hromadne prejmenovani souboru (novy nazev, najit/nahradit, prefix, suffix, cislovani) | nic navic |
| **Komprese** | Zmenseni JPG/PNG/WebP/HEIC, zmena rozliseni, ulozeni jako JPG/PNG/WebP | nic navic, jen pro **HEIC** (fotky z iPhonu) viz nize |
| **PDF** | Spojeni vice PDF do jednoho, vyber stranek z PDF | **Microsoft Word** |

## Pozadavky

- **Windows 10 nebo 11.** PowerShell je soucasti Windows, nic se neinstaluje.
- **Microsoft Word** (desktopova verze, Office 2013 a novejsi / Microsoft 365) pro vsechno, co vytvari PDF.
  Webovy Word ani LibreOffice nestaci.
- **HEIC fotky z iPhonu:** z Microsoft Store nainstaluj **HEIF Image Extensions** a **HEVC Video Extensions**.

## Poznamky

- U PDF (spojeni / vyber stranek) Word obsah prevadi do editovatelne podoby, takze slozite layouty nemusi vyjit uplne presne.
- `cwebp.exe` a `dwebp.exe` jsou oficialni nastroje od Googlu (libwebp), licence je v `LICENSE-libwebp.txt`.
