"""
Erstellt Reverse Charge and IGE.xlsx aus 3216.xlsx + 3226.xlsx
Kreditoren.xlsx wird bevorzugt aus Y:/HRV/ZZ_GK Tools/ geladen;
falls nicht erreichbar, kann sie manuell ausgewaehlt werden.
"""

import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

KREDITOREN_PATH = r"Y:\HRV\ZZ_GK Tools\Kreditoren.xlsx"

# Fix 1: Pfad relativ zum Skript, nicht absolut hardcoded
LOG_FILE = os.path.join(os.path.dirname(__file__), "build_log.txt")
_log = []

REQUIRED_COLS = {"Referenz", "Gegenkto", "Datum"}


def log(msg):
    _log.append(msg)
    print(msg.encode("ascii", "replace").decode())


def ask_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Ordner mit 3216.xlsx und 3226.xlsx auswaehlen")
    root.destroy()
    return folder


# Fix 5: eigene Hilfsfunktion fuer Datei-Dialog (fuer Kreditoren-Fallback)
def ask_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return path


def lookup_service_date(kred, ref, gegenkto):
    ref = str(ref).strip()
    gegenkto = str(gegenkto).strip().replace(".0", "")

    match = kred[(kred["Invoice Number"] == ref) & (kred["Company Account"] == gegenkto)]
    if len(match) == 1:
        return match.iloc[0]["Service Date"], "OK"

    ref_norm = ref.replace(" ", "")
    match = kred[(kred["Invoice Number normalized"] == ref_norm) & (kred["Company Account"] == gegenkto)]
    if len(match) == 1:
        found = match.iloc[0]["Invoice Number"]
        return match.iloc[0]["Service Date"], f"FUZZY ({ref} -> {found})"

    if len(match) == 0:
        return None, f"KEIN TREFFER: Ref={ref} Kred={gegenkto}"
    return None, f"MEHRDEUTIG: Ref={ref} Kred={gegenkto}"


def main():
    # Ordner auswaehlen
    folder = ask_folder()
    if not folder:
        print("Kein Ordner ausgewaehlt. Abbruch.")
        return

    log(f"Ordner: {folder}")

    # Pflichtdateien pruefen
    for fname in ["3216.xlsx", "3226.xlsx"]:
        if not os.path.exists(os.path.join(folder, fname)):
            messagebox.showerror("Fehler", f"{fname} nicht im Ordner gefunden.")
            return

    # Fix 5: Kreditoren laden — Fallback auf Datei-Dialog wenn Netzwerkpfad nicht erreichbar
    kred_path = KREDITOREN_PATH
    if not os.path.exists(kred_path):
        log(f"Kreditoren.xlsx nicht gefunden unter: {kred_path}")
        messagebox.showwarning(
            "Kreditoren.xlsx nicht gefunden",
            f"Die Datei wurde nicht gefunden:\n{kred_path}\n\nBitte im naechsten Schritt manuell auswaehlen."
        )
        kred_path = ask_file(
            title="Kreditoren.xlsx auswaehlen",
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if not kred_path:
            messagebox.showerror("Fehler", "Keine Kreditoren.xlsx ausgewaehlt. Abbruch.")
            return

    log("Lade Kreditoren.xlsx ...")
    kred = pd.read_excel(kred_path)
    kred["Invoice Number"] = kred["Invoice Number"].astype(str).str.strip()
    kred["Invoice Number normalized"] = kred["Invoice Number"].str.replace(" ", "", regex=False)
    kred["Company Account"] = kred["Company Account"].astype(str).str.strip().str.replace(".0", "", regex=False)

    sheets = {}
    warnings = []

    for source_file, sheet_name in [("3216.xlsx", "Reverse Charge - 3216"), ("3226.xlsx", "IGE - 3226")]:
        path = os.path.join(folder, source_file)
        df = pd.read_excel(path)
        df = df.dropna(subset=["Referenz", "Gegenkto"], how="all")

        # Fix 2: Spaltenvalidierung mit verstaendlicher Fehlermeldung
        missing = REQUIRED_COLS - set(df.columns)
        if missing:
            messagebox.showerror(
                "Spalten fehlen",
                f"{source_file}: Folgende Pflichtspalten fehlen:\n{', '.join(sorted(missing))}"
            )
            return

        service_dates = []
        for _, row in df.iterrows():
            sd, status = lookup_service_date(kred, row["Referenz"], row["Gegenkto"])
            if status != "OK":
                log(f"  [{source_file}] {status}")
                warnings.append(f"[{source_file}] {status}")
            if sd is not None and hasattr(sd, "strftime"):
                sd = sd.strftime("%d.%m.%Y")
            service_dates.append(sd)

        # Datum im deutschen Format
        df["Datum"] = pd.to_datetime(df["Datum"]).dt.strftime("%d.%m.%Y")

        df.insert(1, "Service Date", service_dates)

        # Fix 4: Unnamed-Spalten anhand ihrer tatsaechlichen Position umbenennen
        cols = list(df.columns)
        for i, col in enumerate(cols):
            if col.startswith("Unnamed:"):
                cols[i] = f"Unnamed: {i}"
        df.columns = cols

        sheets[sheet_name] = df
        log(f"{source_file}: {len(df)} Zeilen verarbeitet")

    output_path = os.path.join(folder, "Reverse Charge and IGE.xlsx")

    # Fix 7: Ueberschreib-Schutz
    if os.path.exists(output_path):
        if not messagebox.askyesno(
            "Datei ueberschreiben?",
            f"Die Datei existiert bereits:\n{output_path}\n\nUeberschreiben?"
        ):
            log("Abbruch durch Benutzer (kein Ueberschreiben).")
            return

    # Fix 3: PermissionError abfangen (z.B. Datei noch in Excel geoeffnet)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except PermissionError:
        messagebox.showerror(
            "Datei gesperrt",
            f"Datei konnte nicht gespeichert werden.\n\nBitte zuerst schliessen:\n{output_path}"
        )
        return

    log(f"\nGespeichert: {output_path}")

    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(_log))

    if warnings:
        msg = f"Datei gespeichert.\n\nHinweise ({len(warnings)}):\n" + "\n".join(warnings)
        messagebox.showwarning("Fertig mit Hinweisen", msg)
    else:
        messagebox.showinfo("Fertig", "Reverse Charge and IGE.xlsx wurde erfolgreich erstellt.")


if __name__ == "__main__":
    # Fix 6: Unerwartete Fehler abfangen und als Dialog anzeigen statt stilles Absturz
    try:
        main()
    except Exception as e:
        messagebox.showerror("Unerwarteter Fehler", str(e))
        sys.exit(1)
