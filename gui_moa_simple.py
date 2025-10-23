
import sys
import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    from moa_core import process_csv_to_moa_df, export_moa_excel
except ImportError:
    try:
        from moa_core_fallback import process_csv_to_moa_df, export_moa_excel  # type: ignore
    except ImportError as e:
        raise ImportError("Impossible d'importer moa_core ou moa_core_fallback. Place ce fichier dans le même dossier que moa_core.py.") from e

APP_TITLE = "MOA Extractor"

def convert(csv_path, save_path=None):
    df = process_csv_to_moa_df(csv_path)
    if not save_path:
        root_noext, _ = os.path.splitext(csv_path)
        save_path = root_noext + ".moa.xlsx"
    export_moa_excel(df, save_path)
    return save_path

def run_interactive():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("460x180")

    label_text = APP_TITLE + "\nCSV -> Excel (MOA)"
    label = tk.Label(root, text=label_text, font=("Segoe UI", 12))
    label.pack(pady=10)

    def on_click():
        csv_path = filedialog.askopenfilename(
            title="Choisir le CSV",
            filetypes=[("CSV files", "*.csv")]
        )
        if not csv_path:
            return
        try:
            save_path = filedialog.asksaveasfilename(
                title="Enregistrer l'Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")]
            )
            if not save_path:
                return
            out = convert(csv_path, save_path)
            messagebox.showinfo("Terminé", "Export réussi :\n" + out)
        except Exception as e:
            messagebox.showerror("Erreur", "Une erreur est survenue :\n{0}\n\n{1}".format(e, traceback.format_exc()))

    btn = tk.Button(root, text="Choisir un CSV et exporter l'Excel", command=on_click, width=42)
    btn.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    # If a CSV path is provided as argument -> auto convert without full GUI
    # Enables drag-and-drop of a CSV onto the EXE in Windows
    if len(sys.argv) > 1:
        csv_path = sys.argv[1]
        try:
            out = convert(csv_path)
            print(out)
        except Exception as e:
            # use a tiny Tk root to show a dialog
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erreur", "Echec conversion :\n{0}\n\n{1}".format(e, traceback.format_exc()))
            root.destroy()
    else:
        run_interactive()
