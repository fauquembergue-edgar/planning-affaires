import json
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, colorchooser
from pathlib import Path
import xlwings as xw
import uuid
from PIL import Image, ImageTk  # <--- Pour le logo
from datetime import datetime

EXCEL_FILE = "planning affaires.xlsm"
SHEET_MAIN = "PLANNING AFFAIRES"
OPTIONS_FILE = "options.json"

FIELDS = [
    "Type", "Secteur", "MOA", "MOE",
    "Commune", "Désignation de l'Affaire", "Date", "Heure réponse", "Montant estimé",
    "Principales Quantités", "Var.", "Métrés", "MT", "Responsable MT"
]

COLOR_FIELDS_BG = {"Type", "Chargé d'étude", "Métrés", "Coef.", "Responsable MT"}
COLOR_FIELDS_TEXT = {"Secteur"}
MANUAL_FIELDS = {"Principales Quantités"}

# ---- Palette TPW ----
TPW_COLORS = {
    "bleu": "#003A70",
    "bleu_clair": "#66A9D2",
    "jaune": "#FFD600",
    "gris_clair": "#F5F6F7",
    "blanc": "#FFFFFF",
    "rouge": "#E30613"
}

def load_options():
    if Path(OPTIONS_FILE).exists():
        with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {field: [] for field in FIELDS if field not in MANUAL_FIELDS}

def save_options(options):
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(options, f, indent=2, ensure_ascii=False)

def get_option_color(field, value):
    if field in options:
        for item in options[field]:
            if isinstance(item, dict) and item.get("value") == value:
                return item.get("color")
    return None

def add_option(field):
    new_value = simpledialog.askstring("Nouvelle option", f"Ajouter une valeur à '{field}' :")
    if new_value:
        color = "#FFFFFF"
        if field in COLOR_FIELDS_BG:
            color_tuple = colorchooser.askcolor(title=f"Couleur de fond pour '{new_value}'")
            if color_tuple[1]:
                color = color_tuple[1]
        entry = {"value": new_value}
        if field in COLOR_FIELDS_BG:
            entry["color"] = color
        options[field].append(entry)
        save_options(options)
        values = [item["value"] if isinstance(item, dict) else item for item in options[field]]
        entries[field]['values'] = values
        entries[field].set(new_value)

def remove_option(field):
    values = [item["value"] if isinstance(item, dict) else item for item in options.get(field, [])]
    if not values:
        messagebox.showinfo("Aucune valeur", f"Aucune valeur à supprimer pour '{field}'.")
        return
    
    dialog = tk.Toplevel(root)
    dialog.title(f"Supprimer une valeur pour {field}")
    dialog.geometry("350x100")
    tk.Label(dialog, text=f"Choisissez une valeur à supprimer dans '{field}' :", bg=TPW_COLORS["gris_clair"]).pack(pady=5)

    selected_var = tk.StringVar()
    combo = ttk.Combobox(dialog, textvariable=selected_var, values=values, state="readonly")
    combo.pack(pady=5)

    def confirm_delete():
        selected = selected_var.get()
        if selected:
            options[field] = [
                item for item in options[field]
                if not (
                    (isinstance(item, str) and item == selected)
                    or (isinstance(item, dict) and item.get("value") == selected)
                )
            ]
            save_options(options)
            new_values = [item["value"] if isinstance(item, dict) else item for item in options[field]]
            entries[field]['values'] = new_values
            entries[field].set("")
            dialog.destroy()

    tk.Button(dialog, text="Supprimer", command=confirm_delete, bg=TPW_COLORS["rouge"], fg="white", font=("Segoe UI", 10, "bold")).pack(pady=5)

def choose_text_color(field):
    color_code = colorchooser.askcolor(title=f"Couleur du texte pour {field}")
    if color_code[1]:
        COLOR_FIELDS_TEXT_COLORS[field] = color_code[1]

def submit_form():
    try:
        values = [entries[field].get() for field in FIELDS]
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_FILE)
        ws = wb.sheets[SHEET_MAIN]

        repere = "CONGES   /   R.T.T.   /   FORMATIONS   /   ABSENCES …"
        repere_row = None
        max_row = ws.cells.last_cell.row
        for row in range(1, max_row + 1):
            val = ws.cells(row, 10).value
            if isinstance(val, str) and val.strip() == repere:
                repere_row = row
                break

        if repere_row is None:
            messagebox.showerror("Erreur", "Ligne de repère non trouvée dans la colonne J.")
            wb.close()
            app.quit()
            return

        last_row = repere_row - 2
        ws.api.Rows(last_row).Insert(Shift=1)
        unique_id = uuid.uuid4().hex[:8].upper()
        ws.cells(last_row, 286).value = unique_id

        last_col = ws.range("4:4").end("right").column
        skip_cols = [3, 4, 5, 14, 20, 21]

        excel_col = 2
        i = 0

        while i < len(values) and excel_col <= last_col:
            if excel_col in skip_cols:
                excel_col += 1
                continue

            field = FIELDS[i]
            val = values[i]

            if "date" in field.lower():
                try:
                    val_dt = datetime.strptime(val.strip(), "%d/%m/%Y")
                except Exception:
                    messagebox.showerror("Erreur de date", f"Le champ '{field}' doit être au format JJ/MM/AAAA.")
                    wb.close()
                    app.quit()
                    return
                val = val_dt

            cell = ws.cells(last_row, excel_col)
            cell.value = val

            if field in COLOR_FIELDS_BG:
                color = get_option_color(field, values[i])
                if color:
                    cell.color = color

            if field in COLOR_FIELDS_TEXT_COLORS:
                cell.api.Font.Color = int(COLOR_FIELDS_TEXT_COLORS[field].replace('#', '0x00'), 16)

            for edge in [7, 8, 9, 10]:
                cell.api.Borders(edge).LineStyle = 1
                cell.api.Borders(edge).Weight = 2

            excel_col += 1
            i += 1

        for col in range(1, last_col + 1):
            if col == 1 or ws.cells(4, col).value:
                cell = ws.cells(last_row, col)
                for edge in [7, 8, 9, 10]:
                    cell.api.Borders(edge).LineStyle = 1
                    cell.api.Borders(edge).Weight = 2

        ws.cells(last_row, 7).color = "#83CCEB"
        ws.cells(last_row, 8).color = "#B5E6A2"
        ws.cells(last_row, 11).color = "#F7C7AC"
        ws.cells(last_row, 12).color = "#F7C7AC"

        chargé_val = values[2].strip().lower()
        color_exc = get_option_color("Chargé d'étude", "Exc")

        cell_d = ws.cells(last_row, 4)
        cell_j = ws.cells(last_row, 10)
        cell_t = ws.cells(last_row, 20)

        if chargé_val == "exc" and color_exc:
            cell_d.color = color_exc
            cell_j.color = color_exc
            cell_t.value = "Exc"
            cell_t.color = color_exc
        else:
            cell_d.color = None
            cell_j.color = None
            cell_t.color = None
            if cell_t.value == "Exc":
                cell_t.value = ""

        try:
            wb.app.macro("Feuil1.ColorierPlanningDate")(last_row)
        except Exception as e:
            msg = str(e)
            if "The macro may not be available" not in msg:
                print("Erreur lors de l'appel de ColorierPlanningDate:", msg)

        row_total = last_row + 2
        for col in [13, 14]:
            col_letter = xw.utils.col_name(col)
            formula = f"=SUM({col_letter}5:{col_letter}{last_row})"
            cell_total = ws.cells(row_total, col)
            cell_total.formula = formula
            for edge in [7, 8, 9, 10]:
                cell_total.api.Borders(edge).LineStyle = 1
                cell_total.api.Borders(edge).Weight = 2

        wb.save()
        wb.close()
        app.quit()
        messagebox.showinfo("Succès", "Ligne ajoutée avec bordures, couleurs et formule de total.")

    except Exception as e:
        messagebox.showerror("Erreur", str(e))


# ------ Lancement UI stylisée TPW ------

options = load_options()
COLOR_FIELDS_TEXT_COLORS = {}

root = tk.Tk()
root.title("Ajout Planning Affaires - TPW Calais")
root.geometry("1100x900")
root.configure(bg=TPW_COLORS["gris_clair"])

header = tk.Frame(root, bg=TPW_COLORS["bleu"])
header.pack(side="top", fill="x")

try:
    img_logo = Image.open("tpw_logo.png").resize((110, 60))
    logo_tk = ImageTk.PhotoImage(img_logo)
    logo_label = tk.Label(header, image=logo_tk, bg=TPW_COLORS["bleu"])
    logo_label.pack(side="left", padx=16, pady=10)
except Exception as e:
    logo_label = tk.Label(header, text="TPW", font=("Arial", 28, "bold"), bg=TPW_COLORS["bleu"], fg=TPW_COLORS["jaune"])
    logo_label.pack(side="left", padx=16, pady=10)

title = tk.Label(
    header,
    text="Ajout d'une Affaire • Planning Étude",
    font=("Segoe UI", 24, "bold"),
    bg=TPW_COLORS["bleu"],
    fg="white"
)
title.pack(side="top", pady=10, fill="x")
title.config(anchor="center", justify="center")

frame = tk.Frame(root, bg=TPW_COLORS["gris_clair"])
frame.pack(padx=18, pady=18, fill="both", expand=True)

entries = {}

for field in FIELDS:
    row = tk.Frame(frame, bg=TPW_COLORS["gris_clair"])
    label = tk.Label(
        row, text=field, anchor="w", width=30,
        font=("Segoe UI", 10, "bold"),
        bg=TPW_COLORS["gris_clair"], fg=TPW_COLORS["bleu"]
    )

    if field in MANUAL_FIELDS:
        input_widget = tk.Entry(row, width=65, font=("Segoe UI", 10))
    else:
        values = [item["value"] if isinstance(item, dict) else item for item in options.get(field, [])]
        input_widget = ttk.Combobox(row, width=50, values=values, font=("Segoe UI", 10))

        def make_add_callback(f):
            return lambda: add_option(f)
        def make_remove_callback(f):
            return lambda: remove_option(f)

        add_btn = tk.Button(row, text="+", width=3, command=make_add_callback(field), bg=TPW_COLORS["jaune"], fg=TPW_COLORS["bleu"], font=("Segoe UI", 10, "bold"))
        remove_btn = tk.Button(row, text="-", width=3, command=make_remove_callback(field), bg=TPW_COLORS["rouge"], fg="white", font=("Segoe UI", 10, "bold"))

    row.pack(pady=5, fill="x")
    label.pack(side="left")
    input_widget.pack(side="left", expand=True, fill="x", padx=(5, 0))

    if field not in MANUAL_FIELDS:
        add_btn.pack(side="left", padx=(5, 0))
        remove_btn.pack(side="left", padx=(2, 0))

    if field in COLOR_FIELDS_TEXT:
        color_btn = tk.Button(row, text="🖌 texte", command=lambda f=field: choose_text_color(f), bg=TPW_COLORS["bleu_clair"], fg="white")
        color_btn.pack(side="left", padx=(5, 0))

    entries[field] = input_widget

main_btn = tk.Button(
    root,
    text="➕ Ajouter la ligne au planning",
    font=("Segoe UI", 14, "bold"),
    command=submit_form,
    bg=TPW_COLORS["bleu"],
    fg=TPW_COLORS["jaune"],
    activebackground=TPW_COLORS["jaune"],
    activeforeground=TPW_COLORS["bleu"]
)
main_btn.pack(pady=30, ipadx=14, ipady=6)

footer = tk.Frame(root, bg=TPW_COLORS["bleu"])
footer.pack(side="bottom", fill="x")
tk.Label(
    footer,
    text="TPW Calais — Outil interne de gestion du planning étude • v1.0",
    font=("Segoe UI", 10, "italic"),
    bg=TPW_COLORS["bleu"],
    fg="white"
).pack(side="right", padx=14, pady=7)

root.mainloop()
