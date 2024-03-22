# ========================================= Import des bibliothèques ==============================================
import tkinter.messagebox as messagebox
import openpyxl
import threading
import pandas as pd
import numpy as np
import json
import sys
import itertools
import os
import shutil
import tempfile

import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# =============================================== CONFIGURATIONS ==================================================
bdd = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
file1 = None
file2 = None
color1 = '#012B65'
color2 = '#90CBFB'
color3 = '#FFC559'
font_segoe_ui = ("Segoe UI", 9)
font_mini = ("Book Antiqua", 9)
font_style = ("Book Antiqua", 15, "bold")
font_title = ("Book Antiqua", 24, "bold")
ref_entry1 = ""
date_entry1 = ""
ref_entry2 = ""
date_entry2 = ""
rotated_image = None

# =============================================== Fonction Principale ==============================================
def merge_files():
    global ref_entry1
    global date_entry1
    global ref_entry2
    global date_entry2
    name1 = os.path.basename(file1)
    name2 = os.path.basename(file2)

    if file1 is None or file2 is None:
        print("Please select both Excel files.")
        messagebox.showinfo("Error", "Please select both Excel files.")
        return

    if selected_preset.get() == "Select your preset":
        print("Please select a preset")
        messagebox.showinfo("Error", "Please select a preset")
        return

    if not all((ref_entry1, date_entry1, ref_entry2, date_entry2)):
        print("Please fill in all column names in the preset.")
        messagebox.showinfo("Error", "Please fill in all column names in the preset.")
        return

    # Afficher la boîte de dialogue de progression
    progress_dialog = show_progress_dialog()

    # Lecture des fichiers Excel
    df_file1 = pd.read_excel(file1)
    df_file2 = pd.read_excel(file2)

    # Nettoyage des noms de colonnes et des entrées de référence et de date
    ref_entry1 = ref_entry1.strip().lower()
    date_entry1 = date_entry1.strip().lower()
    ref_entry2 = ref_entry2.strip().lower()
    date_entry2 = date_entry2.strip().lower()
    df_file1.columns = [col.strip().lower() for col in df_file1.columns]
    df_file2.columns = [col.strip().lower() for col in df_file2.columns]

    # Vérification des noms de colonnes
    missing_cols_file1 = [col for col in (ref_entry1, date_entry1) if col not in df_file1.columns]
    if missing_cols_file1:
        progress_dialog.destroy()
        print(f"The following column(s) in file 1 are missing: {', '.join(missing_cols_file1)}")
        messagebox.showinfo("Error", f"The following column(s) in file 1 are missing: {', '.join(missing_cols_file1)}")
        root.update()
        return

    missing_cols_file2 = [col for col in (ref_entry2, date_entry2) if col not in df_file2.columns]
    if missing_cols_file2:
        progress_dialog.destroy()
        print(f"The following column(s) in file 2 are missing: {', '.join(missing_cols_file2)}")
        messagebox.showinfo("Error", f"The following column(s) in file 2 are missing: {', '.join(missing_cols_file2)}")
        root.update()
        return

    df_file1.rename(columns={ref_entry1: 'REFERENCE', date_entry1: 'COMPARE'}, inplace=True)
    df_file2.rename(columns={ref_entry2: 'REFERENCE', date_entry2: 'COMPARE'}, inplace=True)

    print("\nColumns in df_file1:", df_file1.columns)
    print("\nColumns in df_file2:", df_file2.columns)
    print("\nType file 1:", df_file1['COMPARE'].dtype)
    print("\nType file 2:", df_file2['COMPARE'].dtype)

    def keep_eight_digits(series):
        series_as_str = series.astype(str)
        extracted = series_as_str.str.extract('(\d{8})')[0]
        # Remplace par la séquence extraite si elle est trouvée, sinon conserve la valeur originale
        return series.where(extracted.isna(), extracted)

    # Application de la fonction aux colonnes REFERENCE avant le merge
    df_file1['REFERENCE'] = keep_eight_digits(df_file1['REFERENCE'])
    df_file2['REFERENCE'] = keep_eight_digits(df_file2['REFERENCE'])

    # Exemple de suppression des doublons basée uniquement sur la colonne 'REFERENCE'
    df_file1 = df_file1.drop_duplicates(subset=['REFERENCE'], keep='last')
    df_file2 = df_file2.drop_duplicates(subset=['REFERENCE'], keep='last')

    # Vérification du type de données dans les colonnes à comparer
    if df_file1['COMPARE'].dtype == 'datetime64[ns]':
        df_file1['COMPARE'] = pd.to_datetime(df_file1['COMPARE'], errors='coerce', format='%d/%m/%Y')
    if df_file2['COMPARE'].dtype == 'datetime64[ns]':
        df_file2['COMPARE'] = pd.to_datetime(df_file2['COMPARE'], errors='coerce', format='%d/%m/%Y')

    df_file1 = apply_date_conversion(df_file1, ['COMPARE'])
    df_file2 = apply_date_conversion(df_file2, ['COMPARE'])

    # Fusionner les deux DataFrames sur la colonne 'Doc_Achat'
    df_merged = pd.merge(df_file1, df_file2, on='REFERENCE', suffixes=('_1', '_2'))

    # TEST LOGIQUE 1 ! Comparer si les colonnes sont identiques
    df_merged['Result'] = np.where(df_merged['COMPARE_1'] == df_merged['COMPARE_2'], 'identical', 'different')

    # Convertir 'COMPARE_1' et 'COMPARE_2' en datetime, avec 'coerce' pour gérer les non-dates
    df_merged['COMPARE_1_datetime'] = pd.to_datetime(df_merged['COMPARE_1'], errors='coerce', format='%d.%m.%Y')
    df_merged['COMPARE_2_datetime'] = pd.to_datetime(df_merged['COMPARE_2'], errors='coerce', format='%d.%m.%Y')

    # TEST LOGIQUE 2 ! Calculer la différence en jours où les deux colonnes sont des dates
    df_merged['Difference'] = (df_merged['COMPARE_2_datetime'] - df_merged['COMPARE_1_datetime']).dt.days

    # Garder les lignes où au moins une des deux colonnes COMPARE n'est pas NaN
    df_filtered = df_merged[(df_merged['COMPARE_1'].notna()) | (df_merged['COMPARE_2'].notna())]

    # Renommer les colonnes
    df_output = df_filtered[['REFERENCE', 'COMPARE_1', 'COMPARE_2', 'Result', 'Difference']].copy()
    df_output.rename(
        columns={'REFERENCE': ref_entry1_var.get(), 'COMPARE_1': f"{date_entry1_var.get()} \n({name1})",
                 'COMPARE_2': f"{date_entry2_var.get()} \n({name2})"}, inplace=True)

    # Fermer la fenêtre d'attente
    progress_dialog.destroy()

    # Demander à l'utilisateur l'emplacement et le nom du fichier Excel à enregistrer
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)  # Créer un nom de fichier temporaire
        format_excel(df_output, temp_file.name)  # Formatage du fichier Excel temporaire

        # Enregistrer le fichier temporaire à l'emplacement spécifié par l'utilisateur
        shutil.copyfile(temp_file.name, save_path)
        temp_file.close()
        os.unlink(temp_file.name)

        os.startfile(save_path)  # Ouvrir le fichier après l'enregistrement
        # messagebox.showinfo("Success", "Excel file saved successfully. It will now be opened.")
        print("\nFormatted Excel file saved successfully.")

    root.update()

# ======================================== Fonctions utilitaires =====================================================
def format_excel(df_output, excel_file):
    # Exporter le DataFrame vers un fichier Excel
    df_output.to_excel(excel_file, index=False)

    # Charger le fichier Excel pour modifier avec openpyxl
    workbook = openpyxl.load_workbook(excel_file)
    ws = workbook.active

    # Styles existants
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    data_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    data_font = Font(color="000000")  # Noir

    # Appliquer le style aux données et ajuster la largeur des colonnes
    max_width = 27
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        adjusted_width = min(length, max_width)  # Ajuster la largeur à un maximum
        ws.column_dimensions[get_column_letter(column_cells[0].column)].width = (adjusted_width + 2) * 1.2
        ws.row_dimensions[1].height = 43  # Cette valeur peut être ajustée selon vos besoins

        for cell in column_cells:
            cell.fill = data_fill
            cell.font = data_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # Couleurs pour les lignes basées sur la colonne 'Result'
    fill_identical = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Vert clair
    fill_different = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Rouge clair

    # Trouver l'index de la colonne 'Result'
    result_col_idx = None
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value == 'Result':
            result_col_idx = idx
            break

    # Appliquer les couleurs
    if result_col_idx:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            result_cell = row[result_col_idx - 1]  # Ajuster l'index
            fill = fill_identical if result_cell.value == 'identical' else fill_different
            for cell in row:
                cell.fill = fill

    # Appliquer le style au header
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)

    # Sauvegarder les modifications
    workbook.save(excel_file)

def apply_date_conversion(df, date_columns):
    # Liste de formats de dates à essayer
    date_formats = ["%d/%m/%Y", "%d.%m.%Y"]

    def convert_dates(value):
        # Essayer de convertir la date en utilisant les différents formats
        for date_format in date_formats:
            try:
                return pd.to_datetime(value, format=date_format, errors='raise').strftime('%d.%m.%Y')
            except (ValueError, TypeError):
                # Si la conversion échoue, ignorer l'erreur et essayer le format suivant
                pass

        # Si aucun format ne convient, retourner la valeur originale
        return value

    for col in date_columns:
        df[col] = df[col].apply(convert_dates)
    return df

def add_date_difference_column(df_merge):
    # Assurez-vous que df_merge est une copie distincte pour éviter les avertissements
    df_temp = df_merge.copy()

    # Convertir les colonnes 'COMPARE_1' et 'COMPARE_2' en datetime
    df_merge['COMPARE_1'] = pd.to_datetime(df_merge['COMPARE_1'], format="%d/%m/%Y", errors='coerce')
    df_merge['COMPARE_2'] = pd.to_datetime(df_merge['COMPARE_2'], format="%d/%m/%Y", errors='coerce')

    # Initialiser la nouvelle colonne avec des NaN
    df_temp['Date_difference'] = pd.NaT

    # Calculer la différence en jours là où les deux colonnes sont des dates valides
    mask = df_merge['COMPARE_1'].notna() & df_merge['COMPARE_2'].notna()
    df_temp.loc[mask, 'Date_difference'] = (df_merge.loc[mask, 'COMPARE_2'] - df_merge.loc[mask, 'COMPARE_1']).dt.days

    return df_temp

def on_enter1(event):
    if file1 is None:
        event.widget.config(bg=color1, fg=color2)  # Changer la couleur lors du hover

def on_enter2(event):
    if file2 is None:
        event.widget.config(bg=color1, fg=color2)  # Changer la couleur lors du hover

def on_enter3(event):
    if file1 is None or file2 is None:
        event.widget.config(bg=color1, fg=color2)  # Changer la couleur lors du hover

def on_leave1(event):
    if file1 is None:
        event.widget.config(bg=color2, fg=color1)  # Revenir à la couleur d'origine

def on_leave2(event):
    if file2 is None:
        event.widget.config(bg=color2, fg=color1)  # Revenir à la couleur d'origine

def on_leave3(event):
    if file1 is None or file2 is None:
        event.widget.config(bg=color2, fg=color1)  # Revenir à la couleur d'origine

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

def update_image(image_label, image, angle):
    if not image_label.winfo_exists():
        return

    rotated_image2 = image.rotate(angle)  # Faire pivoter l'image
    rotated_image_tk = ImageTk.PhotoImage(rotated_image2)  # Convertir l'image en format Tkinter
    image_label.configure(image=rotated_image_tk)  # Mettre à jour l'image affichée dans le label
    image_label.image = rotated_image_tk  # Garder une référence à l'image pour éviter la suppression par le garbage collector
    root.after(20, update_image, image_label, image, angle-2)  # Appeler cette fonction après 500 millisecondes avec un angle augmenté

def generate_color_transition(start_hex, end_hex, steps):
    start_rgb = [int(start_hex[i1:i1+2], 16) for i1 in range(1, 6, 2)]
    end_rgb = [int(end_hex[i1:i1+2], 16) for i1 in range(1, 6, 2)]
    transition = [
        "#" + "".join(
            ["{:02x}".format(int(start_rgb[j] + ((end_rgb[j] - start_rgb[j]) * i / (steps-1))))
             for j in range(3)]
        ) for i in range(steps)
    ]
    return transition

def change_color():
    new_color = next(colors_cycle)
    bg_label.config(bg=new_color)
    root.after(200, change_color)  # Planifier le changement de couleur

# =========================================== Boutons de l'interface Graphique ======================================
def select_file1():
    global file1
    file1 = filedialog.askopenfilename(title="Select the 1st Excel file", filetypes=[("Excel Files", "*.xlsx; *.xlsb; *.xlsm; *.xls")])

    if file1:
        file_label1.config(text=os.path.basename(file1), bg=color2, fg=color1)
        update_button_state()  # Mettre à jour l'état du bouton
        file_label1.grid()  # Rendre le widget visible
    root.update()  # Mettre à jour l'interface utilisateur

def select_file2():
    global file2
    file2 = filedialog.askopenfilename(title="Select the 2nd Excel file", filetypes=[("Excel Files", "*.xlsx; *.xlsb; *.xlsm; *.xls")])
    if file2:
        file_label2.config(text=os.path.basename(file2), bg=color2, fg=color1)
        update_button_state()  # Mettre à jour l'état du bouton
        file_label2.grid()  # Rendre le widget visible
    root.update()  # Mettre à jour l'interface utilisateur

def show_progress_dialog():
    progress_dialog = tk.Toplevel()
    progress_dialog.title("Processing")
    progress_dialog.configure(background=color2)  # Définir la couleur d'arrière-plan
    progress_dialog.resizable(False, False)  # Désactiver la possibilité de redimensionner
    # progress_dialog.overrideredirect(True)  # Supprimer la barre de titre
    progress_dialog.attributes("-topmost", True)  # Garder la fenêtre au premier plan
    center_window(progress_dialog, 800, 400)  # Centrer la fenêtre sur l'écran
    progress_label = tk.Label(progress_dialog, text="Please wait while processing...", bg=color2, fg=color1, font=font_title)
    progress_label.pack(padx=20, pady=20)

    # Charger l'image
    path = os.path.join(bdd, 'engine.png')
    image = Image.open(path)
    image = image.resize((image.width // 3, image.height // 3))  # Redimensionner l'image
    image_tk = ImageTk.PhotoImage(image)

    # Créer un label pour afficher l'image
    image_label = tk.Label(progress_dialog, image=image_tk, bg=color2)
    image_label.pack()

    # Mettre à jour l'image affichée dans le label avec une rotation toutes les 500 millisecondes
    update_image(image_label, image, 90)

    return progress_dialog

def merge_files_with_progress_dialog():
    # Démarrer un thread pour exécuter la fonction merge_files en arrière-plan
    merge_thread = threading.Thread(target=merge_files)
    merge_thread.start()

def update_button_state():
    global ref_entry1, date_entry1, ref_entry2, date_entry2
    if file1:
        btn_select_file1.config(bg=color1, fg=color2)
    if file2:
        btn_select_file2.config(bg=color1, fg=color2)
    if file1 and file2 and selected_preset.get() != "Select your preset":
        btn_merge_files.config(bg=color3)  # Si les deux fichiers sont sélectionnés, le bouton devient bleu
    else:
        btn_merge_files.config(bg=color2)  # Sinon, le bouton conserve sa couleur par défaut

def save_button_data(button_index, title, ref_col_name_1, ref_col_name_2, date_col_name_1, date_col_name_2):
    button_datas = {
        "Title": title,
        "Ref_Column_Name_1": ref_col_name_1,
        "Ref_Column_Name_2": ref_col_name_2,
        "Date_Column_Name_1": date_col_name_1,
        "Date_Column_Name_2": date_col_name_2
    }
    print (title, ref_col_name_1, date_col_name_1, ref_col_name_2, date_col_name_2)
    with open(f'button_{button_index}_data.txt', 'w') as fff:
        json.dump(button_datas, fff)

def open_button_window(button_index):
    global selected_preset
    button = buttons_list[button_index - 1]  # -1, car les index commencent à 1, mais les listes commencent à 0.

    # Fonction pour effacer le contenu des zones de texte
    def clear_fields():
        for entry in entry_fields:
            entry.delete(0, tk.END)
        title_entry.delete(0, tk.END)  # Effacer le titre également

    # Fonction pour sauvegarder les données dans un fichier texte
    def save_data():
        # Récupérer les valeurs des entrées
        title = title_entry.get().strip()  # Obtenir le titre en supprimant les espaces inutiles
        if not title:  # Si le titre est vide
            title = f"Preset {button_index}"  # Utiliser le texte par défaut
        ref_col_name_1 = ref_col_entries[0].get()
        ref_col_name_2 = ref_col_entries[1].get()
        date_col_name_1 = date_col_entries[0].get()
        date_col_name_2 = date_col_entries[1].get()

        # Sauvegarder les données dans le fichier texte
        save_button_data(button_index, title, ref_col_name_1, ref_col_name_2, date_col_name_1, date_col_name_2)

        # Mettre à jour le texte du bouton avec le titre entré par l'utilisateur
        button.config(text=title)

        # Mettre à jour le menu déroulant avec les nouveaux titres des boutons presets
        update_preset_menu()

        # Mettre à jour la sélection du menu déroulant
        selected_preset.set(button_texts[button_index - 1])

        # Fermer la fenêtre
        window.destroy()

        # Appeler la fonction pour mettre à jour le preset sélectionné
        update_selected_preset(title)

    # Créer une nouvelle fenêtre
    window = tk.Toplevel(root)
    window.configure(bg=color2)
    window.title("Button Details")
    window.geometry("290x245")
    window.update_idletasks()
    window.update_idletasks()  # Assurer que toutes les tâches en attente sont effectuées
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    window_w = window.winfo_width()
    window_h = window.winfo_height()
    x9 = (screen_w - window_w) // 2
    y9 = (screen_h - window_h) // 2
    window.geometry(f"+{x9}+{y9}")  # Positionner la fenêtre au centre de l'écran

    # Créer les widgets dans la fenêtre
    title_label = tk.Label(window, text="Preset title:", font=font_mini, bg=color2)
    title_label.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    title_entry = tk.Entry(window, font=font_mini)
    title_entry.grid(row=0, column=1, columnspan=2, sticky="we", padx=10, pady=10)

    file_label = tk.Label(window, text="File 1                                          File 2", font=("Book Antiqua", 10, "bold"), bg=color2)
    file_label.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=5, pady=10)

    ref_col_label = tk.Label(window, text="Reference column name:", font=font_mini, bg=color2)
    ref_col_label.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=5, pady=0)

    date_col_label = tk.Label(window, text="Comparison column name:", font=font_mini, bg=color2)
    date_col_label.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=5, pady=0)

    # Créer les widgets pour les colonnes de référence
    ref_col_entry1 = tk.Entry(window, font=font_mini)
    ref_col_entry1.grid(row=3, column=0, sticky="nsew", padx=10, pady=(2,10))

    ref_col_entry2 = tk.Entry(window, font=font_mini)
    ref_col_entry2.grid(row=3, column=1, sticky="nsew", padx=10, pady=(2,10))

    # Créer les widgets pour les colonnes de date
    date_col_entry1 = tk.Entry(window, font=font_mini)
    date_col_entry1.grid(row=5, column=0, sticky="nsew", padx=10, pady=(2,10))

    date_col_entry2 = tk.Entry(window, font=font_mini)
    date_col_entry2.grid(row=5, column=1, sticky="nsew", padx=10, pady=(2,10))

    # Créer les listes contenant les widgets
    ref_col_entries = [ref_col_entry1, ref_col_entry2]
    date_col_entries = [date_col_entry1, date_col_entry2]

    # Créer une liste regroupant tous les champs d'entrée
    entry_fields = ref_col_entries + date_col_entries

    # Charger les données depuis le fichier texte s'il existe
    try:
        with open(f'button_{button_index}_data.txt', 'r') as f1:
            button_var = json.load(f1)
            title_entry.insert(0, button_var["Title"])
            ref_col_entries[0].insert(0, button_var["Ref_Column_Name_1"])
            ref_col_entries[1].insert(0, button_var["Ref_Column_Name_2"])
            date_col_entries[0].insert(0, button_var["Date_Column_Name_1"])
            date_col_entries[1].insert(0, button_var["Date_Column_Name_2"])
    except FileNotFoundError:
        pass

    # Créer les boutons "clear" et "save"
    clear_button = tk.Button(window, text="Clear", bg="red", fg="white", font=("Book Antiqua", 10, "bold"), command=clear_fields)
    clear_button.grid(row=6, column=0, sticky="we", padx=10, pady=10)

    save_button = tk.Button(window, text="Save", bg="green", fg="white", font=("Book Antiqua", 10, "bold"), command=save_data)
    save_button.grid(row=6, column=1, sticky="we", padx=10, pady=10)

def update_selected_preset(preset):
    global ref_entry1
    global ref_entry2
    global date_entry1
    global date_entry2

    print("Selected preset:", preset)

    # Réinitialiser la couleur de tous les boutons
    for button in buttons_list:
        button.config(bg=color2, fg=color1)

    # Trouver l'index du preset sélectionné dans la liste des titres des boutons
    selected_index = button_texts.index(preset)

    # Changer la couleur du bouton associé au preset sélectionné
    buttons_list[selected_index].config(bg=color3, fg=color1)

    # Changer le texte sélectionné dans le menu déroulant
    selected_preset.set(button_texts[selected_index])

    # Petite condition
    if file1 and file2 and selected_preset.get() != "Select your preset":
        btn_merge_files.config(bg=color3)  # Si les deux fichiers sont sélectionnés + le preset, le bouton devient jaune

    try:
        # Charger les données depuis le fichier texte correspondant au titre du preset sélectionné
        with open(f'button_{button_texts.index(preset) + 1}_data.txt', 'r') as f5:
            button_data_txt = json.load(f5)
            ref_entry1 = button_data_txt.get("Ref_Column_Name_1", "")
            date_entry1 = button_data_txt.get("Date_Column_Name_1", "")
            ref_entry2 = button_data_txt.get("Ref_Column_Name_2", "")
            date_entry2 = button_data_txt.get("Date_Column_Name_2", "")
            # Mettre à jour les variables globales avec les nouvelles valeurs
            ref_entry1_var.set(ref_entry1)
            date_entry1_var.set(date_entry1)
            ref_entry2_var.set(ref_entry2)
            date_entry2_var.set(date_entry2)
            # Imprimer les valeurs mises à jour pour vérification
            print("ref1: ", ref_entry1, " - ref2: ", date_entry1, "\ndate1:", ref_entry2, " - date2:", date_entry2)
    except FileNotFoundError:
        # Si le fichier du preset sélectionné n'est pas trouvé, laisser les variables inchangées
        print("Preset data file not found for", preset)

    root.update()

def update_preset_menu():
    # Vide le menu déroulant actuel
    preset_menu['menu'].delete(0, 'end')

    # Liste pour stocker les nouveaux titres des boutons presets
    new_button_texts = []

    # Parcourez les index des boutons de 1 à 7 pour récupérer les titres des boutons presets
    for i7 in range(1, 8):
        try:
            # Ouvrez le fichier correspondant au bouton et chargez les données JSON
            with open(f'button_{i7}_data.txt', 'r') as f7:
                button_data7 = json.load(f7)
                title_button7 = button_data7.get("Title",
                                               f"Preset {i7}")  # Obtenez le titre du bouton ou utilisez "Preset {i}" par défaut
                new_button_texts.append(title_button7)  # Ajoutez le titre à la liste des textes de bouton
        except FileNotFoundError:
            # Si le fichier n'est pas trouvé, utilisez "Preset {i}" par défaut
            new_button_texts.append(f"Preset {i7}")

    # Mettre à jour le menu déroulant avec les nouveaux titres des boutons presets
    for txt2 in new_button_texts:
        preset_menu['menu'].add_command(label=txt2, command=lambda preset=txt2: update_selected_preset(preset))

    # Mettre à jour la liste button_texts
    button_texts[:] = new_button_texts

    # Si un preset était déjà sélectionné, le remettre à jour avec le nouveau texte
    selected_preset_text = selected_preset.get()
    if selected_preset_text in new_button_texts:
        selected_preset.set(selected_preset_text)


# ================================================= Interface Graphique ============================================
root = tk.Tk()
root.title("Merge Excel Files")

# Définir l'icône de la fenêtre
icon_path = os.path.join(bdd, 'engine.ico')
root.iconbitmap(icon_path)

ref_entry1_var = tk.StringVar()
date_entry1_var = tk.StringVar()
ref_entry2_var = tk.StringVar()
date_entry2_var = tk.StringVar()

# Configuration de la grille
for i in range(7):
    root.grid_rowconfigure(i, weight=1)  # Toutes les lignes ont le même poids
    root.grid_columnconfigure(i, weight=1)  # Toutes les colonnes ont le même poids
    root.grid_rowconfigure(i, weight=1)

# Création des widgets
bg_path = os.path.join(bdd, 'intro1.png')
bg_img = tk.PhotoImage(file=bg_path)
bg_img = bg_img.subsample(1)

# Créer une étiquette pour l'image de fond
bg_label = tk.Label(root, image=bg_img, bg=color1)
bg_label.grid(row=0, column=0, rowspan=7, columnspan=7, sticky="nsew")  # Spanning sur toute la grille
bg_label.lower()

# Ajouter une étiquette pour afficher le nom du fichier sélectionné
file_label1 = tk.Label(root, text="", font=font_mini, bg=color1, fg=color2, padx=0, pady=0, wraplength=200)
file_label1.grid(row=2, column=1, padx=40, pady=20, sticky="nsew")
file_label1.grid_remove()  # Rendre le widget invisible au départ

file_label2 = tk.Label(root, text="", font=font_mini, bg=color1, fg=color2, padx=0, pady=2, wraplength=200)
file_label2.grid(row=2, column=5, padx=40, pady=20, sticky="nsew")
file_label2.grid_remove()  # Rendre le widget invisible au départ

# Ajouter une étiquette pour expliquer ce que fait l'application
app_explanation = tk.Label(root, text="Select two Excel files to calculate date differences based on the column of your choice.",
                           font=font_style, bg=color1, fg=color2)
app_explanation.grid(row=0, column=0, columnspan=7, sticky="nsew")

# Bouton pour sélectionner le fichier 1
btn_select_file1 = tk.Button(root, text="Select 1st File", command=select_file1, bg=color2, font=font_style,
                             relief=tk.SOLID, bd=2, fg=color1, width=8, height=2, cursor="target", wraplength=200)
btn_select_file1.grid(row=1, column=1, padx=40, pady=(20,0), sticky="nsew")
btn_select_file1.bind("<Enter>", lambda event: on_enter1(event))
btn_select_file1.bind("<Leave>", lambda event: on_leave1(event))

# Bouton pour sélectionner le fichier 2
btn_select_file2 = tk.Button(root, text="Select 2nd File", command=select_file2, bg=color2, font=font_style,
                             relief=tk.SOLID, bd=2, fg=color1, width=8, height=2, cursor="target", wraplength=200)
btn_select_file2.grid(row=1, column=5, padx=40, pady=(20,0), sticky="nsew")
btn_select_file2.bind("<Enter>", lambda event: on_enter2(event))
btn_select_file2.bind("<Leave>", lambda event: on_leave2(event))

# Bouton pour lancer la fusion des données
btn_merge_files = tk.Button(root, text="Fusion", command=merge_files_with_progress_dialog, bg=color2, font=font_title,
                            relief=tk.SOLID, bd=2, fg=color1, width=8, height=1, cursor="target", wraplength=200)
btn_merge_files.grid(row=1, column=3, padx=100, pady=(20,0), sticky="nsew")
btn_merge_files.bind("<Enter>", lambda event: on_enter3(event))
btn_merge_files.bind("<Leave>", lambda event: on_leave3(event))

#-------------------------------------------------- PRESETS-------------------------------------------------------------
preset_frame = tk.Frame(root, bg=color1, bd=2, relief=tk.SOLID)
preset_frame.grid(row=6, column=0, columnspan=7, sticky="nsew")

preset_label = tk.Label(preset_frame, text="Presets", font=font_style, bg=color1, fg=color2, bd=2)
preset_label.pack(fill="x")

# Créer les boutons dans chaque colonne avec un léger padding
button_texts = []  # Initialisez une liste vide pour stocker les titres des boutons

# Parcourez les index des boutons de 1 à 7.
for i in range(1, 8):
    try:
        # Ouvrez le fichier correspondant au bouton et chargez les données JSON
        with open(f'button_{i}_data.txt', 'r') as f:
            button_data = json.load(f)
            title_button = button_data.get("Title", f"Preset {i}")  # Obtenez le titre du bouton ou utilisez "Preset {i}" par défaut
            button_texts.append(title_button)  # Ajoutez le titre à la liste des textes de bouton
    except FileNotFoundError:
        # Si le fichier n'est pas trouvé, utilisez "Preset {i}" par défaut
        button_texts.append(f"Preset {i}")

# Création des boutons presets
buttons_list = []
for i, txt in enumerate(button_texts, start=1):
    btn = tk.Button(preset_frame, text=txt, bg=color2, font=font_mini, relief=tk.SOLID, bd=2, fg=color1, cursor="hand2", command=lambda index=i, text=txt: open_button_window(index))
    btn.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    buttons_list.append(btn)

# Menu déroulant
selected_preset = tk.StringVar(root)
selected_preset.set("Select your preset")  # Définir la valeur par défaut

# Créer le menu déroulant
preset_menu = tk.OptionMenu(root, selected_preset, *button_texts, command=lambda preset: update_selected_preset(preset))

# Personnaliser le menu déroulant
preset_menu.config(font=font_style, bg=color1, fg=color2, height=1)  # Modifier la police, la couleur de fond et la hauteur
preset_menu.grid(row=4, column=3, sticky="nsew", padx=70 , pady=(100,0))

# Changer les couleurs du fond
colors = []
colors += generate_color_transition(color1, color2, 128)  # De color1 à color2
colors += generate_color_transition(color2, color3, 128)  # De color2 à color3
colors += generate_color_transition(color3, color1, 128)  # De color3 à color1
colors_cycle = itertools.cycle(colors)
root.after(500, change_color)  # Change la couleur du fond

# Fin
center_window(root, 1000, 500)  # Redimensionne la fenêtre
root.mainloop()  # démarrage de la boucle principale
