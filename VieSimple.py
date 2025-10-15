import sys
import os
import pandas as pd
import tabula
import PyPDF2
import pdfplumber
from datetime import datetime
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from icalendar import Calendar, Event
import win32com.client as win32
import calendar
import locale
import fitz  # PyMuPDF
import re
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import green, red, black
from reportlab.lib.pagesizes import A4, landscape
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QTabWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QCheckBox,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox
)
from PyQt5.QtCore import Qt

# -------------------------------------------------------------------
# Fonction de conversion PDF -> csv (sans header)
# -------------------------------------------------------------------

def convert_pdf_to_csv(pdf_file):
    # Extraction du texte depuis la première page du PDF
    with open(pdf_file, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        first_page = reader.pages[0]
        text = first_page.extract_text()

    # Première tentative : recherche d'une date au format "DD/MM/YYYY"
    date_match = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    if date_match:
        pdf_date = date_match.group(1)
    else:
        # Sinon, recherche du format DOF/AA MM JJ, par exemple "DOF/250306"
        dof_match = re.search(r"DOF/(\d{6})", text)
        if dof_match:
            dof_str = dof_match.group(1)  # ex : "250306"
            short_year = dof_str[:2]       # AA
            month = dof_str[2:4]           # MM
            day = dof_str[4:6]             # JJ
            if int(short_year) < 50:
                full_year = f"20{short_year}"
            else:
                full_year = f"19{short_year}"
            pdf_date = f"{day}/{month}/{full_year}"
        else:
            pdf_date = "01/01/1970"  # Valeur par défaut

    # Extraction des tableaux dans le PDF avec Tabula
    dfs = tabula.read_pdf(pdf_file, pages="all", multiple_tables=True)
    
    # Recherche d'un tableau comportant 7 colonnes
    table_df = None
    for df in dfs:
        if df.shape[1] == 7:
            table_df = df
            break

    if table_df is not None:
        prgrm_name = "NTR"
        # Traitement par tableau à 7 colonnes
        table_df.columns = [
            "Immatriculation", "Vol", "Origine",
            "Heure_Depart", "Heure_Arrivee", "Destination", "Degagement"
        ]
        output_rows = []
        for index, row in table_df.iterrows():
            immat_raw     = str(row["Immatriculation"]).strip()
            vol_raw       = str(row["Vol"]).strip()
            origine       = str(row["Origine"]).strip()
            heure_depart  = str(row["Heure_Depart"]).strip()
            heure_arrivee = str(row["Heure_Arrivee"]).strip()
            destination   = str(row["Destination"]).strip()
            # Extraction des chiffres du numéro de vol (après "NTR")
            vol_digits_match = re.search(r'NTR(\d+)', vol_raw)
            vol_digits = vol_digits_match.group(1) if vol_digits_match else ""
            # Reformater l'immatriculation ("ABCDE" devient "A-BCDE")
            immat_formatted = (immat_raw[0] + "-" + immat_raw[1:]) if len(immat_raw) > 1 else immat_raw

            new_row = [
                pdf_date,       # Colonne 1 : Date extraite
                "",             # Colonne 2 : Cellule vide
                "NTR",          # Colonne 3 : Texte constant
                vol_digits,     # Colonne 4 : Chiffres du vol
                "",
                immat_formatted,# Colonne 5 : Immatriculation formatée
                "AT72600",      # Colonne 6 : Texte constant
                origine,        # Colonne 7 : Origine
                heure_depart,   # Colonne 8 : Heure de départ
                destination,    # Colonne 9 : Destination
                heure_arrivee,  # Colonne 10: Heure d'arrivée (ordre inversé)
                ""              # Colonne 12: Cellule vide
            ]
            output_rows.append(new_row)

        output_df = pd.DataFrame(output_rows)
        csv_file = os.path.splitext(pdf_file)[0] + ".csv"
        output_df.to_csv(csv_file, index=False, header=True, sep=';')
    else:
        # Si aucun tableau à 7 colonnes n'est détecté, on traite le PDF comme du texte brut (par exemple, un plan de vol)
        prgrm_name = "TOA"
        with open(pdf_file, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            full_text = ""
            for page in reader.pages:
                full_text += page.extract_text() + "\n"

        # Extraction de tous les blocs de texte délimités par des parenthèses
        blocks = re.findall(r"\((.*?)\)", full_text, re.DOTALL)
        if not blocks:
            raise ValueError("Aucun bloc de texte n'a été trouvé dans le PDF !")

        output_rows = []
        for block in blocks:
            # Découpe du bloc en lignes, en éliminant les lignes vides
            lines = [line.strip() for line in block.splitlines() if line.strip()]
            
            # Extraction de l'en-tête du vol (ex. "FPL-TOA42-IN") pour obtenir l'opérateur et le numéro de vol
            header_match = re.search(r"FPL-([A-Z]+)(\d+)-IN", block)
            operator = header_match.group(1) if header_match else ""
            flight_number = header_match.group(2) if header_match else ""
            
            # Extraction du type d'appareil à partir de la deuxième ligne (ex. "-DHC6/...")
            aircraft_type = ""
            if len(lines) >= 2:
                ac_match = re.match(r"-([A-Z0-9]+)", lines[1])
                if ac_match:
                    aircraft_type = ac_match.group(1)
            
            # Extraction des informations sur le terrain de départ à partir de la troisième ligne.
            # La logique ici est : après la ligne du type d'appareil, le terrain de départ est donné
            # sous la forme d’un code de 4 lettres, éventuellement suivi d’un horaire.
            departure_code = ""
            departure_time = ""
            if len(lines) >= 3:
                dep_line = lines[2]
                # Enlever le tiret initial, s'il est présent
                if dep_line.startswith("-"):
                    dep_line = dep_line[1:]
                # Les 4 premiers caractères correspondent au code du terrain de départ
                departure_code = dep_line[:4]
                # S'il reste des chiffres après, on prélève l'heure de départ
                if len(dep_line) > 4:
                    departure_time = dep_line[4:]
            
            # Pour tenter d'extraire une durée de vol et calculer l'heure d'arrivée,
            # on vérifie si une cinquième ligne existe et commence par "-NTTB"
            arr_code = ""
            arr_time = ""
            flight_duration = ""
            dep_time = ""
            if len(lines) >= 5:
                # Vérifier si la cinquième ligne est de la forme "-NTTBXXXX..."
                if lines[4].startswith("-"):
                    # Supposons que la durée est codée sur 4 chiffres après "-NTTB"
                    duration_part = lines[4][6:10]  # "-NTTB" correspond à 5 caractères, on saute le symbole et le code.
                    flight_duration = duration_part
                    arr_code = lines[4][1:5]  # Pour indiquer que c'est le code d'arrivée associé à la durée
                    # Calcul de l'heure d'arrivée si l'heure de départ et la durée sont disponibles
                    if departure_time:
                        try:
                                    
                            dep_hours = int(departure_time[:2])
                            dep_minutes = int(departure_time[2:])
                            dep_total = dep_hours * 60 + dep_minutes
                            dur = int(flight_duration)
                            arr_total = dep_total + dur
                            arr_hours = (arr_total // 60) % 24
                            arr_minutes = arr_total % 60
                            arr_time = f"{arr_hours:02d}:{arr_minutes:02d}"
                            dep_time = f"{dep_hours:02d}:{dep_minutes:02d}"
                        except Exception:
                            arr_time = ""
                            dep_time = ""
                    
                        
            
            # Extraction de l'immatriculation (la chaîne qui suit "REG/")
            reg_match = re.search(r"REG/([A-Z0-9]+)", block)
            registration = reg_match.group(1) if reg_match else ""
            # Reformater l'immatriculation ("ABCDE" devient "A-BCDE")
            registration_formatted = (registration[0] + "-" + registration[1:]) if len(registration) > 1 else registration
            heure_min = datetime.strptime('00:00', '%H:%M').time()
            heure_max = datetime.strptime('12:01', '%H:%M').time()
            arr_time_converted = datetime.strptime(arr_time, '%H:%M').time()
            dep_time_converted = datetime.strptime(dep_time, '%H:%M').time()
            pdf_date_obj = datetime.strptime(pdf_date, '%d/%m/20%y')
            # Constitution de la ligne (row) avec des colonnes vides insérées entre Operator, FlightNumber et AircraftType
            row = [
                (pdf_date_obj - timedelta(days=1)) if (
                (heure_min < arr_time_converted <= heure_max) or
                (heure_min < dep_time_converted <= heure_max)
                ) else pdf_date, # Colonne 1
                "",              # Colonne vide
                operator,        # Colonne "Operator"
                flight_number,   # Colonne "FlightNumber"
                "",              # Colonne vide
                registration_formatted,    # Immatriculation
                aircraft_type,   # Colonne "AircraftType"
                departure_code,  # Code du terrain de départ (extraction de la ligne 3)
                dep_time,        # Heure de départ, s'il y a des chiffres après le code
                arr_code,        # Code d'arrivée (si extrait)
                arr_time,        # Heure d'arrivée calculée
                ""               #Colonne vide      
            ]
            output_rows.append(row)

        # Création du DataFrame avec les colonnes souhaitées, y compris les colonnes vides
        df = pd.DataFrame(
            output_rows,
            columns=[
                "date",
                "",              # Colonne vide
                "Operator",      # Colonne pour l'opérateur
                "FlightNumber",  # Colonne pour le numéro de vol
                "",              # Colonne vide
                "Registration",   # Immatriculation
                "AircraftType",  # Colonne pour le type d'appareil
                "DepCode",       # Code du terrain de départ
                "DepTime",       # Heure de départ (renseignée après le code, si présente)
                "ArrCode",       # Code d'arrivée (le cas échéant)
                "ArrTime",       # Heure d'arrivée calculée (le cas échéant)    
                ""
            ]
        )

        csv_file = os.path.splitext(pdf_file)[0] + "_converted.csv"
        df.to_csv(csv_file, index=False, sep=';')
    
    # Renommage du fichier CSV pour inclure la date (jour et mois en lettres françaises)
    try:
        locale.setlocale(locale.LC_TIME, 'fr_FR')
    except locale.Error:
        locale.setlocale(locale.LC_TIME, 'fr_FR.utf8')
    day, month, _ = pdf_date.split('/')
    month_name = calendar.month_name[int(month)]
    new_filename = f"PgVol{prgrm_name} {day} {month_name}.csv"
    directory = os.path.dirname(csv_file)
    new_file_path = os.path.join(directory, new_filename)
    os.rename(csv_file, new_file_path)
    
    return new_file_path, pdf_date, prgrm_name
    
    
def extract_shifts_from_pdf(pdf_path,start_date):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    
    agents_par_jour = {}
    jour_offset = 0  # Permet de compter correctement les jours après mars
    lines = text.splitlines()
    year = date_obj.year
    month = date_obj.month
    nb_jour_mois = calendar.monthrange(year, month)[1]
    
    for line in lines:
        # Chercher une ligne qui contient un numéro de semaine sous le format S1, S2, ...
        if re.search(r'\bS[0-9]{1,2}\b', line):  
            matches = re.findall(r"(\d{1,2})\s+([A-Z]{3})\s+([A-Z]{3})", line)
            for jour, j1, j2 in matches:
                jour_num = int(jour)
                if jour_num == 1 and jour_offset == 0:
                    jour_offset = nb_jour_mois  # Dès qu'on atteint la fin du mois, on ajoute un décalage
                if jour_offset > 0:
                    jour_num += jour_offset
                agents_par_jour[jour_num] = (j1, j2)

    return dict(sorted(agents_par_jour.items()))  # Tri des jours pour garantir l'ordre correct

def generate_ics(shifts_by_agent, output_dir="./"):
    for agent, shifts in shifts_by_agent.items():
        if not shifts:
            continue
        
        cal = Calendar()
        for start, end in shifts:
            event = Event()
            event.add("summary", f"Vacation {agent}")
            event.add("dtstart", start)
            event.add("dtend", end)
            event.add("description", "Vacation")
            cal.add_component(event)
        
        ics_filename = os.path.join(output_dir, f"{agent}.ics")
        with open(ics_filename, "wb") as f:
            f.write(cal.to_ical())
        print(f"Fichier généré : {ics_filename}")
        
        

#--------------------------------------------------------------------
#Widget pour la generation de strip
#--------------------------------------------------------------------
import fitz  # PyMuPDF
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.colors import green, red, black
from datetime import datetime, timedelta
import re
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog, QLabel

def extraire_texte_pdf_compat(chemin_pdf):
            rows = extraire_table_par_colonnes(chemin_pdf)   # ta fonction existante
            texte = normalize_and_join_rows(rows)
            return texte

def extraire_table_par_colonnes(chemin_pdf, seuil_ligne=2.0, seuil_colonne=20.0, debug=False):
    """
    Extrait le texte du PDF en regroupant les fragments par lignes et colonnes.
    Si une colonne est vide, insère une cellule vide pour garder la structure.
    """
    doc = fitz.open(chemin_pdf)
    lignes_finales = []
    positions_colonnes = []

    # --- Première passe : détection des positions X globales des colonnes ---
    for page in doc:
        blocs = page.get_text("blocks")
        x_positions_page = sorted(set(round(b[0], 1) for b in blocs))
        positions_colonnes.extend(x_positions_page)

    # Grouper les X proches pour définir les colonnes globales
    positions_colonnes = sorted(set(round(x) for x in positions_colonnes))
    colonnes_groupées = []
    for x in positions_colonnes:
        if not colonnes_groupées or abs(x - colonnes_groupées[-1]) > seuil_colonne:
            colonnes_groupées.append(x)

    if debug:
        print(f"Colonnes détectées ({len(colonnes_groupées)}): {colonnes_groupées}")

    # --- Deuxième passe : extraction structurée page par page ---
    for page in doc:
        blocs = page.get_text("blocks")
        lignes_temp = {}

        for b in blocs:
            x0, y0, _, _, contenu = b[:5]
            contenu = contenu.strip()
            if not contenu:
                continue

            # Trouver la ligne correspondante (verticalement proche)
            y_found = None
            for y_exist in lignes_temp.keys():
                if abs(y_exist - y0) < seuil_ligne:
                    y_found = y_exist
                    break
            if y_found is None:
                y_found = y0
                lignes_temp[y_found] = {}

            # Trouver la colonne la plus proche
            col_index = min(range(len(colonnes_groupées)), key=lambda i: abs(colonnes_groupées[i] - x0))
            lignes_temp[y_found][col_index] = contenu

        # Trier et compléter les lignes
        for y in sorted(lignes_temp.keys()):
            ligne = []
            for i in range(len(colonnes_groupées)):
                ligne.append(lignes_temp[y].get(i, "_"))  # "_" si colonne absente
            lignes_finales.append(ligne)

            if debug:
                print(f"Ligne y={y:.1f}: {ligne}")

    return lignes_finales


def normalize_and_join_rows(rows):
    """
    rows : list of lists (chaque row = colonnes détectées)
    Retourne : texte normalisé (une info par ligne)
    """

    # --- Mots parasites à ignorer ---
    mots_ignores = {
        "IRGHO", "ITHO", "CHTR", "SUPP", "IRDATE", "IRGAV", "IRSUPP",
        "IRHOU", "IRDEP", "IRARR", "IRSU", "IRSUP", "IRCHTR", "IRFAB",
        "IRTRA", "IRTHO", "IR", "IRDATEITHO"
    }

    vols = []              # liste de vols (chaque vol = liste de tokens)
    vol_courant = []       # tokens du vol en cours
    lignes_finales = []    # texte brut final
    date_regex = re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b")

    # --- 1️⃣ Parcours des lignes extraites du PDF ---
    for row in rows:
        tokens = [c.strip() for c in row if c and c.strip() and c.strip() != "_"]
        if not tokens:
            continue

        ligne = " ".join(tokens)

        # 🟩 Conserver toutes les lignes contenant une date ou "Programme"
        if any(date_regex.search(tok) for tok in tokens) or "Programme" in ligne or "BRIA" in ligne:
            lignes_finales.append(ligne)
            continue

        # --- Fusion et nettoyage ---
        merged = []
        i = 0
        while i < len(tokens):
            t = tokens[i]

            if t.upper() in mots_ignores:
                i += 1
                continue

            # Cas VT + numéro
            if t.upper() == "VT" and i + 1 < len(tokens) and re.match(r'^\d+$', tokens[i + 1]):
                merged.append("VT" + tokens[i + 1])
                i += 2
                continue

            # Cas immatriculation coupée
            if re.fullmatch(r'F-?', t, re.IGNORECASE) and i + 1 < len(tokens):
                nxt = tokens[i + 1].strip().upper()
                if re.match(r'^[A-Z0-9]{3,6}$', nxt):
                    merged.append("F-" + nxt)
                    i += 2
                    continue

            # Cas FOPFN -> F-OPFN
            if re.fullmatch(r'F[A-Z0-9]{3,6}', t, re.IGNORECASE) and not t.startswith("F-"):
                merged.append(t[0] + "-" + t[1:].upper())
                i += 1
                continue

            # Cas NTR + numéro
            if t.upper() == "NTR" and i + 1 < len(tokens) and re.match(r'^\d+$', tokens[i + 1]):
                merged.append("NTR" + tokens[i + 1])
                i += 2
                continue

            merged.append(t)
            i += 1

        # Nettoyage final : suppression des mots parasites isolés
        cleaned = [m for m in merged if m.upper() not in mots_ignores]

        # --- Détection de début de vol ---
        for token in cleaned:
            if re.match(r'^(VT|NT|NTR)\d+', token):
                # Nouveau vol => sauvegarder le précédent
                if vol_courant:
                    vols.append(vol_courant)
                    vol_courant = []
            vol_courant.append(token)

    # Ajouter le dernier vol
    if vol_courant:
        vols.append(vol_courant)

    # --- 2️⃣ Formatage final : chaque info sur une ligne ---
    lignes_finales.append("")  # saut de ligne avant les vols
    for idx, vol in enumerate(vols, start=1):
        lignes_finales.append(str(idx))
        lignes_finales.extend(vol)
        lignes_finales.append("")  # ligne vide entre vols

    texte_final = "\n".join(lignes_finales)
    texte_final = re.sub(r"\s{2,}", " ", texte_final).strip()

    return texte_final




class StrippingWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        self.label_info = QLabel("Sélectionnez un TDS Air Tahiti au format PDF pour générer les strips :")
        layout.addWidget(self.label_info)

        # Sélection du fichier PDF
        self.lineEdit_pdf = QLineEdit()
        self.lineEdit_pdf.setReadOnly(True)
        layout.addWidget(self.lineEdit_pdf)

        self.button_select_file = QPushButton("📁 Sélectionner le fichier PDF")
        self.button_select_file.clicked.connect(self.select_file)
        layout.addWidget(self.button_select_file)

        # Zone de dépôt
        self.dropLabel = DropLabel(self.lineEdit_pdf, self)
        layout.addWidget(self.dropLabel)

        # Bouton pour générer
        self.button_generate = QPushButton("✈️ Générer les strips")
        self.button_generate.clicked.connect(self.generer_strips)
        layout.addWidget(self.button_generate)

        self.label_status = QLabel("")
        layout.addWidget(self.label_status)

    def select_file(self):
        chemin, _ = QFileDialog.getOpenFileName(self, "Sélectionner un fichier PDF", "", "Fichiers PDF (*.pdf)")
        if chemin:
            self.lineEdit_pdf.setText(chemin)
            self.label_status.setText(f"✅ Fichier sélectionné : {chemin}")


    
    def generer_strips(self):
        chemin_pdf = self.lineEdit_pdf.text()
        if not chemin_pdf:
            self.label_status.setText("⚠️ Aucun fichier sélectionné.")
            return

        texte = extraire_texte_pdf_compat(chemin_pdf)
        print(texte)
        
        vols = self.extraire_vols(texte)
        jour, mois, annee = self.extraire_date(texte)
        date_locale = datetime(annee, mois, jour)
        vols_depart = [v for v in vols if v["DEP"] == "NTTB"]
        vols_arrivee = [v for v in vols if v["ARR"] == "NTTB"]
        self.label_status.setText(f"Vols départ : {len(vols_depart)}, Vols arrivée : {len(vols_arrivee)}")

        if vols_depart or vols_arrivee:
            if vols_depart[0]['FLT'][:2] == 'VT' :
                operateur = 'Air Tahiti'
            elif vols_depart[0]['FLT'][:2] == 'NT' :
                operateur = 'Air Moana'
            date_str = date_locale.strftime("%Y%m%d")
            aerodrome = 'NTTB'
            nom_pdf = f"strips_{aerodrome}_{operateur}_{date_str}.pdf"
            self.creer_strips_pdf(vols_depart, vols_arrivee, date_locale,output_pdf= nom_pdf)
            
            self.label_status.setText("✅ Strips générés : strips_nttb_final.pdf")
        else:
            self.label_status.setText("⚠️ Aucun vol trouvé pour NTTB.")
        

    # === Fonctions utilitaires ===
    



    def extraire_date(self, texte):
        match = re.search(r"\b(\d{2})/(\d{2})/(\d{4})\b", texte)
        if match:
            return int(match.group(1)), int(match.group(2)), int(match.group(3))
        return 0, 0, 0

    def extraire_vols(self, texte):
        lignes = [l.rstrip() for l in texte.splitlines()]
        vols = []
        re_time = re.compile(r"^\d{2}:\d{2}$")
        re_oaci = re.compile(r"^NT[A-Z]{2}$")
        re_immat = re.compile(r"^F-?[A-Z0-9]{4}$")
        re_vt_num = re.compile(r"^VT\d+$")
        re_vt = re.compile(r"^VT$|^VT\d+")
        re_ntr = re.compile(r"^NTR\d+")
        jour, mois, annee = self.extraire_date(texte)


        i = 0
        while i < len(lignes):
            L = (lignes[i] or "").strip()

            # === Air Moana ===
            if re_ntr.match(L):
                start = max(0, i - 1)
                end = min(len(lignes), i + 6)
                block = [(lignes[k] or "").strip() for k in range(start, end)]
                while len(block) < 7:
                    block.append("—")

                immat = block[0]
                if immat.startswith("F") and not immat.startswith("F-"):
                    immat = immat[:1] + "-" + immat[1:]

                vol = {
                    "IMMAT": immat,
                    "FLT": block[1],
                    "DEP": block[2],
                    "STD": block[3],
                    "STA": block[4],
                    "ARR": block[5],
                    "ALT": block[6],
                    "TYPE": "AT76"
                }
                vols.append(vol)
                i = end
                continue

            # === Air Tahiti ===
            if (L.startswith("VT") and (i+1)<len(lignes) and (lignes[i+1] or "").strip().isdigit()) or re_vt_num.match(L):
                # reconstruct flight number
                if re_vt_num.match(L):
                    flt = L
                    start_idx = i
                else:
                    flt = L + (lignes[i+1] or "").strip()
                    start_idx = i

                # collect forward until next VT start or up to a safe limit
                j = start_idx
                block = []
                while j < len(lignes):
                    cand = (lignes[j] or "").strip()
                    # break on next VT start (but not current)
                    if j != start_idx and ( (cand.startswith("VT") and (j+1)<len(lignes) and (lignes[j+1] or "").strip().isdigit()) or re_vt_num.match(cand) ):
                        break
                    block.append(cand)
                    j += 1

                # now analyze block by extracting tokens of interest (in order)
                # tokens = sequence of non-empty meaningful tokens
                tokens = [t for t in block if t != "" and t != None]

                # find immat: first token that matches immat pattern (skip VT parts)
                immat = "—"
                for t in tokens:
                    t = t.strip()
                    if re_immat.match(t) and not re_vt_num.match(t) and t != flt:
                        immat = t
                        break
                    # ---- Extraction du type d'appareil ----
                type_appareil = None
                for t in tokens:
                    if re.match(r"AT\d{2}", t):
                        type_appareil = t
                        break

                # find all times and OACI codes (in order)
                times = [t for t in tokens if re_time.match(t)]
                oacis = [t for t in tokens if re_oaci.match(t)]


                # Heuristique assignation:
                # - If there are at least two OACI codes, assume DEP = first, ARR = second (or more)
                # - If only one OACI, try to infer from surrounding context: check block order
                if len(oacis) >= 2:
                    dep = oacis[0]
                    arr = oacis[1]
                elif len(oacis) == 1:
                    # try to locate the single OACI position in block to decide if it's DEP or ARR.
                    idx_oaci = None
                    for k,entry in enumerate(block):
                        if entry == oacis[0]:
                            idx_oaci = k
                            break
                    # if OACI appears late in block -> consider ARR, else DEP
                    if idx_oaci is not None and idx_oaci > len(block)//2:
                        arr = oacis[0]
                        dep = "—"
                    else:
                        dep = oacis[0]
                        arr = "—"
                else:
                    dep = "—"
                    arr = "—"
                    

                # times assignment: first time = STD, second = STA (if present)
                std = times[0] if len(times) >= 1 else "—"
                sta = times[1] if len(times) >= 2 else "—"

        


                

                vol = {
                    "IMMAT": immat,
                    "FLT": flt,
                    "DEP": dep,
                    "STD": std,
                    "ARR": arr,
                    "STA": sta,
                    "ALT": "—",
                    "TYPE" : type_appareil
                }

                vols.append(vol)
                i = j
                continue
            i += 1
        return vols



    def creer_strips_pdf(self,vols_depart, vols_arrivee,date_locale, output_pdf):
    


        strip_width = 26 * cm
        strip_height = 2.9 * cm
        bande_width = 0.5 * cm
        margin_h = 0 * cm
        margin_v = 0 * cm
        page_width, page_height = landscape(A4)

        strips_per_row = int(page_width // (strip_width + margin_h))
        strips_per_col = int(page_height // (strip_height + margin_v))
        total_per_page = strips_per_row * strips_per_col
        margin_x = (page_width - strips_per_row * (strip_width + margin_h) + margin_h) / 2
        margin_y = (page_height - strips_per_col * (strip_height + margin_v) + margin_v) / 2

        c = canvas.Canvas(output_pdf, pagesize=landscape(A4))

        def calcul_date_utc(pdf_date, heure_str, sens_vol):
            """
            pdf_date : datetime.date ou (jour, mois, annee) tuple
            heure_str : HH:MM UTC depuis le PDF
            sens_vol : "DEP" ou "ARR"
            Retourne : jour_utc, mois_utc, annee_utc
            """
            if isinstance(pdf_date, tuple):
                dt_local = datetime(pdf_date[2], pdf_date[1], pdf_date[0])
            else:
                dt_local = datetime(pdf_date.year, pdf_date.month, pdf_date.day)
            
            try:
                h, m = map(int, heure_str.split(":"))
                heure_vol = datetime(dt_local.year, dt_local.month, dt_local.day, h, m)
            except:
                heure_vol = dt_local.replace(hour=0, minute=0)

            # Logique : vols UTC entre 00:00 et 09:59 → date UTC = jour suivant
            heure_min = datetime(dt_local.year, dt_local.month, dt_local.day, 0, 0)
            heure_max = datetime(dt_local.year, dt_local.month, dt_local.day, 9, 59)

            if heure_min < heure_vol <= heure_max:
                heure_vol += timedelta(days=1)

            return heure_vol.day, heure_vol.month, heure_vol.year


        # Normalisation de la date locale
        date_locale_dt = datetime(date_locale.year, date_locale.month, date_locale.day)





        buffer = 6
        offset = 0
        batcha = vols_arrivee
        batchd = vols_depart
        while offset < len(vols_arrivee) :
            # VERSO - Arrivées
            #for start in range(offset, buffer, 6):           
            batch = batcha[offset:offset+buffer]
            #batch = vols_arrivee
            print(f"total par page{total_per_page}")
            
            for idx, vol in enumerate(batch):
                row = idx // strips_per_row
                col = idx % strips_per_row
                x = margin_x + col * (strip_width + margin_h)
                y = page_height - margin_y - (row + 1) * strip_height - row * margin_v
                
                
                jour_utc, mois_utc, annee_utc = calcul_date_utc(date_locale_dt, vol['STA'], "ARR")
                # Bande verte et contour
                c.setFillColor(green)
                c.rect(x, y, bande_width, strip_height, fill=1)
                c.setStrokeColor(green)
                c.line(x+4*cm,y+0*cm,x+4*cm,y+2.9*cm)
                c.line(x+4*cm,y+1.45*cm,x+26*cm,y+1.45*cm)
                c.line(x+6*cm,y+0*cm,x+6*cm,y+2.9*cm)
                c.line(x+8*cm,y+0*cm,x+8*cm,y+2.9*cm)
                c.line(x+12*cm,y+0*cm,x+12*cm,y+2.9*cm)
                c.line(x+14.5*cm,y+0*cm,x+14.5*cm,y+2.9*cm)
                c.line(x+17*cm,y+0*cm,x+17*cm,y+2.9*cm)
                c.line(x+19.5*cm,y+0*cm,x+19.5*cm,y+2.9*cm)
                c.line(x+22*cm,y+0*cm,x+22*cm,y+2.9*cm)
                c.line(x+24*cm,y+0*cm,x+24*cm,y+2.9*cm)
                c.setStrokeColor(black)
                c.rect(x, y, strip_width, strip_height, fill=0)

                # Texte
                c.setFont("Helvetica-Bold", 10)
                c.setFillColor(black)
                c.drawString(x + 2.8*cm, y + 0.2*cm, vol['DEP'])
                c.drawString(x + 1.5*cm, y + 1.45*cm, vol['FLT'])
                c.drawString(x + 0.6*cm, y + 2.5*cm, vol['IMMAT'])
                c.drawString(x + 9.5*cm, y + 2.55*cm, vol['STA'])
                c.drawString(x + 2.3*cm, y + 2.5*cm, vol['TYPE'])
                c.drawString(x + 2.1*cm, y + 2.5*cm, '/')
                c.drawString(x + 0.055*cm, y + 2.5*cm,f"{jour_utc:02d}")
                c.drawString(x + 0.055*cm, y + 1.5*cm,f"{mois_utc:02d}")
                c.drawString(x + 0.055*cm, y + 0.5*cm,f"{annee_utc % 100:02d}")
            c.showPage()


            # RECTO - Départs
            #for start in range(offset, buffer, 6):
            'batch = vols_depart[start:start+total_per_page]'
            batch = batchd[offset:offset+buffer]
            #batch = vols_depart
            for idx, vol in enumerate(batch):
                row = idx // strips_per_row
                col = idx % strips_per_row
                x = margin_x + col * (strip_width + margin_h)
                y = page_height - margin_y - (row + 1) * strip_height - row * margin_v
                
        
                jour_utc, mois_utc, annee_utc_dep = calcul_date_utc(date_locale_dt, vol['STD'], "DEP")
                # Bande rouge et contour
                c.setFillColor(red)
                c.rect(x, y, bande_width, strip_height, fill=1)
                c.setStrokeColor(red)
                c.line(x+6*cm,y+1.45*cm,x+12*cm,y+1.45*cm)
                c.line(x+16*cm,y+1.45*cm,x+20*cm,y+1.45*cm)
                c.line(x+6*cm,y+0*cm,x+6*cm,y+2.9*cm)
                c.line(x+12*cm,y+0*cm,x+12*cm,y+2.9*cm)
                c.line(x+16*cm,y+0*cm,x+16*cm,y+2.9*cm)
                c.line(x+18*cm,y+1.45*cm,x+18*cm,y+2.9*cm)
                c.line(x+20*cm,y+0*cm,x+20*cm,y+2.9*cm)
                c.line(x+23*cm,y+0*cm,x+23*cm,y+2.9*cm)
                c.setStrokeColor(black)
                c.rect(x, y, strip_width, strip_height, fill=0)

                # Texte
                c.setFont("Helvetica-Bold", 10)
                c.setFillColor(black)
                c.drawString(x + 0.055*cm, y + 2.5*cm,f"{jour_utc:02d}")
                c.drawString(x + 0.055*cm, y + 1.5*cm,f"{mois_utc:02d}")
                c.drawString(x + 0.055*cm, y + 0.5*cm,f"{annee_utc % 100:02d}")
                if vol['ARR'] != 'NTTR' :
                    c.drawString(x+ 21*cm, y + 1.7*cm,f"134,7")

                c.drawString(x + 4.9*cm, y + 0.2*cm, vol['ARR'])
                c.drawString(x + 1.5*cm, y + 1.45*cm, vol['FLT'])
                c.drawString(x + 0.6*cm, y + 2.50*cm, vol['IMMAT'])
                c.drawString(x + 2.1*cm, y + 2.50*cm, '/')
                c.drawString(x + 2.3*cm, y + 2.50*cm, vol['TYPE'])
                c.drawString(x + 13.5*cm, y + 1.45*cm, vol['STD'])
            c.showPage()
            
            offset = offset + buffer


        c.save()
        print(f"PDF généré : {output_pdf}")

# -------------------------------------------------------------------
# Widget pour la conversion (anciennement MainWindow transformé en widget)
# -------------------------------------------------------------------
class DropLabel(QLabel):
    """
    Zone de dépôt "drag & drop" pour le fichier PDF.
    Met à jour le QLineEdit associé lors du dépôt.
    """
    def __init__(self, line_edit, parent=None):
        super().__init__(parent)
        self.line_edit = line_edit
        self.setText("\n\nDéposez ici votre fichier PDF\n\n")
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("QLabel { border: 4px dashed #aaa; font-size: 16px; }")
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(".pdf"):
                self.line_edit.setText(file_path)
                self.setText(f"{os.path.basename(file_path)} déposé.")
            else:
                self.setText("Fichier non-PDF. Veuillez déposer un fichier PDF.")

class BaseWidget(QWidget):
    """
    Classe de base pour homogénéiser les styles et l'apparence de toutes nos interfaces.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_common_style()

    def setup_common_style(self):
        # Style commun appliqué à tous les widgets hérités.
        self.setStyleSheet("""
            QWidget {
                background-color: #FFFFFF;
                font-family: Arial, sans-serif;
                font-size: 13px;
            }
            QLabel {
                color: #333333;
                font-size: 14px;
                margin: 4px;
            }
            QLineEdit {
                border: 1px solid #CCCCCC;
                padding: 5px;
                border-radius: 3px;
                margin: 4px;
            }
            QPushButton {
                background-color: #007ACC;
                color: #FFFFFF;
                padding: 8px;
                border: none;
                border-radius: 3px;
                margin: 4px;
            }
            QPushButton:hover {
                background-color: #005F9E;
            }
            QCheckBox {
                margin: 4px;
            }
        """)


class ConversionWidget(BaseWidget):
    """
    Widget encapsulant l'interface graphique de conversion PDF → csv.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)

        # Message d'information
        info_label = QLabel("Sélectionnez le programme Air Moana pour le convertir en programme pour statADD")
        layout.addWidget(info_label)

        # Affichage du chemin du fichier
        self.lineEdit = QLineEdit()
        self.lineEdit.setReadOnly(True)
        layout.addWidget(self.lineEdit)

        # Bouton "Parcourir..."
        browseButton = QPushButton("Parcourir...")
        browseButton.clicked.connect(self.open_file)
        layout.addWidget(browseButton)

        # Zone de dépôt (supposée personnalisée via DropLabel)
        self.dropLabel = DropLabel(self.lineEdit, self)
        layout.addWidget(self.dropLabel)

        # Bouton de conversion
        convertButton = QPushButton("Convertir")
        convertButton.clicked.connect(self.on_convert_clicked)
        layout.addWidget(convertButton)

    def open_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionnez un fichier PDF",
            "",
            "Fichiers PDF (*.pdf);;Tous les fichiers (*)",
            options=options,
        )
        if file_path:
            self.lineEdit.setText(file_path)
            self.dropLabel.setText(f"{os.path.basename(file_path)} sélectionné.")

    def convert_file(self):
        pdf_path = self.lineEdit.text()
        if not pdf_path or not pdf_path.lower().endswith(".pdf"):
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un fichier PDF valide.")
            return None, None
        try:
            csv_file = convert_pdf_to_csv(pdf_path)
            QMessageBox.information(self, "Succès", f"La conversion a réussi !\nFichier créé :\n{csv_file}")
            return csv_file
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Une erreur est survenue :\n{str(e)}")
            return None, None
    def envoyer_email(self, csv_file_path, pdf_date, prgrm_name):
        try:
            # Création d'une instance Outlook
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)  # 0 correspond à un mail

            # Définir le destinataire
            mail.To = "seac-pf-nttb-bf@aviation-civile.gouv.fr"

            # Récupérer la date actuelle au format souhaité
            date_str = datetime.now().strftime("%Y-%m-%d")

            # Préparer l'objet et le corps du mail avec la date
            mail.Subject = (
                f"Programme des vols {prgrm_name}"
            )
            mail.Body = (
                f"Vous trouverez ci-joint le programme des vols {prgrm_name} du {pdf_date}.\n\n"
                "Cordialement."
            )

            # Attacher le fichier CSV
            if os.path.exists(csv_file_path):
                mail.Attachments.Add(csv_file_path)
            else:
                QMessageBox.warning(self, "Erreur", "Le fichier CSV n'existe pas. Email non envoyé.")
                return

            # Pour tester, vous pouvez utiliser mail.Display() pour voir le mail dans Outlook
            mail.Display()
            print("Email envoyé avec succès.")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Une erreur est survenue lors de l'envoi de l'email :\n{e}")
            print(f"Erreur : {e}")
        
    def on_convert_clicked(self):
        # Appel de la méthode de conversion
        csv_file_path, pdf_date, prgrm_name  = self.convert_file()
        if csv_file_path:
            self.envoyer_email(csv_file_path, pdf_date, prgrm_name)


class CalendrierWidget(BaseWidget):
    """
    Interface de génération d'ICS depuis un planning PDF.
    Permet de sélectionner le PDF, de lire la date de début et de choisir les agents.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.detected_agents = set()
        self.agent_checkboxes = {}
        self.agents_par_jour = {}
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # Partie sélection du fichier PDF
        file_layout = QHBoxLayout()
        self.lineEdit_pdf = QLineEdit()
        self.lineEdit_pdf.setReadOnly(True)
        file_layout.addWidget(self.lineEdit_pdf)
        self.button_select_file = QPushButton("Sélectionner le TDS au format PDF")
        self.button_select_file.clicked.connect(self.select_file)
        file_layout.addWidget(self.button_select_file)
        layout.addLayout(file_layout)

        # Partie saisie de la date de début
        date_layout = QHBoxLayout()
        date_label = QLabel("Date de début (YYYY-MM-DD) :")
        date_layout.addWidget(date_label)
        self.lineEdit_start_date = QLineEdit()
        date_layout.addWidget(self.lineEdit_start_date)
        layout.addLayout(date_layout)

        # Liste des agents détectés
        agents_label = QLabel("Sélectionnez les agents :")
        layout.addWidget(agents_label)

        self.agents_widget = QWidget()
        self.agents_layout = QGridLayout(self.agents_widget)
        layout.addWidget(self.agents_widget)
        
        # Champ d'affichage complémentaire (utilisé avec le drop)
        self.lineEdit = QLineEdit()
        self.lineEdit.setReadOnly(True)
        layout.addWidget(self.lineEdit)

        # Zone de dépôt
        self.dropLabel = DropLabel(self.lineEdit, self)
        layout.addWidget(self.dropLabel)
        
        # Bouton pour lancer la génération ICS
        self.button_generate_ics = QPushButton("Générer ICS")
        self.button_generate_ics.clicked.connect(self.start_conversion)
        layout.addWidget(self.button_generate_ics)

        layout.addStretch()

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionnez un fichier PDF",
            "",
            "Fichiers PDF (*.pdf);;Tous les fichiers (*)"
        )
        if file_path:
            self.lineEdit_pdf.setText(file_path)
            # Extraction de la date à partir du nom du fichier
            start_date = self.extract_date_from_filename(os.path.basename(file_path))
            if start_date:
                self.lineEdit_start_date.setText(start_date)
            self.update_agents_list(file_path,start_date)

    def extract_date_from_filename(self, filename):
        """
        Extrait la date de début à partir du nom du fichier PDF.
        Pattern : ^(YYYY).(MM).*du (DD)(MM)
        Exemple : "2025.01...du 0601..." donnera "2025-01-06"
        """
        match = re.search(r"^(\d{4})\.(\d{2}).*du (\d{2})(\d{2})", filename)
        if match:
            year = match.group(1)
            month = match.group(2)
            day = match.group(3)
            file_month = match.group(4)
            if file_month != month:
                print("Avertissement : le mois du début du fichier et celui du jour ne correspondent pas.")
            try:
                start_date = datetime.strptime(f"{year}-{file_month}-{day}", "%Y-%m-%d")
                return start_date.strftime("%Y-%m-%d")
            except Exception as e:
                print(e)
        return None

    def update_agents_list(self, pdf_path):
        self.agents_par_jour = extract_shifts_from_pdf(pdf_path,start_date)
        detected_agents = set()
        for jour, (j1, j2) in self.agents_par_jour.items():
            detected_agents.add(j1)
            detected_agents.add(j2)
        if "NONE" in detected_agents:
            detected_agents.remove("NONE")
        self.detected_agents = detected_agents
        self.update_agents_ui()

    def update_agents_ui(self):
        # On efface les anciens widgets (pour la mise à jour de la liste)
        for i in reversed(range(self.agents_layout.count())):
            widget = self.agents_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        self.agent_checkboxes = {}
        row, col = 0, 0
        for agent in sorted(self.detected_agents):
            checkbox = QCheckBox(agent)
            self.agents_layout.addWidget(checkbox, row, col)
            self.agent_checkboxes[agent] = checkbox
            col += 1
            if col >= 3:
                col = 0
                row += 1

    def start_conversion(self):
        pdf_path = self.lineEdit_pdf.text()
        start_date_str = self.lineEdit_start_date.text()
        selected_agents = [agent for agent, cb in self.agent_checkboxes.items() if cb.isChecked()]

        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.critical(self, "Erreur", "Veuillez sélectionner un fichier PDF valide.")
            return

        if not selected_agents:
            QMessageBox.critical(self, "Erreur", "Veuillez sélectionner au moins un agent.")
            return

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        except ValueError:
            QMessageBox.critical(self, "Erreur", "Format de date invalide. Utilisez YYYY-MM-DD.")
            return

        agents_par_jour = extract_shifts_from_pdf(pdf_path,start_date)
        shifts_by_agent = {agent: [] for agent in selected_agents}

        # On aligne la date des shifts avec celle indiquée
        premier_jour = min(agents_par_jour.keys())
        jours_tries = sorted(agents_par_jour.keys())
        dernier_jour = max(agents_par_jour.keys())

        for jour in jours_tries:
            if jour > dernier_jour:
                break

            jour_date = start_date + timedelta(days=(jour - premier_jour))
            # Définition des deux shifts par jour
            j1_start = jour_date.replace(hour=6, minute=45)
            j1_end   = jour_date.replace(hour=13, minute=30)
            j2_start = jour_date.replace(hour=13, minute=45)
            j2_end   = jour_date.replace(hour=20, minute=45)

            j1, j2 = self.agents_par_jour[jour]
            if j1 in shifts_by_agent:
                shifts_by_agent[j1].append((j1_start, j1_end))
            if j2 in shifts_by_agent:
                shifts_by_agent[j2].append((j2_start, j2_end))

        output_dir = os.path.dirname(pdf_path)
        generate_ics(shifts_by_agent, output_dir)
        QMessageBox.information(self, "Succès", "Fichiers ICS générés avec succès !")
        
# -------------------------------------------------------------------
# MainWindow qui intègre les trois applications dans une interface à onglets
# -------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outil NTTB")
        self.resize(800, 600)

        # Création d'un QTabWidget pour intégrer plusieurs applications
        tabs = QTabWidget()
        self.setCentralWidget(tabs)
        
        # Onglet 1 : Interface de conversion PDF -> csv
        conversion_tab = ConversionWidget()
        tabs.addTab(conversion_tab, "StatAdd")
        
        #Onglet 2 : interface stripping
        Strip_tab = StrippingWidget()
        tabs.addTab(Strip_tab, "Strips")
        
        # Onglet 3 : conversion TDS to ICS
        Calendrier_tab = CalendrierWidget()
        tabs.addTab(Calendrier_tab, "Calendrier")

# -------------------------------------------------------------------
# Lancement de l'application
# -------------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
