import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docx import Document
import os
from typing import Optional
import openpyxl
import docx
import sys
import os
import numpy
from PIL import Image, ImageTk

# === Fonctions pour les documents Word ===
def generate_convention(data_row):
    doc = Document()
    doc.add_heading('Convention de Stage', 0)
    doc.add_paragraph(f"Nom: {data_row['Nom']}")
    doc.add_paragraph(f"Formation: {data_row['Formation']}")
    doc.add_paragraph(f"Date: {data_row['Date']}")
    doc.save('Convention_Stage.docx')


def generate_emargement(filtered_data):
    doc = Document()
    doc.add_heading('Feuille d\'Emargement', 0)
    table = doc.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Nom'
    hdr_cells[1].text = 'Formation'
    hdr_cells[2].text = 'Signature'

    for index, row in filtered_data.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Nom'])
        row_cells[1].text = str(row['Formation'])
        row_cells[2].text = ''

    doc.save('Feuille_Emargement.docx')


def generate_chevalets(filtered_data):
    for index, row in filtered_data.iterrows():
        doc = Document()
        doc.add_heading(f"Chevalet: {row['Nom']}", 0)
        doc.add_paragraph(f"Formation: {row['Formation']}")
        doc.save(f"Chevalet_{row['Nom']}.docx")


def export_filtered_excel(filtered_data, formation_name="", date_filter=""):
    """Exporte les données filtrées vers un fichier Excel pour publipostage"""
    if filtered_data.empty:
        return False
    
    # Créer un nom de fichier basé sur les filtres
    filename = f"Donnees_Filtrees"
    if formation_name:
        filename += f"_{formation_name}"
    if date_filter:
        filename += f"_{date_filter}"
    filename += ".xlsx"
    
    # Exporter vers Excel
    try:
        filtered_data.to_excel(filename, index=False)
        return filename
    except Exception as e:
        print(f"Erreur lors de l'export Excel: {e}")
        return False


# === Interface Tkinter ===
class Application(tk.Tk):
    def open_convention_file(self):
        """Ouvre un fichier Word avec l'application par défaut dans le dossier Datas/documents (toujours accessible)"""
        initial_dir = resource_path(os.path.join("Datas", "documents"))
        file_path = filedialog.askopenfilename(
            filetypes=[("Word Files", "*.docx")],
            title="Ouvrir un fichier Word",
            initialdir=initial_dir
        )
        if file_path:
            os.startfile(file_path)
    def open_word_file(self):
        """Ouvre un fichier Word avec l'application par défaut"""
        file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")], title="Ouvrir un fichier Word")
        if file_path:
            import os
            os.startfile(file_path)

    def filter_and_export_excel(self):
        """Filtre les données et exporte vers un fichier Excel pour publipostage"""
        filtered_df = self.filter_data()
        if filtered_df is None:
            return

        # Demander où sauvegarder le fichier
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Sauvegarder le fichier Excel filtré"
        )
        
        if save_path:
            try:
                filtered_df.to_excel(save_path, index=False)
                messagebox.showinfo("Succès", f"Fichier Excel exporté avec succès!\n{len(filtered_df)} lignes exportées.\n\nVous pouvez maintenant utiliser ce fichier pour le publipostage dans Word.")
                
                # Ouvrir le dossier contenant le fichier
                import os
                os.startfile(os.path.dirname(save_path))
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")

    def __init__(self):
        super().__init__()
        self.iconbitmap("logo-Toyota-Solo.ico")
        self.title("Rev-20250812-01")
        self.minsize(700, 400)  # Augmenté la taille minimale pour accommoder l'image

        self.file_path = None
        self.df = None

        # Créer le layout principal avec deux frames
        main_frame = tk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Frame de gauche pour l'image
        left_frame = tk.Frame(main_frame, width=200)
        left_frame.pack(side="left", fill="y", padx=(0, 20))
        left_frame.pack_propagate(False)  # Empêche le frame de se redimensionner

        # Frame de droite pour les widgets
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True)

        # Ajouter l'image sur le côté gauche
        try:
            # Charger et redimensionner l'image
            image_path = "LogoTMH.png"
            if os.path.exists(image_path):
                # Charger l'image PNG avec PIL et la redimensionner
                pil_image = Image.open(image_path)
                # Redimensionner l'image pour qu'elle tienne dans le frame (max 150px de large)
                pil_image.thumbnail((150, 150), Image.Resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(pil_image)
                logo_label = tk.Label(left_frame, image=self.logo_image)
                logo_label.pack(pady=20)
            else:
                # Fallback si l'image n'existe pas
                logo_label = tk.Label(left_frame, text="Logo TMH", font=("Arial", 16, "bold"))
                logo_label.pack(pady=20)
        except Exception as e:
            # En cas d'erreur, afficher un texte à la place
            logo_label = tk.Label(left_frame, text="Logo TMH", font=("Arial", 16, "bold"))
            logo_label.pack(pady=20)

        # Widgets dans le frame de droite
        # Titre de l'application
        title_label = tk.Label(right_frame, text="Générateur de Documents de Formation", 
                              font=("Arial", 14, "bold"), justify="center")
        title_label.pack(pady=10)
        
        tk.Button(right_frame, text="Charger Fichier Excel", command=self.load_excel).pack(pady=10)
        
        tk.Label(right_frame, text="Date (AAAA-MM-JJ):").pack()
        self.date_entry = tk.Entry(right_frame)
        self.date_entry.pack()

        tk.Label(right_frame, text="Type de Formation:").pack()
        self.formation_entry = tk.Entry(right_frame)
        self.formation_entry.pack()

        # Remettre le bouton d'export Excel
        tk.Button(right_frame, text="Filtrer et Exporter Excel", command=self.filter_and_export_excel).pack(pady=10)
        
        tk.Button(right_frame, text="Afficher Toutes les Données", command=self.show_all_data).pack(pady=5)
        
        tk.Button(right_frame, text="Ouvrir feuille d'emargement", command=self.open_word_file).pack(pady=10)

        tk.Button(right_frame, text="Ouvrir convention", command=self.open_convention_file).pack(pady=10)

    def show_all_data(self):
        """Affiche toutes les données du fichier Excel"""
        if self.df is None:
            messagebox.showerror("Erreur", "Veuillez d'abord charger un fichier Excel.")
            return
        
        # Créer une fenêtre pour afficher les données
        data_window = tk.Toplevel(self)
        data_window.title("Toutes les Données")
        data_window.geometry("800x600")
        
        # Créer un widget Text pour afficher les données
        text_widget = tk.Text(data_window, wrap=tk.WORD)
        scrollbar = tk.Scrollbar(data_window, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Afficher les données
        text_widget.insert(tk.END, f"Total des lignes: {len(self.df)}\n\n")
        text_widget.insert(tk.END, f"Colonnes: {list(self.df.columns)}\n\n")
        
        # Afficher les premières 20 lignes
        text_widget.insert(tk.END, "Premières 20 lignes:\n")
        text_widget.insert(tk.END, str(self.df.head(20)))
        
        # Afficher les valeurs uniques pour les colonnes importantes
        if 'datedebutsession' in self.df.columns:
            dates = self.df['datedebutsession'].dropna().unique()
            text_widget.insert(tk.END, f"\n\nDates uniques ({len(dates)}):\n")
            for date in dates[:10]:  # Afficher les 10 premières
                text_widget.insert(tk.END, f"- {date}\n")
            if len(dates) > 10:
                text_widget.insert(tk.END, f"... et {len(dates) - 10} autres\n")
        
        if 'course full name' in self.df.columns:
            formations = self.df['course full name'].dropna().unique()
            text_widget.insert(tk.END, f"\nFormations uniques ({len(formations)}):\n")
            for formation in formations[:10]:  # Afficher les 10 premières
                text_widget.insert(tk.END, f"- {formation}\n")
            if len(formations) > 10:
                text_widget.insert(tk.END, f"... et {len(formations) - 10} autres\n")

    def load_excel(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.file_path:
            self.df = pd.read_excel(self.file_path)
            # Afficher les informations sur les colonnes de filtrage
            info_text = "Fichier chargé avec succès!\n\n"
            
            if 'datedebutsession' in self.df.columns:
                dates_uniques = self.df['datedebutsession'].dropna().unique()
                info_text += f"Dates disponibles: {', '.join(map(str, dates_uniques[:5]))}\n"
                if len(dates_uniques) > 5:
                    info_text += f"... et {len(dates_uniques) - 5} autres dates\n"
            
            if 'course full name' in self.df.columns:
                formations_uniques = self.df['course full name'].dropna().unique()
                info_text += f"\nFormations disponibles: {', '.join(map(str, formations_uniques[:3]))}\n"
                if len(formations_uniques) > 3:
                    info_text += f"... et {len(formations_uniques) - 3} autres formations"
            
            messagebox.showinfo("Succès", info_text)

    def filter_data(self) -> Optional[pd.DataFrame]:
        """Filtre les données selon les critères saisis"""
        if self.df is None:
            messagebox.showerror("Erreur", "Veuillez d'abord charger un fichier Excel.")
            return None

        date_filter = self.date_entry.get()
        formation_filter = self.formation_entry.get()

        filtered_df: pd.DataFrame = self.df.copy()
        
        # Debug: afficher le nombre de lignes initial
        initial_count = len(filtered_df)
        debug_info = f"Données initiales: {initial_count} lignes\n"

        if date_filter:
            # Filtrage plus flexible pour la date
            filtered_df = filtered_df[filtered_df['datedebutsession'].astype(str).str.contains(date_filter, na=False)]
            debug_info += f"Après filtre date '{date_filter}': {len(filtered_df)} lignes\n"
            
        if formation_filter:
            # Filtrage très flexible pour la formation (insensible à la casse et par mots-clés)
            formation_filter_lower = formation_filter.lower()
            
            # Recherche par mots-clés séparés
            mots_cles = formation_filter_lower.split()
            if len(mots_cles) > 1:
                # Si plusieurs mots-clés, chercher ceux qui contiennent TOUS les mots
                mask = pd.Series([True] * len(filtered_df), index=filtered_df.index)
                for mot in mots_cles:
                    mask = mask & filtered_df['course full name'].astype(str).str.lower().str.contains(mot, na=False)
                filtered_df = filtered_df[mask]
                debug_info += f"Après filtre formation '{formation_filter}' (recherche par mots-clés): {len(filtered_df)} lignes\n"
            else:
                # Si un seul mot, recherche simple
                filtered_df = filtered_df[filtered_df['course full name'].astype(str).str.lower().str.contains(formation_filter_lower, na=False)]
                debug_info += f"Après filtre formation '{formation_filter}' (recherche insensible à la casse): {len(filtered_df)} lignes\n"
        
        # Si aucun filtre n'est appliqué, utiliser toutes les données
        if not date_filter and not formation_filter:
            debug_info += "Aucun filtre appliqué - utilisation de toutes les données\n"

        if len(filtered_df) == 0:
            messagebox.showinfo("Aucun Résultat", f"Aucune donnée ne correspond aux critères.\n\n{debug_info}")
            return None
            
        return filtered_df

    def filter_and_generate(self):
        """Filtre les données et génère les documents Word"""
        filtered_df = self.filter_data()
        if filtered_df is None:
            return

        generate_emargement(filtered_df)
        generate_chevalets(filtered_df)

        # Générer convention pour la première ligne filtrée
        if len(filtered_df) > 0:
            generate_convention(filtered_df.iloc[0])

        messagebox.showinfo("Succès", "Documents Word générés avec succès!")


def resource_path(relative_path):
    """Obtenir le chemin absolu vers une ressource, compatible dev et .exe PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


if __name__ == "__main__":
    app = Application()
    app.mainloop()
