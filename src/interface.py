import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from ttkthemes import ThemedStyle
from datetime import datetime
from tkinter import *
from tkinter import filedialog, messagebox
from scraper.ZLME_scrap import ZLME_scrap
from scraper.EURX_scrap import EURX_scrap
import configparser
import os
from excel.excel_manager import write_to_excel
from tkcalendar import DateEntry
import subprocess

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Price € / $")
        self.geometry("1000x600")
        self.resizable(False, False)
        
        self.style = ThemedStyle(self)
        self.style.set_theme("arc")  # Choisissez le thème que vous préférez
        
        self.create_widgets()

    def create_widgets(self):
        
        # Charger le chemin d'accès depuis le fichier de configuration
        self.config = configparser.ConfigParser()
        self.config.read('./config.ini')
        self.excel_path = self.config.get('DEFAULT', 'ExcelPath', fallback="")
        print("chemin", self.excel_path)
        print(self.config.read('./config.ini'))
        
         # Configure la fenêtre principale pour s'adapter dynamiquement
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=1)  # Donne un poids à la ligne pour qu'elle s'adapte
        
        # Création des cadres principaux pour la division gauche et droite
        left_frame = ttk.Frame(self)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        right_frame = ttk.Frame(self)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        # Assurez-vous que les cadres s'adaptent également
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        
        right_frame.grid_rowconfigure(0, weight=1)  # Permet au contenu de s'étendre verticalement
        for i in range(5):  # S'assurer que tous les éléments de la partie droite peuvent s'étendre
            right_frame.grid_rowconfigure(i, weight=1)
        
        self.create_left_side(left_frame)
        self.create_right_side(right_frame)

    def create_left_side(self, frame):
        # Configuration pour les logs dans la partie gauche
        self.log_text = scrolledtext.ScrolledText(frame, font=('Arial', 10), background="white")
        self.log_text.pack(expand=True, fill="both")

    def create_right_side(self, frame):
        
        # Configure les colonnes de frame pour utiliser l'espace disponible
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)
    
        # Titre
        title = ttk.Label(frame, text="Prices € / $", font=('Arial', 12, 'bold'))
        title.grid(row=0, column=1, columnspan=3, sticky="w", padx=0, pady=(10, 20))
        
        # Champ pour le chemin d'accès
        self.path_entry = ttk.Entry(frame, font=('Arial', 10))
        self.path_entry.grid(row=1, column=0, columnspan=3,padx=5, pady=5, sticky="ew")
        self.path_entry.insert(0, self.excel_path)

        # Bouton avec texte "Modifier" pour modifier le chemin
        edit_button = ttk.Button(frame, text="Modifier", command=self.edit_path)
        edit_button.grid(row=2, column=0, padx=5, pady=5)

        # Bouton "Ouvrir"
        open_button = ttk.Button(frame, text="Ouvrir", command=self.open_path)
        open_button.grid(row=2, column=2, padx=5, pady=5)

        # Champs pour les dates - Ajustement des numéros de rangée
        ttk.Label(frame, text="Date de début :").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.date_from_entry = DateEntry(frame, width=17, background='darkblue',
                                         foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_from_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # Ajout du label et du champ de sélection pour la date de fin
        ttk.Label(frame, text="Date de fin :").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.date_to_entry = DateEntry(frame, width=17, background='darkblue',
                                       foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_to_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # Checkbutton pour activer ou non l'option "utiliser date à date"
        self.use_date_range = tk.BooleanVar()
        date_range_check = ttk.Checkbutton(frame, text="Utiliser date à date", variable=self.use_date_range)
        date_range_check.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="w")

        # Bouton principal pour lancer le programme
        launch_button = ttk.Button(frame, text="Lancer le programme", command=self.launch_program)
        launch_button.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="ew")


    def edit_path(self):
        path = filedialog.askopenfilename(title='Sélectionner un fichier Excel', filetypes=(('Excel Files', '*.xlsx *.xls'), ('All Files', '*.*')))
        if path:
            # Afficher le chemin choisi (optionnel)
            print("Chemin sélectionné:", path)
            
            # Sauvegarder le chemin dans config.ini
            config = configparser.ConfigParser()
            config['DEFAULT'] = {'ExcelPath': path}
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
            messagebox.showinfo("Information", "Chemin modifié.")

    def open_path(self):
        try:
            subprocess.run(["start", self.excel_path], shell=True, check=True)
            self.log('Fichier ouvert.')
        except subprocess.CalledProcessError as e:
            self.log('Fichier non trouvé.')
    def launch_program(self):
        # Vérifiez l'état de la checkbox pour déterminer les dates à utiliser
        if self.use_date_range.get():  # Correction pour l'utilisation de get() avec BooleanVar
            start_date = self.date_from_entry.get_date()  # Correction pour récupérer la date
            end_date = self.date_to_entry.get_date()
        else:
            start_date = end_date = datetime.today().date()

        # Scraping ZLME
        result_zlme = ZLME_scrap(start_date, end_date)
        if isinstance(result_zlme, str):
            self.log(f"Erreur ZLME: {result_zlme}")
            return

        # Scraping EURX
        result_eurx = EURX_scrap(start_date, end_date)
        if isinstance(result_eurx, str):
            self.log(result_eurx)
            self.log(f"Erreur EURX: {result_eurx}")
            return

        if not os.path.exists(self.excel_path):
            self.log(f"Fichier Excel non trouvé à {self.excel_path}.")
            return

        # Écriture dans Excel
        try:
            write_to_excel(self.excel_path, result_zlme, result_eurx, self.log)
            self.log("Données enregistrées avec succès.")
        except Exception as e:
            self.log(f"Erreur écriture Excel: {e}")


    def log(self, message):
        if isinstance(message, list):
            message = ' '.join(map(str, message))
        # Affichez le message dans la console ou dans un widget de log dans votre interface
        print(message)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)


if __name__ == "__main__":
    app = Application()
    app.mainloop()
