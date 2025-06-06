import tkinter as tk
from tkinter import ttk, PhotoImage
from PIL import Image, ImageTk
from database.db_operations import DBOperations
from scraping.spacex_scraper import SpaceXScraper
from utils.utils import show_error, show_info, show_warning, ask_confirmation, create_matplotlib_graph
import os

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("SpaceX Launches Tracker")
        self.db = DBOperations()
        self.scraper = SpaceXScraper()

        # Palette de couleurs
        self.bg_color = "#FFFFFF"  # Blanc
        self.fg_color = "#0C0C0C"  # Noir
        self.accent_color = "#00BFFF"  # Bleu cyan
        self.header_color = "#3E5C76"  # Bleu acier
        self.border_color = "#B0B3B8"  # Gris métal

        self.setup_ui()

        # Vérifier et charger les données au lancement
        if self.db.get_launches_count() > 0:
            self.update_status("Données chargées depuis la base")
            self.display_latest_data()
        else:
            self.update_status("Système prêt - Aucune donnée en base")

    def setup_ui(self):
        # Configuration de la fenêtre
        self.root.configure(bg=self.bg_color)
        self.root.geometry("1100x700")
        self.root.minsize(800, 600)  # Taille minimale

        # Configuration du grid principal
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Configuration du logo en haut à gauche
        self.setup_header()

        # Zone principale
        self.main_frame = ttk.Frame(self.root, padding=(20, 10))
        self.main_frame.grid(row=1, column=0, sticky="nsew")
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Configuration des styles
        self.setup_styles()

        # Contrôles principaux
        self.setup_controls()

        # Zone de données
        self.setup_data_display()

        # Barre de statut
        self.setup_status_bar()

    def setup_header(self):
        """Configure le header avec logo en haut à gauche"""
        header_frame = ttk.Frame(self.root, style='Header.TFrame')
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header_frame.grid_columnconfigure(0, weight=1)

        # Conteneur pour le logo et le titre
        header_content = ttk.Frame(header_frame, style='Header.TFrame')
        header_content.grid(row=0, column=0, sticky="w", padx=20, pady=10)

        # Logo SpaceX
        try:
            logo_img = Image.open("assets/spacex_logo.png")
            logo_img = logo_img.resize((300, 60), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)

            logo_label = ttk.Label(
                header_frame,
                image=self.logo,
                background=self.header_color
            )
            logo_label.grid(row=0, column=0, padx=20, pady=10, sticky="ne")

            # Titre à droite du logo
            ttk.Label(
                header_content,
                text="SYSTÈME DE SUIVI DES LANCEMENTS",
                style='Header.TLabel'
            ).grid(row=0, column=1, padx=(20, 0), sticky="w")

        except Exception as e:
            print(f"Erreur de chargement du logo: {e}")
            # Fallback textuel
            ttk.Label(
                header_frame,
                text="SPACEX LAUNCH TRACKER",
                style='Header.TLabel'
            ).grid(row=0, column=0, padx=20, sticky="w")

    def setup_styles(self):
        """Configure tous les styles visuels"""
        style = ttk.Style()

        # Style global
        style.theme_use('clam')

        # Header
        style.configure('Header.TFrame',
                      background=self.header_color)
        style.configure('Header.TLabel',
                      background=self.header_color,
                      foreground=self.bg_color,
                      font=('Segoe UI', 12, 'bold'))

        # Boutons
        style.configure('Control.TButton',
                      background=self.accent_color,
                      foreground=self.bg_color,
                      font=('Segoe UI', 10, 'bold'),
                      borderwidth=0,
                      padding=8)
        style.map('Control.TButton',
                background=[('active', self.header_color)])

        # Treeview
        style.configure('Data.Treeview',
                      background=self.bg_color,
                      foreground=self.fg_color,
                      fieldbackground=self.bg_color,
                      rowheight=28,
                      font=('Segoe UI', 9))
        style.configure('Data.Treeview.Heading',
                      background=self.header_color,
                      foreground=self.bg_color,
                      font=('Segoe UI', 10, 'bold'),
                      relief='flat')
        style.map('Data.Treeview',
                background=[('selected', self.border_color)])

        # Barre de statut
        style.configure('Status.TLabel',
                      background=self.border_color,
                      foreground=self.fg_color,
                      font=('Consolas', 9),
                      padding=5)

    def setup_controls(self):
        """Configure les boutons de contrôle"""
        control_frame = ttk.Frame(self.main_frame)
        control_frame.grid(row=0, column=0, pady=(0, 15), sticky="ew")

        ttk.Button(
            control_frame,
            text="ACTUALISER LES DONNÉES",
            command=self.fetch_data,
            style='Control.TButton'
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            control_frame,
            text="EFFACER LA BASE",
            command=self.clear_database,
            style='Control.TButton'
        ).grid(row=0, column=1, padx=5)

        ttk.Button(
            control_frame,
            text="GÉNÉRER STATISTIQUES",
            command=self.show_stats,
            style='Control.TButton'
        ).grid(row=0, column=2, padx=5)

    def setup_data_display(self):
        """Configure l'affichage des données"""
        data_frame = ttk.LabelFrame(
            self.main_frame,
            text=" HISTORIQUE DES LANCEMENTS ",
            padding=15
        )
        data_frame.grid(row=1, column=0, sticky="nsew")
        data_frame.grid_rowconfigure(0, weight=1)
        data_frame.grid_columnconfigure(0, weight=1)

        # Treeview
        self.tree = ttk.Treeview(
            data_frame,
            columns=("mission", "date", "url"),
            show="headings",
            style='Data.Treeview'
        )

        # Configuration des colonnes
        self.tree.heading("mission", text="MISSION", anchor="w")
        self.tree.heading("date", text="DATE", anchor="center")
        self.tree.heading("url", text="LIEN", anchor="w")

        self.tree.column("mission", width=300, minwidth=200)
        self.tree.column("date", width=150, minwidth=100)
        self.tree.column("url", width=400, minwidth=200)

        self.tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbar
        scrollbar = ttk.Scrollbar(
            data_frame,
            orient="vertical",
            command=self.tree.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

    def setup_status_bar(self):
        """Configure la barre de statut en bas"""
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            style='Status.TLabel',
            anchor="w"
        )
        status_bar.grid(row=2, column=0, sticky="ew")

    def fetch_data(self):
        if self.database_has_data():
            show_error("La base de données contient déjà des données.\n"
                  "Veuillez d'abord effacer la base avant de charger de nouvelles données.")
            self.update_status("Erreur: Base déjà chargée")
            return
        self.update_status("Téléchargement des données SpaceX en cours...")
        try:
            data = self.scraper.scrape_launches()
            if data:
                self.db.insert_data(data)
                self.update_status(f"{len(data)} lancements ajoutés à la base")
                show_info(f"{len(data)} lancements ajoutés à la base")
                self.display_data(data[:])
            else:
                self.update_status("Aucune donnée récupérée")
                show_warning("Aucune donnée n'a été récupérée")
        except Exception as e:
            self.update_status("Erreur lors du téléchargement")
            show_error(f"Erreur lors du téléchargement: {str(e)}")

    def display_data(self, data):
        """Affiche les données dans le Treeview"""
        self.tree.delete(*self.tree.get_children())

        for item in data:
            self.tree.insert("", "end",
                            values=(item['mission_name'],
                                   item['mission_date'],
                                   item['mission_url']))

    def clear_database(self):
        if ask_confirmation("Voulez-vous vraiment effacer toutes les données?"):
            self.db.clear_collection()
            self.update_status("Base de données effacée")
            show_info("Base de données effacée")
            self.tree.delete(*self.tree.get_children())

    def show_stats(self):
        try:
            # Remplacer l'ancienne requête par la nouvelle
            stats = self.db.get_launches_by_organization()

            if not stats:
                show_info("Aucune donnée disponible pour afficher les statistiques")
                return

            data = {
                'labels': [item['_id'] for item in stats],
                'values': [item['count'] for item in stats]
            }

            create_matplotlib_graph(
                self.root,
                data,
                title="Nombre de lancements par organisation/cliente"
            )

            self.update_status("Statistiques par organisation affichées")
        except Exception as e:
            self.update_status("Erreur lors de l'affichage des stats")
            show_error(f"Erreur lors de l'affichage des statistiques: {str(e)}")

    def update_status(self, message):
        self.status_var.set(f"STATUS: {message}")
        self.root.update_idletasks()

    def database_has_data(self):
        """Vérifie si la base de données contient déjà des données"""
        return self.db.get_launches_count() > 0

    def display_latest_data(self):
        try:

            latest_launches = list(self.db.collection.find().sort("_id", 1))

            if latest_launches:
                self.display_data(latest_launches)
                self.update_status(f"{len(latest_launches)} lancements chargés depuis la base")
            else:
                self.update_status("Aucune donnée à afficher")
        except Exception as e:
            self.update_status("Erreur lors du chargement des données")
            show_error(f"Erreur lors du chargement des données: {str(e)}")