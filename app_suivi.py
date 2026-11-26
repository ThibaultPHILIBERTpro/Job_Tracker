import customtkinter as ctk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import threading
import xlsxwriter
import os
import webbrowser

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DB_NAME = "suivi_emploi.db"

class JobTrackerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Tracker de Candidatures - V8 (Complet)")
        
        largeur = self.winfo_screenwidth()
        hauteur = self.winfo_screenheight()
        self.geometry(f"{largeur}x{hauteur}+0+0")
        self.after(200, lambda: self.state('zoomed'))
        
        self.init_db()
        self.build_ui()
        self.load_data()
        self.export_excel(silent=True)

    def init_db(self):
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS candidatures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entreprise TEXT, poste TEXT, lieu TEXT, source TEXT, url TEXT,
                date_candidature TEXT, statut TEXT, date_relance_1 TEXT, date_relance_2 TEXT
            )
        ''')
        c.execute("PRAGMA table_info(candidatures)")
        colonnes_existantes = [col[1] for col in c.fetchall()]
        
        colonnes_a_verifier = {
            "contact": "TEXT", "remarques": "TEXT", "type_poste": "TEXT",
            "salaire": "TEXT", "avantages": "TEXT", "offre_texte": "TEXT",
            "lettre_motivation": "TEXT", "date_entretien": "TEXT",
            "details_entretien": "TEXT", "motif_refus": "TEXT",
            "date_entretien_2": "TEXT", "details_entretien_2": "TEXT",
            "date_entretien_3": "TEXT", "details_entretien_3": "TEXT",
            "cv_envoye": "TEXT"
        }
        for col_name, col_type in colonnes_a_verifier.items():
            if col_name not in colonnes_existantes:
                c.execute(f"ALTER TABLE candidatures ADD COLUMN {col_name} {col_type}")

        # Remplissage rétroactif par défaut pour les anciennes candidatures
        c.execute("UPDATE candidatures SET cv_envoye = 'CV Hybride' WHERE cv_envoye IS NULL OR cv_envoye = ''")
        conn.commit()
        conn.close()

    def build_ui(self):
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.tab_dashboard = self.tabs.add("1. 📊 Accueil & KPIs")
        self.tab_auto = self.tabs.add("2. ⚡ Mode Rapide")
        self.tab_manuel = self.tabs.add("3. ✍️ Mode Saisie")
        self.tab_db = self.tabs.add("4. 🗄️ Ma Base & Filtres")

        # ================= ONGLET 1 : DASHBOARD =================
        dash_title = ctk.CTkLabel(self.tab_dashboard, text="Aperçu de ta recherche d'emploi", font=("Arial", 24, "bold"))
        dash_title.pack(pady=20)

        self.cards_frame = ctk.CTkFrame(self.tab_dashboard, fg_color="transparent")
        self.cards_frame.pack(pady=20, expand=True)

        self.kpi_jour = self.create_kpi_card(self.cards_frame, "Candidatures\naujourd'hui", "0", "#2980B9", 0)
        self.kpi_urgentes = self.create_kpi_card(self.cards_frame, "Relances\nà faire", "0", "#C0392B", 1)
        self.kpi_bientot = self.create_kpi_card(self.cards_frame, "Relances\n(< 3 jours)", "0", "#E67E22", 2)
        self.kpi_entretiens = self.create_kpi_card(self.cards_frame, "Processus\nen cours", "0", "#27AE60", 3)

        # ================= ONGLET 2 : AUTO =================
        ctk.CTkLabel(self.tab_auto, text="Colle l'URL de l'offre.", font=("Arial", 14, "italic")).pack(pady=30)
        self.url_auto_entry = ctk.CTkEntry(self.tab_auto, width=700, height=40)
        self.url_auto_entry.pack(pady=15)
        self.btn_scrape = ctk.CTkButton(self.tab_auto, text="Analyser l'URL", command=self.trigger_autofill, fg_color="#E67E22", height=40, font=("Arial", 14, "bold"))
        self.btn_scrape.pack(pady=10)

        # ================= ONGLET 3 : SAISIE =================
        form_frame = ctk.CTkScrollableFrame(self.tab_manuel, fg_color="transparent")
        form_frame.pack(pady=5, padx=10, fill="both", expand=True)

        ctk.CTkLabel(form_frame, text="Entreprise:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entreprise_entry = ctk.CTkEntry(form_frame, width=200)
        self.entreprise_entry.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="Poste:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.poste_entry = ctk.CTkEntry(form_frame, width=250)
        self.poste_entry.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="Lieu:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.lieu_combo = ctk.CTkComboBox(form_frame, values=["Paris", "Orléans", "Télétravail", "Autre"], width=150)
        self.lieu_combo.grid(row=0, column=5, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="Type:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.type_combo = ctk.CTkComboBox(form_frame, values=["CDI", "CDD", "Intérim", "Stage", "Alternance", "Freelance"], width=200)
        self.type_combo.grid(row=1, column=1, padx=5, pady=5)
        self.type_combo.set("")

        ctk.CTkLabel(form_frame, text="Salaire:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        self.salaire_entry = ctk.CTkEntry(form_frame, width=250)
        self.salaire_entry.grid(row=1, column=3, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="Avantages:").grid(row=1, column=4, padx=5, pady=5, sticky="e")
        self.avantages_entry = ctk.CTkEntry(form_frame, width=150)
        self.avantages_entry.grid(row=1, column=5, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="Source:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.source_combo = ctk.CTkComboBox(form_frame, values=["Welcome to the Jungle", "APEC", "France Travail", "LinkedIn", "HelloWork", "Indeed", "Meteojob", "Autre"], width=200)
        self.source_combo.grid(row=2, column=1, padx=5, pady=5)
        self.source_combo.set("")

        ctk.CTkLabel(form_frame, text="Contact:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.contact_entry = ctk.CTkEntry(form_frame, width=250)
        self.contact_entry.grid(row=2, column=3, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="URL:").grid(row=2, column=4, padx=5, pady=5, sticky="e")
        self.url_manuel_entry = ctk.CTkEntry(form_frame, width=150)
        self.url_manuel_entry.grid(row=2, column=5, padx=5, pady=5)

        ctk.CTkLabel(form_frame, text="CV envoyé:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.cv_combo = ctk.CTkComboBox(form_frame, values=["CV Hybride", "CV Business", "CV Tech", "Autre"], width=200)
        self.cv_combo.grid(row=3, column=1, padx=5, pady=5)
        self.cv_combo.set("")

        ctk.CTkLabel(form_frame, text="Remarques:").grid(row=4, column=0, padx=5, pady=5, sticky="ne")
        self.remarques_textbox = ctk.CTkTextbox(form_frame, width=800, height=40)
        self.remarques_textbox.grid(row=4, column=1, columnspan=5, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(form_frame, text="Texte offre:").grid(row=5, column=0, padx=5, pady=5, sticky="ne")
        self.offre_texte_textbox = ctk.CTkTextbox(form_frame, width=800, height=120)
        self.offre_texte_textbox.grid(row=5, column=1, columnspan=5, padx=5, pady=5, sticky="w")
        
        ctk.CTkLabel(form_frame, text="Lettre Motiv:").grid(row=6, column=0, padx=5, pady=5, sticky="ne")
        self.lettre_textbox = ctk.CTkTextbox(form_frame, width=800, height=120)
        self.lettre_textbox.grid(row=6, column=1, columnspan=5, padx=5, pady=5, sticky="w")

        self.add_btn = ctk.CTkButton(form_frame, text="Valider & Ajouter", command=self.add_candidature, fg_color="#27AE60", width=300, height=40, font=("Arial", 14, "bold"))
        self.add_btn.grid(row=7, column=1, columnspan=5, pady=15)

        # ================= ONGLET 4 : FILTRES & BASE =================
        filter_frame = ctk.CTkFrame(self.tab_db)
        filter_frame.pack(pady=5, padx=10, fill="x")
        
        ctk.CTkLabel(filter_frame, text="Recherche (Boîte/Poste):").pack(side="left", padx=5)
        self.search_entry = ctk.CTkEntry(filter_frame, width=150)
        self.search_entry.pack(side="left", padx=5)

        ctk.CTkLabel(filter_frame, text="Statut:").pack(side="left", padx=5)
        self.filter_statut = ctk.CTkComboBox(filter_frame, values=["Tous", "Envoyé", "Relance 1 Faite", "Relance 2 Faite", "Entretien", "Refus", "Sans réponse"], width=130)
        self.filter_statut.pack(side="left", padx=5)
        self.filter_statut.set("Tous")

        ctk.CTkLabel(filter_frame, text="Lieu:").pack(side="left", padx=5)
        self.filter_lieu = ctk.CTkComboBox(filter_frame, values=["Tous", "Paris", "Orléans", "Télétravail", "Autre"], width=100)
        self.filter_lieu.pack(side="left", padx=5)
        self.filter_lieu.set("Tous")

        ctk.CTkButton(filter_frame, text="🔍 Filtrer", command=self.load_data, width=80).pack(side="left", padx=10)
        ctk.CTkButton(filter_frame, text="✖ Reset", command=self.reset_filters, fg_color="#7F8C8D", width=80).pack(side="left", padx=5)

        self.tree_frame = ctk.CTkFrame(self.tab_db)
        self.tree_frame.pack(pady=5, padx=10, fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=30, fieldbackground="#2b2b2b", borderwidth=0, font=('Arial', 10))
        style.map('Treeview', background=[('selected', '#1f538d')])
        style.configure("Treeview.Heading", background="#1f538d", foreground="white", font=('Arial', 11, 'bold'))

        columns = ("ID", "Entreprise", "Poste", "Type", "Lieu", "Date", "Statut", "Prochaine Action")
        self.tree = ttk.Treeview(self.tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            width_col = 100
            if col == "Poste": width_col = 220
            elif col == "Entreprise": width_col = 150
            elif col == "ID": width_col = 40
            elif col == "Type" or col == "Lieu": width_col = 80
            self.tree.column(col, width=width_col)
        self.tree.pack(fill="both", expand=True)

        self.tree.tag_configure('alerte', background='#781c1c') 
        self.tree.tag_configure('ok', background='#2b2b2b')     
        self.tree.tag_configure('mort', background='#4a4a4a')
        self.tree.tag_configure('victoire', background='#1E8449') 

        self.action_frame = ctk.CTkFrame(self.tab_db)
        self.action_frame.pack(pady=10, padx=10, fill="x")
        
        self.export_btn = ctk.CTkButton(self.action_frame, text="📥 Archiver Excel", command=self.export_excel, fg_color="#8E44AD")
        self.export_btn.pack(side="left", padx=10)

        self.open_fiche_btn = ctk.CTkButton(self.action_frame, text="Ouvrir Fiche Complète", command=self.open_fiche_complete, fg_color="#2980B9", font=("Arial", 12, "bold"))
        self.open_fiche_btn.pack(side="left", padx=20)

        self.relance_btn = ctk.CTkButton(self.action_frame, text="Appliquer Statut", command=self.update_statut, fg_color="#C0392B")
        self.relance_btn.pack(side="right", padx=10)
        self.statut_combo = ctk.CTkComboBox(self.action_frame, values=["Envoyé", "Relance 1 Faite", "Relance 2 Faite", "Entretien", "Refus", "Sans réponse"])
        self.statut_combo.pack(side="right", padx=10)

    def create_kpi_card(self, parent, titre, valeur, couleur, col):
        card = ctk.CTkFrame(parent, fg_color=couleur, corner_radius=15, width=200, height=120)
        card.grid(row=0, column=col, padx=20, pady=20)
        card.pack_propagate(False)
        ctk.CTkLabel(card, text=titre, font=("Arial", 16, "bold"), text_color="white", justify="center").pack(pady=(15,5))
        lbl_val = ctk.CTkLabel(card, text=valeur, font=("Arial", 36, "bold"), text_color="white")
        lbl_val.pack()
        return lbl_val

    def update_dashboard(self, rows):
        aujourdhui_dt = datetime.now().date()
        aujourdhui_str = aujourdhui_dt.strftime("%Y-%m-%d")
        
        c_jour, r_urgentes, r_bientot, entretiens = 0, 0, 0, 0

        for row in rows:
            date_cand = row[5]
            st = row[6]
            r1 = row[7]
            r2 = row[8]

            if date_cand == aujourdhui_str: c_jour += 1
            if st == "Entretien": entretiens += 1

            target = None
            if st == "Envoyé" and r1: target = r1
            elif st == "Relance 1 Faite" and r2: target = r2

            if target:
                try:
                    target_dt = datetime.strptime(target, "%Y-%m-%d").date()
                    delta = (target_dt - aujourdhui_dt).days
                    if delta <= 0: r_urgentes += 1
                    elif 0 < delta <= 3: r_bientot += 1
                except: pass

        self.kpi_jour.configure(text=str(c_jour))
        self.kpi_urgentes.configure(text=str(r_urgentes))
        self.kpi_bientot.configure(text=str(r_bientot))
        self.kpi_entretiens.configure(text=str(entretiens))

    def reset_filters(self):
        self.search_entry.delete(0, 'end')
        self.filter_statut.set("Tous")
        self.filter_lieu.set("Tous")
        self.load_data()

    def trigger_autofill(self):
        url = self.url_auto_entry.get()
        if not url: return
        self.url_manuel_entry.delete(0, 'end')
        self.url_manuel_entry.insert(0, url)
        self.btn_scrape.configure(text="Recherche en cours...")
        threading.Thread(target=self.scrape_info, args=(url,)).start()

    def scrape_info(self, url):
        try:
            response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
            soup = BeautifulSoup(response.text, 'html.parser')
            title = soup.title.string if soup.title else ""
            self.poste_entry.delete(0, 'end')
            self.poste_entry.insert(0, title[:50]) 
            if "welcometothejungle" in url: self.source_combo.set("Welcome to the Jungle")
            elif "apec" in url: self.source_combo.set("APEC")
            elif "linkedin" in url: self.source_combo.set("LinkedIn")
            elif "francetravail" in url or "pole-emploi" in url: self.source_combo.set("France Travail")
            elif "hellowork" in url: self.source_combo.set("HelloWork")
            elif "indeed" in url: self.source_combo.set("Indeed")
            elif "meteojob" in url: self.source_combo.set("Meteojob")
            self.tabs.set("3. ✍️ Mode Saisie")
        except: pass
        finally: self.btn_scrape.configure(text="Analyser l'URL")

    def add_candidature(self):
        ent = self.entreprise_entry.get().strip()
        poste = self.poste_entry.get().strip()
        type_poste = self.type_combo.get().strip()
        source = self.source_combo.get().strip()
        cv = self.cv_combo.get().strip()
        if not all([ent, poste, type_poste, source, cv]):
            return messagebox.showwarning("Attention", "Entreprise, poste, type de contrat, source et CV envoyé sont obligatoires.")
        
        date_cand = datetime.now().strftime("%Y-%m-%d")
        date_r1 = (datetime.now() + timedelta(days=15)).strftime("%Y-%m-%d")
        
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''INSERT INTO candidatures 
                    (entreprise, poste, lieu, type_poste, salaire, avantages, source, url, date_candidature, statut, date_relance_1, contact, remarques, offre_texte, lettre_motivation, cv_envoye)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                (ent, poste, self.lieu_combo.get(), self.type_combo.get(), self.salaire_entry.get(), self.avantages_entry.get(), self.source_combo.get(), self.url_manuel_entry.get(), date_cand, "Envoyé", date_r1, self.contact_entry.get(), self.remarques_textbox.get("1.0", "end-1c"), self.offre_texte_textbox.get("1.0", "end-1c"), self.lettre_textbox.get("1.0", "end-1c"), self.cv_combo.get()))
        conn.commit()
        conn.close()
        
        for w in [self.entreprise_entry, self.poste_entry, self.contact_entry, self.url_manuel_entry, self.url_auto_entry, self.salaire_entry, self.avantages_entry]: w.delete(0, 'end')
        self.remarques_textbox.delete("1.0", "end")
        self.offre_texte_textbox.delete("1.0", "end")
        self.lettre_textbox.delete("1.0", "end")
        self.cv_combo.set("")
        self.type_combo.set("")
        self.source_combo.set("")
        
        self.load_data()
        self.export_excel(silent=True)
        self.tabs.set("1. 📊 Accueil & KPIs")
        messagebox.showinfo("Super", "Candidature ajoutée !")

    def load_data(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        
        recherche = f"%{self.search_entry.get()}%"
        statut = self.filter_statut.get()
        lieu = self.filter_lieu.get()
        
        query = "SELECT id, entreprise, poste, type_poste, lieu, date_candidature, statut, date_relance_1, date_relance_2 FROM candidatures WHERE (entreprise LIKE ? OR poste LIKE ?)"
        params = [recherche, recherche]
        
        if statut != "Tous":
            query += " AND statut = ?"
            params.append(statut)
        if lieu != "Tous":
            query += " AND lieu = ?"
            params.append(lieu)

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(query, params)
        rows = c.fetchall()
        
        c.execute("SELECT id, entreprise, poste, type_poste, lieu, date_candidature, statut, date_relance_1, date_relance_2 FROM candidatures")
        all_rows = c.fetchall()
        conn.close()

        self.update_dashboard(all_rows)

        aujourdhui = datetime.now().strftime("%Y-%m-%d")
        for row in rows:
            id_cand, ent, poste, type_poste, lieu_cand, date_cand, st, r1, r2 = row
            tag, action = 'ok', "En cours"
            if st in ["Sans réponse", "Refus"]: tag, action = 'mort', "Terminé"
            elif st == "Entretien": tag, action = 'victoire', "Préparer Entretien"
            elif st == "Envoyé":
                action = f"Relance 1 le {r1}"
                if aujourdhui >= (r1 or ""): tag = 'alerte'
            elif st == "Relance 1 Faite":
                action = f"Relance 2 le {r2}" if r2 else "En attente"
                if r2 and aujourdhui >= r2: tag = 'alerte'
            self.tree.insert("", "end", values=(id_cand, ent, poste, type_poste, lieu_cand, date_cand, st, action), tags=(tag,))

    def update_statut(self):
        selected = self.tree.selection()
        if not selected: return messagebox.showwarning("Erreur", "Sélectionne une offre dans le tableau.")
        cand_id = self.tree.item(selected)['values'][0]
        nouveau_statut = self.statut_combo.get()
        
        if nouveau_statut == "Entretien":
            self.popup_entretien(cand_id)
            return
        elif nouveau_statut == "Refus":
            self.popup_refus(cand_id)
            return
            
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        if nouveau_statut == "Relance 1 Faite":
            date_r2 = (datetime.now() + timedelta(days=15)).strftime("%Y-%m-%d")
            c.execute("UPDATE candidatures SET statut=?, date_relance_2=? WHERE id=?", (nouveau_statut, date_r2, cand_id))
        else:
            c.execute("UPDATE candidatures SET statut=? WHERE id=?", (nouveau_statut, cand_id))
        conn.commit()
        conn.close()
        self.load_data()
        self.export_excel(silent=True)

    def popup_entretien(self, cand_id):
        pop = ctk.CTkToplevel(self)
        pop.title("Félicitations ! 🎉 Préparation Entretien")
        pop.geometry("600x450")
        pop.grab_set()

        ctk.CTkLabel(pop, text="Étape de l'entretien :").pack(pady=(10,0))
        combo_etape = ctk.CTkComboBox(pop, values=["1er Entretien (RH / Préqual)", "2ème Entretien (Tech / Manager)", "3ème Entretien (Final)"], width=250)
        combo_etape.pack(pady=5)

        ctk.CTkLabel(pop, text="Date et Heure :").pack(pady=(10,0))
        e_date = ctk.CTkEntry(pop, width=300, placeholder_text="ex: 28/10/2026 à 14h30")
        e_date.pack(pady=5)

        ctk.CTkLabel(pop, text="Lieu / Lien Visio & Choses à préparer :").pack(pady=(10,0))
        e_details = ctk.CTkTextbox(pop, width=500, height=120)
        e_details.pack(pady=5)

        def save():
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            etape = combo_etape.get()
            if "1er" in etape:
                c.execute("UPDATE candidatures SET statut=?, date_entretien=?, details_entretien=? WHERE id=?", 
                          ("Entretien", e_date.get(), e_details.get("1.0", "end-1c"), cand_id))
            elif "2ème" in etape:
                c.execute("UPDATE candidatures SET statut=?, date_entretien_2=?, details_entretien_2=? WHERE id=?", 
                          ("Entretien", e_date.get(), e_details.get("1.0", "end-1c"), cand_id))
            else:
                c.execute("UPDATE candidatures SET statut=?, date_entretien_3=?, details_entretien_3=? WHERE id=?", 
                        ("Entretien", e_date.get(), e_details.get("1.0", "end-1c"), cand_id))
            conn.commit()
            conn.close()
            self.load_data()
            self.export_excel(silent=True)
            pop.destroy()
            
        ctk.CTkButton(pop, text="Valider l'entretien", command=save, fg_color="#27AE60").pack(pady=20)

    def popup_refus(self, cand_id):
        pop = ctk.CTkToplevel(self)
        pop.title("Candidature Refusée")
        pop.geometry("500x250")
        pop.grab_set()

        ctk.CTkLabel(pop, text="Motif du refus (si fourni) :").pack(pady=(20,5))
        e_motif = ctk.CTkTextbox(pop, width=400, height=100)
        e_motif.pack(pady=5)

        def save():
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute("UPDATE candidatures SET statut=?, motif_refus=? WHERE id=?", 
                    ("Refus", e_motif.get("1.0", "end-1c"), cand_id))
            conn.commit()
            conn.close()
            self.load_data()
            self.export_excel(silent=True)
            pop.destroy()
            
        ctk.CTkButton(pop, text="Valider le refus", command=save, fg_color="#C0392B").pack(pady=15)

    def open_fiche_complete(self):
        selected = self.tree.selection()
        if not selected: return messagebox.showwarning("Sélection", "Clique sur une offre d'abord.")
        cand_id = self.tree.item(selected)['values'][0]
        
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute('''SELECT entreprise, poste, lieu, type_poste, salaire, avantages, source, contact, url, 
                    remarques, offre_texte, lettre_motivation, date_entretien, details_entretien, motif_refus, statut,
                    date_entretien_2, details_entretien_2, date_entretien_3, details_entretien_3, date_candidature, cv_envoye
                    FROM candidatures WHERE id=?''', (cand_id,))
        row = c.fetchone()
        conn.close()
        if not row: return

        popup = ctk.CTkToplevel(self)
        popup.title(f"Fiche détaillée - {row[0]}")
        popup.after(200, lambda: popup.state('zoomed'))
        popup.grab_set() 

        bottom_frame = ctk.CTkFrame(popup, fg_color="transparent")
        bottom_frame.pack(side="bottom", fill="x", pady=15)

        scroll = ctk.CTkScrollableFrame(popup, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        f_haut = ctk.CTkFrame(scroll, fg_color="transparent")
        f_haut.pack(fill="x")

        ctk.CTkLabel(f_haut, text="Entreprise:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        e_ent = ctk.CTkEntry(f_haut, width=150)
        e_ent.insert(0, row[0] or "")
        e_ent.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Poste:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        e_poste = ctk.CTkEntry(f_haut, width=200)
        e_poste.insert(0, row[1] or "")
        e_poste.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Lieu:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        e_lieu = ctk.CTkComboBox(f_haut, values=["Paris", "Orléans", "Télétravail", "Autre"], width=150)
        e_lieu.set(row[2] if row[2] else "Paris")
        e_lieu.grid(row=0, column=5, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Type:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        e_type = ctk.CTkComboBox(f_haut, values=["CDI", "CDD", "Intérim", "Stage", "Alternance", "Freelance"], width=150)
        e_type.set(row[3] if row[3] else "CDI")
        e_type.grid(row=1, column=1, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Salaire:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        e_sal = ctk.CTkEntry(f_haut, width=200)
        e_sal.insert(0, row[4] or "")
        e_sal.grid(row=1, column=3, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Avantages:").grid(row=1, column=4, padx=5, pady=5, sticky="e")
        e_av = ctk.CTkEntry(f_haut, width=150)
        e_av.insert(0, row[5] or "")
        e_av.grid(row=1, column=5, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Source:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        e_src = ctk.CTkComboBox(f_haut, values=["Welcome to the Jungle", "APEC", "France Travail", "LinkedIn", "HelloWork", "Indeed", "Meteojob", "Autre"], width=150)
        e_src.set(row[6] if row[6] else "APEC")
        e_src.grid(row=2, column=1, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Contact:").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        e_cont = ctk.CTkEntry(f_haut, width=200)
        e_cont.insert(0, row[7] or "")
        e_cont.grid(row=2, column=3, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Statut:").grid(row=2, column=4, padx=5, pady=5, sticky="e")
        e_stat = ctk.CTkComboBox(f_haut, values=["Envoyé", "Relance 1 Faite", "Relance 2 Faite", "Entretien", "Refus", "Sans réponse"], width=150)
        e_stat.set(row[15] if row[15] else "Envoyé")
        e_stat.grid(row=2, column=5, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="URL de l'offre:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        e_url = ctk.CTkEntry(f_haut, width=550)
        e_url.insert(0, row[8] or "")
        e_url.grid(row=3, column=1, columnspan=4, padx=5, pady=5, sticky="ew")

        def ouvrir_url():
            lien = e_url.get().strip()
            if not lien:
                return messagebox.showwarning("Lien manquant", "Aucune URL enregistrée pour cette offre.")
            if not lien.startswith(("http://", "https://")):
                lien = "https://" + lien
            webbrowser.open(lien)

        ctk.CTkButton(f_haut, text=" Ouvrir l'offre", command=ouvrir_url, fg_color="#2980B9", width=150).grid(row=3, column=5, padx=5, pady=5)

        ctk.CTkLabel(f_haut, text="Date candidature:", font=("Arial", 12, "bold")).grid(row=4, column=0, padx=5, pady=5, sticky="e")
        e_date_cand = ctk.CTkEntry(f_haut, width=150)
        e_date_cand.insert(0, row[20] or "")
        e_date_cand.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(f_haut, text="CV envoyé:", font=("Arial", 12, "bold")).grid(row=4, column=2, padx=5, pady=5, sticky="e")
        e_cv = ctk.CTkComboBox(f_haut, values=["CV Hybride", "CV Business", "CV Tech", "Autre"], width=200)
        e_cv.set(row[21] if row[21] else "CV Hybride")
        e_cv.grid(row=4, column=3, padx=5, pady=5, sticky="w")

        ctk.CTkLabel(scroll, text="Remarques :").pack(anchor="w", padx=10, pady=(10,0))
        e_rem = ctk.CTkTextbox(scroll, width=1000, height=60)
        e_rem.insert("1.0", row[9] or "")
        e_rem.pack(padx=10, pady=5)
        
        f_spec = ctk.CTkFrame(scroll)
        f_spec.pack(fill="x", pady=5, padx=10)
        
        ctk.CTkLabel(f_spec, text="Entretien 1:").grid(row=0, column=0, padx=5, pady=5)
        e_d_ent1 = ctk.CTkEntry(f_spec, width=120)
        e_d_ent1.insert(0, row[12] or "")
        e_d_ent1.grid(row=0, column=1, padx=5, pady=5)
        e_det_ent1 = ctk.CTkEntry(f_spec, width=250)
        e_det_ent1.insert(0, row[13] or "")
        e_det_ent1.grid(row=0, column=2, padx=5, pady=5)

        ctk.CTkLabel(f_spec, text="Entretien 2:").grid(row=1, column=0, padx=5, pady=5)
        e_d_ent2 = ctk.CTkEntry(f_spec, width=120)
        e_d_ent2.insert(0, row[16] or "")
        e_d_ent2.grid(row=1, column=1, padx=5, pady=5)
        e_det_ent2 = ctk.CTkEntry(f_spec, width=250)
        e_det_ent2.insert(0, row[17] or "")
        e_det_ent2.grid(row=1, column=2, padx=5, pady=5)

        ctk.CTkLabel(f_spec, text="Entretien 3:").grid(row=2, column=0, padx=5, pady=5)
        e_d_ent3 = ctk.CTkEntry(f_spec, width=120)
        e_d_ent3.insert(0, row[18] or "")
        e_d_ent3.grid(row=2, column=1, padx=5, pady=5)
        e_det_ent3 = ctk.CTkEntry(f_spec, width=250)
        e_det_ent3.insert(0, row[19] or "")
        e_det_ent3.grid(row=2, column=2, padx=5, pady=5)

        ctk.CTkLabel(f_spec, text="Motif Refus:").grid(row=0, column=3, padx=20, pady=5, sticky="ne")
        e_m_ref = ctk.CTkTextbox(f_spec, width=250, height=80)
        e_m_ref.insert("1.0", row[14] or "")
        e_m_ref.grid(row=0, column=4, rowspan=3, padx=5, pady=5)

        ctk.CTkLabel(scroll, text="Lettre de motivation :").pack(anchor="w", padx=10)
        e_lettre = ctk.CTkTextbox(scroll, width=1000, height=150)
        e_lettre.insert("1.0", row[11] or "")
        e_lettre.pack(padx=10, pady=5)

        ctk.CTkLabel(scroll, text="Texte complet de l'annonce :").pack(anchor="w", padx=10)
        e_offre = ctk.CTkTextbox(scroll, width=1000, height=250)
        e_offre.insert("1.0", row[10] or "")
        e_offre.pack(padx=10, pady=5)

        def save_mods():
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute('''UPDATE candidatures SET entreprise=?, poste=?, lieu=?, type_poste=?, salaire=?, avantages=?, 
                        source=?, contact=?, url=?, statut=?, date_candidature=?, cv_envoye=?, remarques=?, offre_texte=?, lettre_motivation=?, 
                        date_entretien=?, details_entretien=?, motif_refus=?,
                        date_entretien_2=?, details_entretien_2=?, date_entretien_3=?, details_entretien_3=? WHERE id=?''', 
                    (e_ent.get(), e_poste.get(), e_lieu.get(), e_type.get(), e_sal.get(), e_av.get(), 
                    e_src.get(), e_cont.get(), e_url.get().strip(), e_stat.get(), e_date_cand.get().strip(), e_cv.get(), e_rem.get("1.0", "end-1c"), e_offre.get("1.0", "end-1c"), 
                    e_lettre.get("1.0", "end-1c"), e_d_ent1.get(), e_det_ent1.get(), e_m_ref.get("1.0", "end-1c"),
                    e_d_ent2.get(), e_det_ent2.get(), e_d_ent3.get(), e_det_ent3.get(), cand_id))
            conn.commit()
            conn.close()
            self.load_data()
            self.export_excel(silent=True)
            popup.destroy()

        ctk.CTkButton(bottom_frame, text="💾 Sauvegarder modifications", command=save_mods, fg_color="#27AE60", height=40, font=("Arial", 14, "bold")).pack()

    def export_excel(self, silent=False):
        try:
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute('''SELECT id, entreprise, poste, lieu, type_poste, salaire, avantages, source, url, 
                        date_candidature, statut, cv_envoye, date_relance_1, date_relance_2, contact, remarques, 
                        date_entretien, details_entretien, date_entretien_2, details_entretien_2, 
                        date_entretien_3, details_entretien_3, motif_refus FROM candidatures''')
            rows = c.fetchall()
            columns = [desc[0].upper() for desc in c.description]
            conn.close()

            nom_fichier = "Suivi_Candidatures_Automatique.xlsx" if silent else f"Export_Archives_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            workbook = xlsxwriter.Workbook(nom_fichier)
            worksheet = workbook.add_worksheet("Recherche")
            fmt = workbook.add_format({'bold': True, 'bg_color': '#4F81BD', 'font_color': 'white'})

            for col_num, col_name in enumerate(columns):
                worksheet.write(0, col_num, col_name, fmt)
                worksheet.set_column(col_num, col_num, 16)
            for row_num, row_data in enumerate(rows, 1):
                for col_num, cell_data in enumerate(row_data):
                    worksheet.write(row_num, col_num, cell_data)
            workbook.close()
            
            if not silent: messagebox.showinfo("Export Réussi", f"Fichier généré : {nom_fichier}")
        except Exception as e:
            if not silent: messagebox.showerror("Erreur d'export", f"Fermez le fichier Excel s'il est ouvert.\nErreur : {e}")

if __name__ == "__main__":
    app = JobTrackerApp()
    app.mainloop()