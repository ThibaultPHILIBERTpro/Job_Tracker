# Job Tracker : Application Desktop de Suivi de Candidatures

`v1.1 - En développement` · `Python 3.13` · `SQLite3` · `CustomTkinter`

## Objectif du Projet
Concevoir un outil local et autonome pour centraliser, piloter et analyser une recherche d'emploi. L'application remplace le traditionnel tableau Excel manuel par une base de données relationnelle SQLite couplée à une interface graphique moderne et des alertes automatiques de relance.

## Fonctionnalités actuelles (v1.1)
* **Tableau de bord & KPIs en temps réel :** Suivi des candidatures du jour, détection automatique des relances urgentes (J+15), relances à venir (< 3 jours) et processus d'entretiens actifs.
* **Base de données SQLite évolutive :** Initialisation automatique et gestion dynamique des migrations de schéma (`PRAGMA table_info` + `ALTER TABLE`) sans perte de données.
* **Mode Rapide (Web Scraping — *Bêta*) :** Pré-remplissage du titre et détection de la plateforme (LinkedIn, APEC, WTTJ, France Travail) depuis l'URL de l'offre en arrière-plan (`threading`).
* **Gestion du cycle de recrutement :** Filtres multi-critères, code couleur selon l'urgence, fiche détaillée éditable (suivi jusqu'à 3 tours d'entretiens, sauvegarde de l'annonce, lettre de motivation et motifs de refus).
* **Archivage Excel automatisé :** Génération silencieuse d'une sauvegarde formatée (`xlsxwriter`) à chaque modification + export manuel horodaté.

## Stack Technique
* **Interface (GUI) :** `CustomTkinter`, `Tkinter` (`ttk.Treeview`)
* **Base de données :** `sqlite3` (SQL natif)
* **Automatisation & Scraping :** `requests`, `BeautifulSoup4`, `threading`
* **Reporting :** `xlsxwriter`

## Roadmap & Versionning

### v1.1: Ergonomie & Suivi de Candidature
- [x] Ajout d'un script `Fake_data.py` pour générer une base fictive de démonstration en 1 clic.
- [X] Affichage et accès direct au lien URL de l'offre depuis la fiche détaillée.
- [X] Affichage visible de la date de candidature dans la fiche.
- [x] Ajout d'un champ obligatoire pour tracer la version du **CV envoyé** pour chaque offre (migration rétroactive en "CV Hybride" pour l'historique).
- [x] Ajout de **Meteojob** dans les sources + champs *Type de contrat* et *Source* vides par défaut et obligatoires en saisie.

### v1.2: Gestion Avancée des Relances & Réseau (LinkedIn)
- [ ] Option par annonce pour désactiver les relances ou personnaliser le délai de relance.
- [ ] Nouveau statut *"Relance impossible / contact introuvable"* avec bascule automatique en *"Sans réponse"* après 1 mois et demi d'inactivité.
- [ ] Module **Approche Réseau (LinkedIn)** : fenêtre dédiée lors d'une demande de connexion (profil du contact, message préparé) + nouvel indicateur jaune sur le Dashboard pour vérifier les demandes acceptées.

### v2.0: Refonte Architecture, Paramètres & Scraping
- [ ] Refactorisation complète en architecture modulaire (séparation UI / Base de données / Scraping).
- [ ] Nouvel onglet **Paramètres avancés** : gestion personnalisée des noms de CV (multi-CV), configuration globale des délais de relance, etc.
- [ ] Refonte du moteur de scraping pour extraire automatiquement le nom de l'entreprise, le lieu et le texte de l'offre.
- [ ] Refonte visuelle et amélioration de l'expérience utilisateur (UI/UX).