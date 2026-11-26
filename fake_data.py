import sqlite3
import os
import shutil
from datetime import datetime, timedelta

DB_NAME = "suivi_emploi.db"


def date_offset(jours: int) -> str:
    """Retourne une date au format YYYY-MM-DD décalée de X jours par rapport à aujourd'hui."""
    return (datetime.now() + timedelta(days=jours)).strftime("%Y-%m-%d")


def securiser_base_existante() -> bool:
    """Protège la vraie base de données avant de générer les données de démo."""
    if os.path.exists(DB_NAME):
        print(f"ATTENTION : Le fichier '{DB_NAME}' existe déjà !")
        print("Pour éviter toute perte de données, une sauvegarde automatique sera créée si tu continues.")
        choix = input("Veux-tu remplacer la base par des données fictives ? (tape 'OUI' pour confirmer) : ").strip()
        if choix != "OUI":
            print("Opération annulée. Ta base actuelle n'a pas été touchée.")
            return False

        horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"suivi_emploi_backup_{horodatage}.db"
        shutil.copy2(DB_NAME, backup_name)
        os.remove(DB_NAME)
        print(f"Sauvegarde créée : {backup_name}")
    return True


def generer_fake_data():
    if not securiser_base_existante():
        return

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Création de la table avec exactement le même schéma que app_suivi.py
    c.execute('''
        CREATE TABLE IF NOT EXISTS candidatures (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entreprise TEXT, poste TEXT, lieu TEXT, source TEXT, url TEXT,
            date_candidature TEXT, statut TEXT, date_relance_1 TEXT, date_relance_2 TEXT,
            contact TEXT, remarques TEXT, type_poste TEXT, salaire TEXT, avantages TEXT,
            offre_texte TEXT, lettre_motivation TEXT,
            date_entretien TEXT, details_entretien TEXT, motif_refus TEXT,
            date_entretien_2 TEXT, details_entretien_2 TEXT,
            date_entretien_3 TEXT, details_entretien_3 TEXT
        )
    ''')

    # Jeu de données réaliste couvrant tous les statuts et KPIs
    fausses_candidatures = [
        # 1. Candidature du jour (KPI Bleu)
        (
            "DataPulse Analytics", "Junior Data Analyst", "Paris", "Welcome to the Jungle",
            "https://www.welcometothejungle.com/fr/companies/datapulse/jobs/junior-data-analyst",
            date_offset(0), "Envoyé", date_offset(15), None,
            "Claire Moreau (RH)", "Équipe data de 6 personnes, stack SQL + Power BI.",
            "CDI", "38-42k€", "Télétravail 2j/sem, Tickets Resto",
            "Missions : création de dashboards Power BI, requêtage SQL, automatisation Python.",
            "Madame, Monsieur, diplômé d'un Master 2 en analyse de données...",
            None, None, None, None, None, None, None
        ),
        # 2. Relance 1 urgente dépassée (KPI Rouge + ligne rouge)
        (
            "FinTech Horizon", "Business Analyst Junior", "Paris", "LinkedIn",
            "https://www.linkedin.com/jobs/view/123456789",
            date_offset(-18), "Envoyé", date_offset(-3), None,
            "Marc Lefèvre", "Offre très alignée avec mon parcours économie/data.",
            "CDI", "40k€", "Mutuelle 100%, RTT",
            "Analyse des KPIs commerciaux et accompagnement des équipes métiers.",
            "Bonjour, je vous soumets ma candidature pour le poste de Business Analyst...",
            None, None, None, None, None, None, None
        ),
        # 3. Relance 2 urgente aujourd'hui (KPI Rouge + ligne rouge)
        (
            "GreenGrid Énergie", "Chargé d'Études Statistiques", "Orléans", "APEC",
            "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/987654",
            date_offset(-30), "Relance 1 Faite", date_offset(-15), date_offset(0),
            "Sophie Laurent", "Relance 1 envoyée par mail, pas encore de retour.",
            "CDI", "36-39k€", "Intéressement, 12 RTT",
            "Études de consommation énergétique sous Python et SQL.",
            "Madame Laurent, suite à votre offre sur l'APEC...",
            None, None, None, None, None, None, None
        ),
        # 4. Relance bientôt (< 3 jours) (KPI Orange)
        (
            "NovaRetail", "Data Analyst Marketing", "Télétravail", "HelloWork",
            "https://www.hellowork.com/fr-fr/emplois/555444333.html",
            date_offset(-13), "Envoyé", date_offset(2), None,
            "Service Recrutement", "Full remote possible avec 1 déplacement/mois à Paris.",
            "CDI", "39k€", "Forfait télétravail",
            "Suivi des campagnes CRM, segmentation clients et reporting mensuel.",
            "Bonjour, passionné par la valorisation des données clients...",
            None, None, None, None, None, None, None
        ),
        # 5. Processus d'entretien en cours - 2 tours (KPI Vert + ligne verte)
        (
            "AéroData Systems", "Data Analyst BI", "Paris", "Welcome to the Jungle",
            "https://www.welcometothejungle.com/fr/companies/aerodata/jobs/data-analyst-bi",
            date_offset(-20), "Entretien", date_offset(-5), None,
            "Julien Mercier (Lead Data)", "Très bon feeling au 1er entretien RH !",
            "CDI", "42k€", "CE, Tickets Resto, 2j TT",
            "Développement de pipelines Python et tableaux de bord décisionnels.",
            "Bonjour, mon expérience en automatisation Python et Power BI...",
            f"{date_offset(-5)} à 14h00", "Visio RH - Présentation du parcours et des prétentions.",
            None,
            f"{date_offset(3)} à 10h30", "Entretien Technique SQL/Python avec Julien (Lead Data).",
            None, None
        ),
        # 6. Processus d'entretien en cours - 1er tour (KPI Vert + ligne verte)
        (
            "Loire Conseil", "Consultant Data Junior", "Orléans", "France Travail",
            "https://candidat.francetravail.fr/offres/recherche/detail/1112223",
            date_offset(-10), "Entretien", date_offset(5), None,
            "Amélie Garnier", "Cabinet à taille humaine, clients variés.",
            "CDI", "35-38k€", "Voiture de fonction ou Navigo",
            "Accompagnement PME sur Power BI et structuration de bases SQL.",
            "Madame, Monsieur, disponible immédiatement sur Orléans...",
            f"{date_offset(2)} à 15h00", "Entretien dans les locaux à Orléans (apporter CV papier).",
            None, None, None, None, None
        ),
        # 7. Refus avec motif (Ligne grise)
        (
            "Banque Atlantique", "Analyste Risques & Données", "Paris", "APEC",
            "https://www.apec.fr/candidat/recherche-emploi.html/emploi/detail-offre/222333",
            date_offset(-25), "Refus", date_offset(-10), None,
            "DRH Pôle Risques", "Retour rapide et constructif.",
            "CDI", "43k€", "Participation, PERCO",
            "Modélisation des risques de crédit et reporting réglementaire.",
            "Madame, Monsieur, titulaire d'un Master 2 en Économie...",
            None, None, "Profil junior apprécié, mais ils ont retenu un candidat avec 3 ans d'expérience bancaire.",
            None, None, None, None
        ),
        # 8. Sans réponse (Ligne grise)
        (
            "MegaCorp Logistics", "Data Analyst Supply Chain", "Autre", "Indeed",
            "https://fr.indeed.com/viewjob?jk=abcdef123456",
            date_offset(-45), "Sans réponse", date_offset(-30), date_offset(-15),
            "", "Aucune nouvelle malgré 2 relances.",
            "CDD", "34k€", "Prime annuelle",
            "Analyse des stocks et optimisation des flux logistiques.",
            "Bonjour, je vous adresse ma candidature...",
            None, None, None, None, None, None, None
        ),
    ]

    c.executemany('''
        INSERT INTO candidatures (
            entreprise, poste, lieu, source, url,
            date_candidature, statut, date_relance_1, date_relance_2,
            contact, remarques, type_poste, salaire, avantages,
            offre_texte, lettre_motivation,
            date_entretien, details_entretien, motif_refus,
            date_entretien_2, details_entretien_2,
            date_entretien_3, details_entretien_3
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', fausses_candidatures)

    conn.commit()
    conn.close()
    print(f"Base de démonstration '{DB_NAME}' générée avec {len(fausses_candidatures)} candidatures fictives !")


if __name__ == "__main__":
    generer_fake_data()