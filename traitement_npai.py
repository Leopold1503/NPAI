import os
import zipfile
import win32com.client
import pandas as pd
from datetime import datetime

# ============================================================
#                PARAMÈTRES GÉNÉRAUX
# ============================================================
DOSSIER_BASE = r"U:\Business Assurance\Revenue Assurance\RA_2024\FE2026\PND-NPAI\Fichiers Asterion\B2C\Etude Asterion 2025"
DOSSIER_CSV = os.path.join(DOSSIER_BASE, "Fichiers traités")

FICHIER_COLONNES = os.path.join(DOSSIER_BASE, "NPAI Léopold.xlsx")
FICHIER_COMPLET = os.path.join(DOSSIER_BASE, "NPAI 2025.xlsx")
FICHIER_CONSIGNE = os.path.join(DOSSIER_BASE, "Consigne_NPAI.xlsx")

COLONNES_VOULUES = ["ENTITÉ", "TYPE DE DOCUMENT", "SCS-CONTRAT", "DATE RÉCEPTION", "DATE TRAITEMENT PND"]
DOSSIER_TEMP = os.path.join(DOSSIER_BASE, "tmp_zip")
DATE_COMPARAISON = pd.Timestamp(2020, 1, 1)

# ============================================================
#         1. CHARGER OU CRÉER LA CONSIGNE
# ============================================================
def charger_consigne():
    if os.path.exists(FICHIER_CONSIGNE):
        return pd.read_excel(FICHIER_CONSIGNE)
    else:
        return pd.DataFrame(columns=["Date", "Fichier", "Statut"])

# ============================================================
#         2. RÉCUPÉRER LES ZIP DANS OUTLOOK
# ============================================================
def telecharger_zip_outlook():
    print("📩 Connexion à Outlook…")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    destinataire = outlook.CreateRecipient("SFR-RA-NPAI@sfr.com")
    destinataire.Resolve()
    if not destinataire.Resolved:
        raise Exception("❌ Impossible de trouver la boîte mail partagée SFR-RA-NPAI.")

    inbox = outlook.GetSharedDefaultFolder(destinataire, 6)  # 6 = Boîte de réception
    os.makedirs(DOSSIER_TEMP, exist_ok=True)

    nouveaux_fichiers = []
    for message in inbox.Items:
        try:
            if message.Attachments.Count > 0:
                for att in message.Attachments:
                    if att.FileName.lower().endswith(".zip"):
                        chemin_zip = os.path.join(DOSSIER_TEMP, att.FileName)
                        if not os.path.exists(chemin_zip):
                            att.SaveAsFile(chemin_zip)
                            print(f"📥 ZIP téléchargé : {att.FileName}")
                            nouveaux_fichiers.append(chemin_zip)
        except Exception as e:
            print(f"⚠️ Erreur lecture mail : {e}")

    return nouveaux_fichiers

# ============================================================
#         3. DÉZIPPER LES NOUVEAUX FICHIERS
# ============================================================
def extraire_zip(fichiers_zip):
    csv_extraits = []
    os.makedirs(DOSSIER_CSV, exist_ok=True)

    for fichier_zip in fichiers_zip:
        try:
            with zipfile.ZipFile(fichier_zip, "r") as zip_ref:
                zip_ref.extractall(DOSSIER_CSV)
                csv_extraits.extend(zip_ref.namelist())
                print(f"📂 Dézippé : {fichier_zip}")
            os.remove(fichier_zip)
        except Exception as e:
            print(f"⚠️ Erreur sur {fichier_zip} : {e}")
    return csv_extraits

# ============================================================
#         4. METTRE À JOUR LES AGRÉGATS
# ============================================================
def maj_aggregats(reconstruction_totale=False):
    df_log = charger_consigne()
    dfs_colonnes, dfs_complet = [], []

    fichiers_a_traiter = []
    for fichier in os.listdir(DOSSIER_CSV):
        if fichier.lower().endswith(".csv"):
            if reconstruction_totale or fichier not in df_log["Fichier"].values:
                fichiers_a_traiter.append(fichier)

    for fichier in fichiers_a_traiter:
        chemin = os.path.join(DOSSIER_CSV, fichier)
        print(f"📑 Lecture du CSV : {fichier}")

        try:
            df = pd.read_csv(chemin, encoding="utf-8", sep=None, engine="python")
        except UnicodeDecodeError:
            df = pd.read_csv(chemin, encoding="latin1", sep=None, engine="python")
        except Exception as e:
            print(f"⚠️ Erreur lecture {fichier} : {e}")
            continue

        colonnes_dispo = [col for col in COLONNES_VOULUES if col in df.columns]
        dfs_colonnes.append(df[colonnes_dispo])
        dfs_complet.append(df)

        # Mise à jour log
        df_log = pd.concat([
            df_log,
            pd.DataFrame([[datetime.today().date(), fichier, "X"]],
                         columns=["Date", "Fichier", "Statut"])
        ], ignore_index=True)

    if dfs_colonnes:
        fusion_colonnes = pd.concat(dfs_colonnes, ignore_index=True)

        # Onglet principal = toutes les colonnes demandées
        fusion_colonnes_clean = fusion_colonnes.drop_duplicates()

        # Onglet "Sans doublons" = une seule ligne par SCS-CONTRAT, avec date la plus proche de DATE_COMPARAISON
        sans_doublons = (
            fusion_colonnes
            .dropna(subset=["SCS-CONTRAT", "DATE TRAITEMENT PND"])
            .assign(Diff=lambda x: (pd.to_datetime(x["DATE TRAITEMENT PND"], errors="coerce") - DATE_COMPARAISON).abs())
            .sort_values("Diff")
            .drop_duplicates(subset=["SCS-CONTRAT"], keep="first")
            .drop(columns=["Diff"])
        )

        with pd.ExcelWriter(FICHIER_COLONNES, engine="openpyxl") as writer:
            fusion_colonnes_clean.to_excel(writer, sheet_name="Complet", index=False)
            sans_doublons.to_excel(writer, sheet_name="Sans doublons", index=False)

        print(f"✅ NPAI Léopold mis à jour avec feuille 'Sans doublons'")

    if dfs_complet:
        fusion_complet = pd.concat(dfs_complet, ignore_index=True)
        fusion_complet.drop_duplicates(inplace=True)
        fusion_complet.to_excel(FICHIER_COMPLET, index=False, engine="openpyxl")
        print(f"✅ NPAI 2025 mis à jour")

    df_log.drop_duplicates(subset=["Fichier"], inplace=True)
    df_log.to_excel(FICHIER_CONSIGNE, index=False)
    print(f"📝 Consigne mise à jour")

# ============================================================
#             5. PIPELINE GLOBAL
# ============================================================
def pipeline(reconstruction_totale=False):
    print("=== DÉMARRAGE DU PROCESS ===")
    fichiers_zip = telecharger_zip_outlook()
    if fichiers_zip:
        extraire_zip(fichiers_zip)
    maj_aggregats(reconstruction_totale=reconstruction_totale)
    print("=== PROCESS TERMINÉ ✅ ===")

# ============================================================
#             LANCEMENT DU SCRIPT
# ============================================================
if __name__ == "__main__":
    # Première exécution = reconstruire depuis zéro
    pipeline(reconstruction_totale=True)
