import io
import re
import time
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

API_URL = "https://recherche-entreprises.api.gouv.fr/search"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
USER_AGENT = "EnrichissementSIRET-Streamlit/1.0"

st.set_page_config(page_title="Enrichissement SIRET", page_icon="🏢", layout="wide")


def normaliser_siret(valeur) -> str:
    """Transforme une valeur Excel en chaîne de 14 chiffres lorsque possible."""
    if pd.isna(valeur):
        return ""
    texte = str(valeur).strip()
    if texte.endswith(".0"):
        texte = texte[:-2]
    return re.sub(r"\D", "", texte)


def trouver_etablissement(resultat: dict, siret: str):
    """Retourne l'établissement correspondant exactement au SIRET demandé."""
    siege = resultat.get("siege") or {}
    if str(siege.get("siret", "")) == siret:
        return siege

    for etablissement in resultat.get("matching_etablissements") or []:
        if str(etablissement.get("siret", "")) == siret:
            return etablissement
    return None


def appeler_api(siret: str, session: requests.Session, tentatives: int = 4) -> dict:
    """Interroge l'API, gère les erreurs temporaires et retourne une ligne enrichie."""
    ligne = {
        "SIRET": siret,
        "SIREN": siret[:9] if len(siret) >= 9 else "",
        "Raison sociale": "",
        "Adresse": "",
        "Code postal": "",
        "Commune": "",
        "Code NAF": "",
        "État administratif": "",
        "Établissement siège": "",
        "Statut du traitement": "",
        "Message d'erreur": "",
    }

    if not re.fullmatch(r"\d{14}", siret):
        ligne["Statut du traitement"] = "SIRET invalide"
        ligne["Message d'erreur"] = "Le SIRET doit contenir exactement 14 chiffres"
        return ligne

    for tentative in range(1, tentatives + 1):
        try:
            reponse = session.get(
                API_URL,
                params={"q": siret, "page": 1, "per_page": 10},
                headers={"User-Agent": USER_AGENT, "Accept": "application/json"},
                timeout=20,
            )

            if reponse.status_code == 429:
                attente = int(reponse.headers.get("Retry-After", "2"))
                time.sleep(max(attente, 1))
                continue

            if 500 <= reponse.status_code < 600 and tentative < tentatives:
                time.sleep(2 ** (tentative - 1))
                continue

            reponse.raise_for_status()
            donnees = reponse.json()

            for resultat in donnees.get("results", []):
                etablissement = trouver_etablissement(resultat, siret)
                if etablissement:
                    ligne.update({
                        "SIREN": str(resultat.get("siren") or siret[:9]),
                        "Raison sociale": resultat.get("nom_raison_sociale") or resultat.get("nom_complet") or "",
                        "Adresse": etablissement.get("adresse") or "",
                        "Code postal": str(etablissement.get("code_postal") or ""),
                        "Commune": etablissement.get("libelle_commune") or "",
                        "Code NAF": etablissement.get("activite_principale") or resultat.get("activite_principale") or "",
                        "État administratif": etablissement.get("etat_administratif") or "",
                        "Établissement siège": "Oui" if etablissement.get("est_siege") is True else "Non",
                        "Statut du traitement": "Trouvé",
                    })
                    return ligne

            ligne["Statut du traitement"] = "Non trouvé"
            ligne["Message d'erreur"] = "Aucun établissement correspondant exactement au SIRET"
            return ligne

        except requests.RequestException as exc:
            if tentative == tentatives:
                ligne["Statut du traitement"] = "Erreur technique"
                ligne["Message d'erreur"] = str(exc)[:250]
                return ligne
            time.sleep(2 ** (tentative - 1))
        except ValueError as exc:
            ligne["Statut du traitement"] = "Erreur technique"
            ligne["Message d'erreur"] = f"Réponse JSON inexploitable : {exc}"
            return ligne

    ligne["Statut du traitement"] = "Erreur technique"
    ligne["Message d'erreur"] = "Nombre maximal de tentatives atteint"
    return ligne


def produire_excel(resultats: pd.DataFrame) -> bytes:
    """Produit un classeur formaté avec les feuilles Résultats et Synthèse."""
    total = len(resultats)
    trouves = int((resultats["Statut du traitement"] == "Trouvé").sum())
    invalides = int((resultats["Statut du traitement"] == "SIRET invalide").sum())
    non_trouves = int((resultats["Statut du traitement"] == "Non trouvé").sum())
    erreurs = int((resultats["Statut du traitement"] == "Erreur technique").sum())

    synthese = pd.DataFrame({
        "Indicateur": [
            "SIRET fournis", "Établissements trouvés", "SIRET invalides",
            "Établissements non trouvés", "Erreurs techniques", "Taux de réussite"
        ],
        "Valeur": [total, trouves, invalides, non_trouves, erreurs, trouves / total if total else 0],
    })

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resultats.to_excel(writer, sheet_name="Résultats", index=False)
        synthese.to_excel(writer, sheet_name="Synthèse", index=False)

        ws = writer.book["Résultats"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        couleur_entete = "1F4E78"
        for cellule in ws[1]:
            cellule.fill = PatternFill("solid", fgColor=couleur_entete)
            cellule.font = Font(color="FFFFFF", bold=True)
            cellule.alignment = Alignment(horizontal="center")

        largeurs = {
            "A": 18, "B": 14, "C": 38, "D": 52, "E": 14, "F": 25,
            "G": 14, "H": 20, "I": 20, "J": 24, "K": 55,
        }
        for colonne, largeur in largeurs.items():
            ws.column_dimensions[colonne].width = largeur

        col_statut = list(resultats.columns).index("Statut du traitement") + 1
        for ligne in range(2, ws.max_row + 1):
            statut = ws.cell(ligne, col_statut).value
            couleur = "C6EFCE" if statut == "Trouvé" else "FFC7CE"
            for cellule in ws[ligne]:
                cellule.fill = PatternFill("solid", fgColor=couleur)
                cellule.alignment = Alignment(vertical="top", wrap_text=True)
            ws.cell(ligne, 1).number_format = "@"
            ws.cell(ligne, 2).number_format = "@"

        ws2 = writer.book["Synthèse"]
        ws2.freeze_panes = "A2"
        ws2.column_dimensions["A"].width = 34
        ws2.column_dimensions["B"].width = 18
        for cellule in ws2[1]:
            cellule.fill = PatternFill("solid", fgColor=couleur_entete)
            cellule.font = Font(color="FFFFFF", bold=True)
        ws2["B7"].number_format = "0.0%"

    return buffer.getvalue()


st.title("Enrichissement d'une liste de SIRET")
st.caption("Importez un fichier Excel, interrogez l'API Recherche d'entreprises, puis téléchargez le résultat enrichi")

with st.expander("Format attendu", expanded=False):
    st.markdown("""
    - Fichier `.xlsx`
    - Une ligne par établissement
    - Une colonne contenant le SIRET
    - Les SIRET doivent comporter 14 chiffres
    """)

fichier = st.file_uploader("Déposer le fichier Excel", type=["xlsx"])

if fichier is not None:
    try:
        classeur = pd.ExcelFile(fichier, engine="openpyxl")
        feuille = st.selectbox("Feuille à traiter", classeur.sheet_names)
        donnees = pd.read_excel(classeur, sheet_name=feuille, dtype=str)

        if donnees.empty:
            st.warning("La feuille sélectionnée ne contient aucune ligne")
            st.stop()

        colonne_siret = st.selectbox(
            "Colonne contenant le SIRET",
            donnees.columns.tolist(),
            index=donnees.columns.tolist().index("SIRET") if "SIRET" in donnees.columns else 0,
        )
        st.write("Aperçu du fichier")
        st.dataframe(donnees.head(20), use_container_width=True)

        if st.button("Lancer l'enrichissement", type="primary"):
            sirets = donnees[colonne_siret].map(normaliser_siret).tolist()
            progression = st.progress(0, text="Initialisation")
            zone_statut = st.empty()
            lignes = []

            with requests.Session() as session:
                for index, siret in enumerate(sirets, start=1):
                    zone_statut.info(f"Traitement {index}/{len(sirets)} : {siret or 'valeur vide'}")
                    lignes.append(appeler_api(siret, session))
                    progression.progress(index / len(sirets), text=f"{index}/{len(sirets)} SIRET traités")
                    if index < len(sirets):
                        time.sleep(0.16)

            resultats = pd.DataFrame(lignes)
            excel = produire_excel(resultats)
            trouves = int((resultats["Statut du traitement"] == "Trouvé").sum())
            erreurs = len(resultats) - trouves

            zone_statut.success("Traitement terminé")
            col1, col2, col3 = st.columns(3)
            col1.metric("SIRET traités", len(resultats))
            col2.metric("Établissements trouvés", trouves)
            col3.metric("Anomalies ou erreurs", erreurs)
            st.dataframe(resultats, use_container_width=True)
            st.download_button(
                "Télécharger le fichier enrichi",
                data=excel,
                file_name=f"siret_enrichis_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime=XLSX_MIME,
                type="primary",
            )

    except Exception as exc:
        st.error(f"Impossible de traiter le fichier : {exc}")
