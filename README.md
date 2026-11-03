# Enrichissement SIRET avec Streamlit

Application de test permettant :

- d'importer un fichier Excel contenant une liste de SIRET
- d'interroger l'API Recherche d'entreprises de data.gouv.fr
- de récupérer la raison sociale, l'adresse et le code NAF
- de télécharger un classeur Excel enrichi et mis en forme

## Fichiers du projet

- `app.py` : application Streamlit
- `requirements.txt` : dépendances Python
- `exemple_siret.xlsx` : fichier d'entrée de démonstration

## Test local

Prérequis : Python 3.10 ou version ultérieure

### Windows

```bat
cd cheminers\enrichissement_siret_streamlit
python -m venv .venv
.venv\Scriptsctivate
python -m pip install -r requirements.txt
streamlit run app.py
```

### macOS ou Linux

```bash
cd /chemin/vers/enrichissement_siret_streamlit
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

L'application s'ouvre généralement sur `http://localhost:8501`

## Publication sur Posit Connect Cloud

1. Décompresser le projet
2. Créer un dépôt GitHub
3. Ajouter `app.py`, `requirements.txt` et éventuellement `exemple_siret.xlsx`
4. Envoyer les fichiers sur la branche principale
5. Ouvrir Posit Connect Cloud
6. Choisir la publication depuis GitHub
7. Sélectionner le dépôt et la branche
8. Choisir `app.py` comme fichier principal
9. Lancer le déploiement
10. Tester l'import et le téléchargement d'un fichier

## Données et limites

- Le fichier importé est traité en mémoire par l'application
- Les résultats dépendent des données ouvertes retournées par l'API
- Les SIRET à diffusion restreinte peuvent ne pas être retournés
- L'application temporise les appels et gère les réponses HTTP 429
- Pour un usage industriel, prévoir journalisation, supervision, tests et politique de conservation des données
