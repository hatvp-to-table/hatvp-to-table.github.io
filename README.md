# 🏛️ HATVP To Table

Interface web pour explorer le répertoire des représentants d'intérêts (lobbyistes) de la HATVP et l'exporter en Excel.

## Fonctionnalités

- **5 modes de recherche** :
  - **Objets d'activité** : full-text dans plus de 100 000 descriptions libres d'actions de lobbying
  - **Secteurs d'activité** : dans les 31 catégories prédéfinies (Energie, Santé, Numérique…)
  - **Organisations** : dénomination, nom d'usage, sigle
  - **Personnes** : nom / prénom des dirigeants et collaborateurs
  - **Donneurs d'ordre** : clients déclarés par les cabinets mandataires
- **Requêtes booléennes** : critères combinés avec `ET`, `OU`, `SAUF`
  (ET / SAUF prioritaires sur OU : `A ET B OU C SAUF D` = `(A ET B) OU (C SAUF D)`)
- **Matching élargi ou exact** : accents et casse ignorés ; élargi = sous-chaînes et variantes, exact = mot entier
- Affichage interactif + **export Excel en 4 onglets** (Actions, Organisations, Dirigeants & Collaborateurs, Clients & Mandats)
- **Liens profonds** : `?mode=organisations&q=MEDEF&exact=1` ouvre l'app avec la recherche pré-remplie
  (`mode` ∈ `objets`, `secteurs`, `organisations`, `personnes`, `donneurs`)

Le projet comporte deux parties :

- **l'app Streamlit** (`app.py`) : la recherche avancée, hébergée sur Render ;
- **le site statique** (`static_site/`) : une page par organisation et par secteur, toujours en ligne
  et indexable par les moteurs de recherche, régénérée chaque nuit et hébergée sur GitHub Pages.
  Ses pages renvoient vers l'app via les liens profonds.

---

## 🚀 Déploiement sur Render

### Étape 1 — Préparer GitHub

Le repository doit contenir :

- `app.py`
- `hatvp.py`
- `requirements.txt`
- `render.yaml`
- `.python-version`
- `.streamlit/config.toml`

### Étape 2 — Déployer sur Render

1. Créez un compte sur [render.com](https://render.com)
2. Cliquez **"New +"** → **"Blueprint"** et sélectionnez le repo
3. Render lit le `render.yaml` et demande la valeur de `FORMSPREE_FORM_ID`
   (identifiant du formulaire [Formspree](https://formspree.io) qui reçoit les feedbacks)
4. Cliquez **"Deploy"** et attendez 2-3 minutes ⏳

### Étape 3 — Partager l'URL

Envoyez simplement l'URL à vos collègues. Pas d'installation nécessaire de leur côté.

---

## ⚠️ Notes d'hébergement

- Les données HATVP sont téléchargées au premier chargement puis **mises en cache 12h** :
  passé ce délai, le ZIP est retéléchargé (en cas d'échec, l'ancienne version reste utilisée)
- Le `render.yaml` monte un disque persistant pour ce cache (fonctionnalité des plans payants Render)
- Sur le plan gratuit, l'app **s'endort après 15 min** d'inactivité ; le réveil prend ~30-60 secondes

---

## 💻 Lancer en local

```bash
# 1. Installer les dépendances (Python 3.11)
pip install -r requirements.txt

# 2. (Optionnel) activer le formulaire de feedback
export FORMSPREE_FORM_ID=votre_id

# 3. Lancer l'app
streamlit run app.py

# L'app s'ouvre automatiquement sur http://localhost:8501
```

---

## 🌐 Site statique (`static_site/`)

Le générateur produit, à partir des données HATVP :

- une page par organisation (~3 900) : chiffres clés, dépenses déclarées par exercice, responsables publics
  contactés, domaines, types d'actions, dernières activités, dirigeants, clients / cabinets mandatés,
  export Excel (mêmes 4 onglets que l'app) ;
- une page par secteur (31), un index alphabétique, l'accueil avec recherche instantanée ;
- `sitemap.xml`, `robots.txt`, balises canonical / Open Graph / JSON-LD.

Les URL utilisent l'identifiant national (SIREN, RNA ou identifiant HATVP), stable d'un export à l'autre :
`/organisation/mouvement-des-entreprises-de-france-784668618/`.

### Générer en local

```bash
pip install -r static_site/requirements.txt
python static_site/build.py --limit 100          # prototype rapide
python static_site/build.py                      # site complet (~2-3 min)
python -m http.server -d static_site/dist 8000   # aperçu sur http://localhost:8000
```

Options : `--zip fichier.zip` (données locales au lieu du téléchargement), `--no-excel`.
Variables : `SITE_URL` (URL publique du site, `http://localhost:8000` par défaut ; un chemin éventuel
comme `https://orga.github.io/depot` préfixe tous les liens), `APP_URL` (URL de l'app Streamlit).

### Mise en ligne sur GitHub Pages (une seule fois)

Le workflow `.github/workflows/static-site.yml` génère le site chaque nuit et le publie sur GitHub Pages
(dépôt public requis pour le plan gratuit).

1. Dans le dépôt : **Settings → Pages → Build and deployment → Source : GitHub Actions**.
2. **Settings → Secrets and variables → Actions → Variables** : ajoutez `APP_URL`
   (URL de l'app Streamlit sur Render). L'URL du site est détectée automatiquement.
3. Lancez le workflow **Site statique** à la main (onglet Actions → Run workflow).
   Le site est publié sur `https://<orga>.github.io/<depot>/`
   (ou à la racine `https://<orga>.github.io/` si le dépôt s'appelle `<orga>.github.io`).
4. Déclarez le site dans [Google Search Console](https://search.google.com/search-console)
   et soumettez `<URL du site>/sitemap.xml`.

GitHub suspend les workflows planifiés après 60 jours sans activité sur le repo : un commit ou un
lancement manuel suffit à les réactiver.

---

## 📁 Structure du projet

```
Hatvp_query/
├── app.py                 # Interface Streamlit
├── hatvp.py               # Logique métier : chargement, recherche, jointures, export
├── requirements.txt       # Dépendances Python de l'app
├── render.yaml            # Config déploiement Render
├── .python-version        # Version Python
├── .streamlit/config.toml # Thème Streamlit
├── static_site/
│   ├── build.py           # Générateur du site statique
│   ├── requirements.txt   # Dépendances du générateur
│   ├── templates/         # Gabarits Jinja2
│   └── static/            # CSS et recherche instantanée
├── .github/workflows/static-site.yml  # Génération + publication GitHub Pages nocturnes
└── README.md              # Ce fichier
```

---

## 🔗 Sources des données

- Données : [data.gouv.fr — HATVP vues séparées](https://www.data.gouv.fr/datasets/repertoire-des-representants-dinterets-fichiers-en-vues-separees)
- Répertoire officiel : [hatvp.fr/le-repertoire](https://www.hatvp.fr/le-repertoire/)
- Licence : Licence Ouverte Etalab 2.0
