"""
HATVP To Table — Application Streamlit
Interface utilisateur : sidebar, en-tête, recherche, affichage, export.
"""

import html

import streamlit as st
import pandas as pd

from hatvp import (
    load_tables, CACHE_TTL,
    rule_groups,
    search_objets, search_secteurs, search_organisations,
    search_personnes, search_donneurs_ordre,
    enrich_actions,
    build_actions_sheet, build_orgs_sheet, build_persons_sheet, build_clients_sheet,
    build_excel,
    send_feedback_email,
    COL_SECTEUR, COL_OBJET, COL_REP_ID, COL_DENOM,
    COL_EXO_ID, COL_ACT_ID, COL_ARI_ID_APP,
)

# ─── CONFIG ───────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="HATVP To Table — Lobbying & transparence en France",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

load_all_tables = st.cache_resource(show_spinner=False, ttl=CACHE_TTL)(load_tables)

# Liens profonds depuis le site statique : ?mode=organisations&q=TotalEnergies&exact=1
MODES = {
    "objets":        "Objets d'activité",
    "secteurs":      "Secteurs d'activité",
    "organisations": "Organisations",
    "personnes":     "Personnes",
    "donneurs":      "Donneurs d'ordre",
}
# Chaque critère a un id stable : les clés de widgets ne doivent pas dépendre
# de la position, sinon supprimer une ligne décale les valeurs affichées.
if "kw_next_id" not in st.session_state:
    _qp = st.query_params
    st.session_state.kw_rules = [{"id": 0, "keyword": _qp.get("q", ""), "op": "ET"}]
    st.session_state.kw_next_id = 1
    st.session_state.default_mode_idx = (list(MODES).index(_qp["mode"])
                                         if _qp.get("mode") in MODES else 0)
    st.session_state.default_exact = _qp.get("exact") == "1"


def fr(n):
    return f"{n:,}".replace(",", "\u202f")


def show_metrics(items):
    for col, (label, value) in zip(st.columns(4), items):
        col.metric(label, fr(value))


def show_preview(title, items):
    with st.expander(title):
        li = "".join(f"<li>{html.escape(str(x))}</li>" for x in items)
        st.markdown(f"<ul class='preview'>{li}</ul>", unsafe_allow_html=True)

# ─── SEO ──────────────────────────────────────────────────────────────────────

# ⚠️ Mets à jour cette URL si tu configures un domaine personnalisé
SITE_URL = "https://hatvp-to-table.onrender.com"

st.markdown(f"""
<link rel="canonical" href="{SITE_URL}">
<meta name="description" content="Consultez le répertoire HATVP des lobbyistes en France : plus de 95 000 actions de lobbying déclarées, recherche par mot-clé, organisation ou personne, export Excel structuré. Données open data de la Haute Autorité pour la Transparence de la Vie Publique.">
<meta name="keywords" content="HATVP, lobbying, représentants d'intérêts, transparence, open data, répertoire, France, lobbyistes, actions de lobbying, export Excel">
<meta name="robots" content="index, follow">
<meta name="author" content="HATVP To Table">

<meta property="og:title" content="HATVP To Table — Répertoire des lobbyistes en France">
<meta property="og:description" content="Recherchez et téléchargez les données open data HATVP : 95 000 actions de lobbying déclarées, organisations, dirigeants — export Excel en un clic.">
<meta property="og:type" content="website">
<meta property="og:url" content="{SITE_URL}">
<meta property="og:locale" content="fr_FR">
<meta property="og:site_name" content="HATVP To Table">

<meta name="twitter:card" content="summary">
<meta name="twitter:title" content="HATVP To Table — Répertoire des lobbyistes en France">
<meta name="twitter:description" content="Recherchez et téléchargez les données HATVP sur le lobbying en France. Export Excel en un clic.">
<meta name="twitter:site" content="@hatvp">

<script type="application/ld+json">
{{
  "@context": "https://schema.org",
  "@type": "WebApplication",
  "name": "HATVP To Table",
  "url": "{SITE_URL}",
  "description": "Outil de recherche et d'export des données open data de la HATVP sur le lobbying en France. Permet de rechercher parmi plus de 95 000 actions de lobbying déclarées par les représentants d'intérêts inscrits au répertoire officiel.",
  "applicationCategory": "DataVisualization",
  "inLanguage": "fr",
  "isAccessibleForFree": true,
  "featureList": [
    "Recherche full-text dans les actions de lobbying",
    "Recherche par secteur, organisation, personne ou donneur d'ordre",
    "Export Excel structuré en 4 onglets",
    "Matching élargi ou exact"
  ],
  "about": {{
    "@type": "Dataset",
    "name": "Répertoire des représentants d'intérêts — HATVP",
    "url": "https://www.hatvp.fr/le-repertoire/",
    "license": "https://www.etalab.gouv.fr/licence-ouverte-open-licence",
    "publisher": {{
      "@type": "GovernmentOrganization",
      "name": "Haute Autorité pour la Transparence de la Vie Publique",
      "url": "https://www.hatvp.fr",
      "sameAs": "https://fr.wikipedia.org/wiki/Haute_Autorit%C3%A9_pour_la_transparence_de_la_vie_publique"
    }}
  }}
}}
</script>
""", unsafe_allow_html=True)

# ─── STYLES ───────────────────────────────────────────────────────────────────
# Les couleurs viennent du thème (.streamlit/config.toml) : le CSS ci-dessous
# reste neutre pour fonctionner aussi si l'utilisateur choisit le thème clair.

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&display=swap');
h1, h2, h3 { font-family: 'Syne', sans-serif !important; }
h1 { font-weight: 800 !important; letter-spacing: -0.5px; }
[data-testid="stSidebar"] { min-width: 380px !important; max-width: 380px !important; }
[data-testid="stSidebar"] > div:first-child { width: 380px !important; }
[data-testid="stMetricValue"] { font-family: 'Syne', sans-serif; font-weight: 800; }
[data-testid="stMetricLabel"] p { opacity: 0.65; }
.sector-pill { display: inline-block; border: 1px solid rgba(128,128,128,0.35); border-radius: 20px;
               padding: 4px 12px; margin: 3px; font-size: 14px; }
.sector-pill span { opacity: 0.55; margin-left: 4px; }
ul.preview { list-style: none; padding: 0; margin: 0; }
ul.preview li { padding: 8px 0; border-top: 1px solid rgba(128,128,128,0.2); }
ul.preview li:first-child { border-top: 0; }
</style>
""", unsafe_allow_html=True)

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🏛️ HATVP To Table")
    st.caption("Répertoire des représentants d'intérêts")

    FIRST_OPS = {"Inclure": "ET", "Exclure": "SAUF"}
    NEXT_OPS = ["ET", "OU", "SAUF"]

    st.markdown("**Mots-clés**", help=(
        "Combinez les critères avec **ET**, **OU** et **SAUF**. "
        "ET / SAUF sont prioritaires sur OU : "
        "*A ET B OU C SAUF D* = *(A ET B) OU (C SAUF D)*."))
    # Valeurs initiales posées dans session_state et non via value=/index= :
    # sous Streamlit 1.35 ces paramètres entrent dans l'identifiant du widget,
    # qui serait recréé (et la saisie perdue) dès que la valeur change.
    to_delete = None
    for i, rule in enumerate(st.session_state.kw_rules):
        rid = rule["id"]
        st.session_state.setdefault(f"kw_{rid}", rule["keyword"])
        c_op, c_kw, c_del = st.columns([2, 3, 0.6])
        with c_op:
            if i == 0:
                st.session_state.setdefault(f"op_first_{rid}", "Exclure" if rule["op"] == "SAUF" else "Inclure")
                choice = st.selectbox(
                    f"op{rid}", list(FIRST_OPS), key=f"op_first_{rid}",
                    label_visibility="collapsed")
                rule["op"] = FIRST_OPS[choice]
            else:
                st.session_state.setdefault(f"op_next_{rid}", rule["op"])
                rule["op"] = st.selectbox(
                    f"op{rid}", NEXT_OPS, key=f"op_next_{rid}",
                    label_visibility="collapsed")
        with c_kw:
            rule["keyword"] = st.text_input(
                f"kw{rid}", key=f"kw_{rid}",
                label_visibility="collapsed",
                placeholder="ex : énergie, MiCA...")
        with c_del:
            if len(st.session_state.kw_rules) > 1:
                if st.button("✕", key=f"del_{rid}", use_container_width=True):
                    to_delete = i
    if to_delete is not None:
        st.session_state.kw_rules.pop(to_delete)
        st.rerun()
    if st.button("＋ Ajouter un critère", use_container_width=True):
        st.session_state.kw_rules.append(
            {"id": st.session_state.kw_next_id, "keyword": "", "op": "ET"})
        st.session_state.kw_next_id += 1
        st.rerun()

    keyword_rules = [r for r in st.session_state.kw_rules if r["keyword"].strip()]
    groups = rule_groups(keyword_rules)
    group_labels = []
    for group in groups:
        terms = [("SAUF " if negate else ("ET " if j else "")) + f"« {kw} »"
                 for j, (kw, negate) in enumerate(group)]
        label = " ".join(terms)
        group_labels.append(f"({label})" if len(group) > 1 and len(groups) > 1 else label)
    keyword_label = " OU ".join(group_labels)
    keyword_slug = "_".join([kw for g in groups for kw, _ in g][:3]).replace(" ", "-")

    st.write("")
    mode = st.radio("Chercher dans",
        list(MODES.values()), index=st.session_state.default_mode_idx,
        help=(
            "**Objets** : full-text dans les descriptions des actions. "
            "**Secteurs** : 31 catégories. "
            "**Organisations** : nom, sigle. "
            "**Personnes** : nom/prénom des dirigeants et collaborateurs. "
            "**Donneurs d'ordre** : clients déclarés par les cabinets mandataires."
        ))
    mode_key = next(k for k, v in MODES.items() if v == mode)

    matching_mode = st.radio(
        "Correspondance",
        ["Élargi", "Exact"],
        index=1 if st.session_state.default_exact else 0,
        horizontal=True,
        help=(
            "**Élargi** — trouve les variantes et sous-chaînes. "
            "*« seb »* matche *SEB*, *Sebastian*, *Sébastien*...\n\n"
            "**Exact** — mot entier uniquement. "
            "*« seb »* matche *SEB SA* mais pas *Sebastian* ni *Sébastien*."
        )
    )
    exact_match = (matching_mode == "Exact")
    max_preview = st.slider("Lignes affichées", 5, 50, 15)

    st.divider()
    with st.popover("💬 Donner un feedback", use_container_width=True):
        st.markdown("##### Envoyer un retour")
        with st.form("feedback_form", clear_on_submit=True):
            fb_type = st.selectbox(
                "Type de retour",
                ["Suggestion", "Bug / Erreur", "Question", "Autre"])
            fb_msg = st.text_area(
                "Message *",
                placeholder="Décrivez votre retour, suggestion ou problème...",
                height=120)
            fb_email = st.text_input(
                "Votre email (optionnel)",
                placeholder="Pour un suivi éventuel")
            submitted = st.form_submit_button("Envoyer", use_container_width=True)
            if submitted:
                if not fb_msg.strip():
                    st.error("Le message ne peut pas être vide.")
                else:
                    ok, err = send_feedback_email(fb_type, fb_msg.strip(), fb_email.strip())
                    if ok:
                        st.success("Merci pour votre retour !")
                    else:
                        st.error(f"Erreur d'envoi : {err}")
    st.caption("Données open data · [hatvp.fr/le-repertoire](https://www.hatvp.fr/le-repertoire/) · mises à jour toutes les 12 h")

# ─── EN-TÊTE ──────────────────────────────────────────────────────────────────

st.title("🏛️ HATVP To Table")
st.caption("Recherchez dans le répertoire des lobbyistes de la "
           "[HATVP](https://www.hatvp.fr/le-repertoire/) et exportez les résultats en Excel.")

# ─── CHARGEMENT ───────────────────────────────────────────────────────────────

with st.spinner("⏳ Chargement des données HATVP... (première visite ~20 sec)"):
    try:
        tables = load_all_tables()
    except Exception as e:
        st.error(f"❌ Erreur de chargement : {e}")
        st.stop()

required_keys = ["infos", "secteurs", "objets", "exercices",
                 "dirigeants", "collaborateurs", "ari",
                 "actions", "ministeres", "domaines"]
missing = [k for k in required_keys if k not in tables]
if missing:
    st.error(f"❌ Tables manquantes : {missing}. Essayez --refresh-cache.")
    st.stop()

df_infos          = tables["infos"]
df_secteurs       = tables["secteurs"]
df_objets         = tables["objets"]
df_exercices      = tables["exercices"]
df_dirigeants     = tables["dirigeants"]
df_collaborateurs = tables["collaborateurs"]
df_ari            = tables["ari"]
df_actions        = tables["actions"]
df_ministeres     = tables["ministeres"]
df_domaines       = tables["domaines"]
df_clients        = tables.get("clients",       pd.DataFrame())
df_beneficiaires  = tables.get("beneficiaires", pd.DataFrame())

# ─── ACCUEIL ──────────────────────────────────────────────────────────────────

if not keyword_rules:
    st.write("")
    show_metrics([
        ("Représentants", len(df_infos)),
        ("Objets d'activité", len(df_objets)),
        ("Secteurs", df_secteurs[COL_SECTEUR].nunique()),
        ("Actions déclarées", len(df_actions)),
    ])
    st.write("")
    st.info("← Saisissez un mot-clé dans la barre latérale pour commencer.")

    if mode_key == "secteurs":
        st.subheader("Secteurs disponibles")
        all_s = df_secteurs[COL_SECTEUR].dropna().value_counts()
        pills = "".join(
            f'<span class="sector-pill">{html.escape(str(s))}<span>{n}</span></span>'
            for s, n in all_s.items())
        st.markdown(pills, unsafe_allow_html=True)

    st.write("")
    st.subheader("Comment ça marche")
    t_modes, t_criteres, t_export = st.tabs(["Modes de recherche", "Combiner des critères", "Export Excel"])
    with t_modes:
        st.markdown(f"""
- **Objets d'activité** *(recommandé)* : le texte des {fr(len(df_objets))} actions déclarées — `taxe carbone`, `MiCA`, `RGPD`
- **Secteurs d'activité** : les 31 secteurs déclarés — `Energie`, `Santé`, `Numérique`
- **Organisations** : nom ou sigle — `MEDEF`, `FNSEA`, `BNP`
- **Personnes** : dirigeants et collaborateurs — `Dupont` *(correspondance Exact conseillée)*
- **Donneurs d'ordre** : clients déclarés par les cabinets de lobbying — `EDF`, `Total`
""")
    with t_criteres:
        st.markdown("""
- **ET** : les deux termes doivent être présents · **OU** : l'un ou l'autre · **SAUF** : exclut le terme
- ET et SAUF passent avant OU : `nucléaire ET EDF OU hydrogène` = (nucléaire ET EDF) OU hydrogène
- **Élargi** trouve les variantes (`energie` → *énergétique*) ; **Exact** le mot entier (`loi` ≠ *lobbying*)
- Les accents et la casse sont ignorés
""")
    with t_export:
        st.markdown("""
Quatre onglets : **Actions** (objet, période, responsables publics, types d'actions), **Organisations**,
**Personnes** et **Clients** (pour les cabinets mandataires).

- En mode Personnes, l'onglet Personnes ne contient que les personnes trouvées
- En mode Donneurs d'ordre, l'onglet Clients ne contient que les donneurs d'ordre trouvés
- La période correspond à l'exercice déclaratif, pas à la date de publication
""")
    st.stop()

# ─── RECHERCHE ────────────────────────────────────────────────────────────────

st.subheader(f"Résultats pour {keyword_label}")
st.caption(f"{mode} · correspondance {matching_mode.lower()}")

ids_retenus            = []
df_objets_match        = pd.DataFrame()
_dirigeants_for_s3     = df_dirigeants
_collaborateurs_for_s3 = df_collaborateurs

# ── Mode OBJETS ───────────────────────────────────────────────────────────────
if mode_key == "objets":
    df_objets_match = search_objets(keyword_rules, df_objets, exact=exact_match)
    if df_objets_match.empty:
        st.warning(f"Aucun objet d'activité ne correspond à {keyword_label}.")
        st.stop()
    exo_rep = df_exercices[[COL_EXO_ID, COL_REP_ID]].drop_duplicates()
    ids_retenus = (df_objets_match.merge(exo_rep, on=COL_EXO_ID, how="left")
                   [COL_REP_ID].dropna().unique().tolist())
    show_metrics([("Actions trouvées", len(df_objets_match)),
                  ("Organisations", len(ids_retenus))])
    show_preview(f"Aperçu des {min(10, len(df_objets_match))} premiers objets",
                 df_objets_match[COL_OBJET].head(10))

# ── Mode SECTEURS ─────────────────────────────────────────────────────────────
elif mode_key == "secteurs":
    all_sectors_found = search_secteurs(keyword_rules, df_secteurs, exact=exact_match)
    if not all_sectors_found:
        st.warning(f"Aucun secteur ne correspond à {keyword_label}.")
        st.stop()
    val_counts = df_secteurs[COL_SECTEUR].value_counts()
    selected_sectors = st.multiselect(
        "Secteurs trouvés",
        options=all_sectors_found, default=all_sectors_found,
        format_func=lambda s: f"{s}  ({val_counts.get(s, 0)})")
    if not selected_sectors:
        st.info("Sélectionnez au moins un secteur.")
        st.stop()
    ids_retenus = (df_secteurs[df_secteurs[COL_SECTEUR].isin(selected_sectors)]
                   [COL_REP_ID].dropna().unique().tolist())
    exo_ids = df_exercices[df_exercices[COL_REP_ID].isin(ids_retenus)][COL_EXO_ID].unique()
    df_objets_match = df_objets[df_objets[COL_EXO_ID].isin(exo_ids)]
    show_metrics([("Organisations", len(ids_retenus)),
                  ("Actions associées", len(df_objets_match))])

# ── Mode ORGANISATIONS ────────────────────────────────────────────────────────
elif mode_key == "organisations":
    df_orgs_match = search_organisations(keyword_rules, df_infos, exact=exact_match)
    if df_orgs_match.empty:
        st.warning(f"Aucune organisation ne correspond à {keyword_label}.")
        st.stop()
    ids_retenus = df_orgs_match[COL_REP_ID].dropna().unique().tolist()
    exo_ids = df_exercices[df_exercices[COL_REP_ID].isin(ids_retenus)][COL_EXO_ID].unique()
    df_objets_match = df_objets[df_objets[COL_EXO_ID].isin(exo_ids)]
    show_metrics([("Organisations trouvées", len(ids_retenus)),
                  ("Actions associées", len(df_objets_match))])
    show_preview(f"Aperçu des {min(10, len(df_orgs_match))} premières organisations",
                 [str(r.get(COL_DENOM, "")) or str(r.get("nom_usage_hatvp", ""))
                  for _, r in df_orgs_match.head(10).iterrows()])

# ── Mode PERSONNES ────────────────────────────────────────────────────────────
elif mode_key == "personnes":
    df_pers_match = search_personnes(keyword_rules, df_dirigeants, df_collaborateurs, exact=exact_match)
    if df_pers_match.empty:
        st.warning(f"Aucune personne ne correspond à {keyword_label}.")
        st.stop()
    ids_retenus = df_pers_match[COL_REP_ID].dropna().unique().tolist()
    exo_ids = df_exercices[df_exercices[COL_REP_ID].isin(ids_retenus)][COL_EXO_ID].unique()
    df_objets_match = df_objets[df_objets[COL_EXO_ID].isin(exo_ids)]
    _dirigeants_for_s3 = (df_pers_match[df_pers_match["_statut_match"] == "Dirigeant"]
                          .drop(columns=["_statut_match"]))
    _collaborateurs_for_s3 = (df_pers_match[df_pers_match["_statut_match"] == "Collaborateur"]
                               .drop(columns=["_statut_match"]))
    show_metrics([("Personnes trouvées", len(df_pers_match)),
                  ("Organisations", len(ids_retenus)),
                  ("Actions associées", len(df_objets_match))])

    def _person_label(row):
        nom_col = next((c for c in ["nom_prenom_dirigeant", "nom_prenom_collaborateur",
                                    "nom_dirigeant", "nom_collaborateur"] if c in row.index), None)
        return f"{row[nom_col] if nom_col else ''} — {row.get('_statut_match', '')}"
    show_preview(f"Aperçu des {min(10, len(df_pers_match))} premières personnes",
                 [_person_label(r) for _, r in df_pers_match.head(10).iterrows()])

# ── Mode DONNEURS D'ORDRE ─────────────────────────────────────────────────────
elif mode_key == "donneurs":
    df_do_match = search_donneurs_ordre(keyword_rules, df_clients, exact=exact_match)
    if df_do_match.empty:
        st.warning(f"Aucun donneur d'ordre ne correspond à {keyword_label}.")
        st.stop()

    matched_client_names = set(df_do_match["denomination_client"].dropna().unique())
    ids_retenus = df_do_match["representants_id"].dropna().unique().tolist()

    if (not df_beneficiaires.empty
            and COL_ARI_ID_APP in df_beneficiaires.columns
            and "action_menee_en_propre" in df_beneficiaires.columns
            and "beneficiaire_action_menee" in df_beneficiaires.columns):
        df_tiers_do = df_beneficiaires[
            (df_beneficiaires["action_menee_en_propre"].astype(str) == "0") &
            (df_beneficiaires["beneficiaire_action_menee"].isin(matched_client_names))
        ]
        ari_ids_do = df_tiers_do[COL_ARI_ID_APP].dropna().unique()
        if len(ari_ids_do) and not df_ari.empty and COL_ACT_ID in df_ari.columns:
            act_ids_do = df_ari[df_ari[COL_ARI_ID_APP].isin(ari_ids_do)][COL_ACT_ID].dropna().unique()
            df_objets_match = df_objets[df_objets[COL_ACT_ID].isin(act_ids_do)]
        else:
            df_objets_match = pd.DataFrame(columns=df_objets.columns)
    else:
        exo_ids = df_exercices[df_exercices[COL_REP_ID].isin(ids_retenus)][COL_EXO_ID].unique()
        df_objets_match = df_objets[df_objets[COL_EXO_ID].isin(exo_ids)]

    nb_do = len(matched_client_names)
    show_metrics([("Donneurs d'ordre trouvés", nb_do),
                  ("Cabinets mandataires", len(ids_retenus)),
                  ("Actions associées", len(df_objets_match))])
    show_preview(f"Aperçu des {min(10, nb_do)} premiers donneurs d'ordre",
                 sorted(matched_client_names)[:10])

# ─── ENRICHISSEMENT ───────────────────────────────────────────────────────────

if df_objets_match.empty:
    st.warning("Aucune action trouvée.")
    st.stop()

with st.spinner("🔀 Enrichissement des actions (période, types, responsables)..."):
    df_enriched = enrich_actions(
        df_objets_match, df_exercices,
        df_ari, df_actions, df_ministeres, df_domaines, df_beneficiaires)
    if COL_REP_ID in df_enriched.columns:
        df_enriched = df_enriched[df_enriched[COL_REP_ID].isin(ids_retenus)]

# ─── CONSTRUCTION ONGLETS ─────────────────────────────────────────────────────

df_s1 = build_actions_sheet(df_enriched, df_infos)
df_s2 = build_orgs_sheet(ids_retenus, df_infos, df_enriched)
df_s3 = build_persons_sheet(ids_retenus, df_s2, _dirigeants_for_s3, _collaborateurs_for_s3)

if mode_key in ("objets", "secteurs"):
    _ari_ids_matched = (df_enriched[COL_ARI_ID_APP].dropna().unique()
                        if COL_ARI_ID_APP in df_enriched.columns else None)
    _allowed_clients = None
elif mode_key == "donneurs":
    _ari_ids_matched = None
    _allowed_clients = matched_client_names
else:
    _ari_ids_matched = None
    _allowed_clients = None
df_s4 = build_clients_sheet(ids_retenus, df_clients, df_infos,
                             ari_ids=_ari_ids_matched, df_beneficiaires=df_beneficiaires,
                             allowed_clients=_allowed_clients)
_df_s2_export = df_s2.drop(columns=[COL_REP_ID], errors="ignore")

# ─── EXPORT ───────────────────────────────────────────────────────────────────

excel_bytes = build_excel([
    ("Actions de lobbying",         df_s1),
    ("Organisations",               _df_s2_export),
    ("Dirigeants & Collaborateurs", df_s3),
    ("Clients & Mandats",           df_s4),
])
n3 = len(df_s3) if not df_s3.empty else 0
n4 = len(df_s4) if not df_s4.empty else 0
st.download_button(
    label="⬇️  Télécharger l'Excel",
    data=excel_bytes,
    file_name=f"hatvp_{keyword_slug}_{mode_key}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# ─── AFFICHAGE ONGLETS ────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4 = st.tabs([
    f"Actions ({fr(len(df_s1))})",
    f"Organisations ({fr(len(df_s2))})",
    f"Personnes ({fr(n3)})",
    f"Clients ({fr(n4)})",
])

RENAME = {
    COL_REP_ID: "ID", COL_DENOM: "Organisation", "nom_usage_hatvp": "Nom HATVP",
    "sigle_hatvp": "Sigle", "label_categorie_organisation": "Catégorie",
    "ville": "Ville", "pays": "Pays", "site_web": "Site web",
    "page_linkedin": "LinkedIn", "page_twitter": "Twitter",
    "date_premiere_publication": "1ère publication",
    "objets_activite_matches": "Objets matchés",
    "identifiant_national": "SIREN/RNA",
    "statut": "Statut", "civilite": "Civilité", "nom": "Nom",
    "prenom": "Prénom", "fonction": "Fonction", "nom_prenom": "Nom complet",
}

def show_tab(df, label):
    if df is None or df.empty:
        st.info(f"Aucune donnée {label}.")
        return
    df_disp = df.head(max_preview).reset_index(drop=True)
    df_disp = df_disp.rename(columns={k: v for k, v in RENAME.items() if k in df_disp.columns})
    st.dataframe(df_disp, use_container_width=True, height=420, hide_index=True)
    if len(df) > max_preview:
        st.caption(f"{max_preview} lignes affichées sur {fr(len(df))} · l'Excel contient tout.")

with tab1:
    show_tab(df_s1, "d'actions")

with tab2:
    show_tab(_df_s2_export, "d'organisations")

with tab3:
    show_tab(df_s3, "de personnes")

with tab4:
    if df_s4.empty:
        st.info("Aucun client déclaré — les organisations trouvées ne sont pas des cabinets mandataires.")
    else:
        show_tab(df_s4, "de clients")
