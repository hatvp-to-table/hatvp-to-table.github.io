"""
HATVP To Table — Logique métier
Chargement, matching, enrichissement, construction des onglets, export Excel, feedback.
Indépendant de Streamlit : utilisé par l'app (app.py) et par le générateur du site statique (static_site/build.py).
"""

import re
import gzip
import io
import os
import time
import zipfile
import requests
import tempfile
import unicodedata
from pathlib import Path
from io import BytesIO

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# ─── CONSTANTES ───────────────────────────────────────────────────────────────

CACHE_DIR = Path(tempfile.gettempdir()) / "hatvp_cache"
CACHE_TTL = 12 * 3600  # secondes
ZIP_URL   = "https://www.hatvp.fr/agora/opendata/csv/Vues_Separees_CSV.zip"

COL_SECTEUR    = "secteur_activite"
COL_OBJET      = "objet_activite"
COL_REP_ID     = "representants_id"
COL_DENOM      = "denomination"
COL_EXO_ID     = "exercices_id"
COL_ACT_ID     = "activite_id"
COL_ARI_ID_APP = "action_representation_interet_id"

TABLE_KEYWORDS = {
    "secteurs":       ["secteur"],
    "infos":          ["information", "generale"],
    "objets":         ["objet", "activite"],
    "exercices":      ["exercice"],
    "dirigeants":     ["dirigeant"],
    "collaborateurs": ["collaborateur"],
    "ari":            ["observation"],   # 14_observations.csv : activite_id ↔ action_representation_interet_id
    "actions":        ["action", "menee"],
    "ministeres":     ["ministere"],
    "domaines":       ["domaine"],
    "decisions":      ["decision"],
    "clients":        ["client"],
    "beneficiaires":  ["beneficiaire"],
}

def _nc(c):
    return c.strip().lower().replace(" ", "_").strip('"').strip("'")

def _col_filter(fixed=(), patterns=(), exclude=()):
    fixed_set = set(fixed)
    def keep(c):
        nc = _nc(c)
        if exclude and any(e in nc for e in exclude):
            return False
        return nc in fixed_set or any(p in nc for p in patterns)
    return keep

TABLE_COLS = {
    "secteurs":       _col_filter({"secteur_activite", "representants_id"}),
    "infos":          _col_filter({"representants_id", "denomination", "nom_usage_hatvp",
                                   "sigle_hatvp", "label_categorie_organisation", "adresse",
                                   "code_postal", "ville", "pays", "site_web", "page_linkedin",
                                   "page_twitter", "date_premiere_publication",
                                   "identifiant_national", "type_identifiant_national"}),
    "objets":         _col_filter({"objet_activite", "exercices_id", "activite_id"}),
    "exercices":      _col_filter({"exercices_id", "representants_id",
                                   "annee_debut", "annee_fin"}),
    "dirigeants":     _col_filter({"representants_id", "civilite_dirigeant", "nom_dirigeant",
                                   "prenom_dirigeant", "fonction_dirigeant",
                                   "nom_prenom_dirigeant"}),
    "collaborateurs": _col_filter({"representants_id", "civilite_collaborateur",
                                   "nom_collaborateur", "prenom_collaborateur",
                                   "fonction_collaborateur", "nom_prenom_collaborateur"}),
    "ari":            _col_filter({"activite_id", "action_representation_interet_id"}),
    "actions":        _col_filter({"action_representation_interet_id"}, ("action_menee",),
                                  exclude=("autre",)),
    "ministeres":     _col_filter({"action_representation_interet_id"}, ("responsable_public",),
                                  exclude=("autre",)),
    "domaines":       _col_filter({"activite_id"}, ("domaine",)),
    "clients":        _col_filter({"representants_id", "denomination_client",
                                   "identifiant_national_client", "ancienclient",
                                   "datecessation"}),
    "beneficiaires":  _col_filter({"beneficiaire_action_menee",
                                   "action_representation_interet_id",
                                   "action_menee_en_propre"}),
}

# Colonnes recherchées : une version normalisée (minuscules, sans accents) est
# précalculée au chargement pour éviter de renormaliser à chaque requête.
SEARCH_COLS = {
    "objets":         [COL_OBJET],
    "infos":          ["denomination", "nom_usage_hatvp", "sigle_hatvp"],
    "dirigeants":     ["nom_dirigeant", "prenom_dirigeant", "nom_prenom_dirigeant"],
    "collaborateurs": ["nom_collaborateur", "prenom_collaborateur", "nom_prenom_collaborateur"],
    "clients":        ["denomination_client"],
}

# ─── UTILITAIRES ──────────────────────────────────────────────────────────────

def find_csv_in_zip(csv_names, keywords):
    kw_low = [k.lower() for k in keywords]
    candidates = [n for n in csv_names if all(k in n.lower() for k in kw_low)]
    return max(candidates, key=len) if candidates else None

def normalize_cols(df):
    df.columns = [_nc(c) for c in df.columns]
    return df

def read_csv_bytes(raw_bytes, usecols=None):
    if raw_bytes[:2] == b"\x1f\x8b":
        with gzip.open(io.BytesIO(raw_bytes)) as f:
            raw_bytes = f.read()
    for enc in ["utf-8-sig", "utf-8", "latin-1", "cp1252"]:
        try:
            text = raw_bytes.decode(enc); break
        except UnicodeDecodeError:
            continue
    else:
        text = raw_bytes.decode("latin-1", errors="replace")
    first_line = text.split("\n")[0]
    sep = ";" if first_line.count(";") >= first_line.count(",") else ","
    for s in [sep, ("," if sep == ";" else ";")]:
        try:
            df = pd.read_csv(io.StringIO(text), sep=s, low_memory=False,
                             on_bad_lines="skip", usecols=usecols)
            if len(df.columns) > 1:
                return normalize_cols(df)
        except Exception:
            continue
    return pd.DataFrame()

def strip_accents(text):
    return "".join(c for c in unicodedata.normalize("NFD", text)
                   if unicodedata.category(c) != "Mn")

def normalize(text):
    return strip_accents(str(text).lower())

def norm_col(col):
    return f"_norm_{col}"

def add_normalized_cols(df, cols):
    for c in cols:
        if c in df.columns:
            df[norm_col(c)] = df[c].fillna("").map(normalize)
    return df

# ─── CHARGEMENT DONNÉES ───────────────────────────────────────────────────────

def _download_zip(dest):
    tmp = dest.with_suffix(".part")
    resp = requests.get(ZIP_URL, stream=True, timeout=300)
    resp.raise_for_status()
    with open(tmp, "wb") as f:
        for chunk in resp.iter_content(65536):
            f.write(chunk)
    tmp.replace(dest)

def ensure_zip():
    """Retourne le chemin du ZIP HATVP en cache, retéléchargé s'il a plus de CACHE_TTL."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    zip_cache = CACHE_DIR / "vues_separees.zip"
    is_fresh = zip_cache.exists() and time.time() - zip_cache.stat().st_mtime < CACHE_TTL
    if not is_fresh:
        try:
            _download_zip(zip_cache)
        except requests.RequestException:
            # HATVP injoignable : on garde l'ancienne version plutôt que de planter
            if not zip_cache.exists():
                raise
    return zip_cache

def read_tables(zip_path, table_cols=None):
    table_cols = TABLE_COLS if table_cols is None else table_cols
    tables = {}
    with zipfile.ZipFile(zip_path) as z:
        names = z.namelist()
        csv_names = [n for n in names if n.lower().endswith(".csv")]
        for key, keywords in TABLE_KEYWORDS.items():
            match = find_csv_in_zip(csv_names, keywords)
            if match:
                df = read_csv_bytes(z.read(match), usecols=table_cols.get(key))
                tables[key] = add_normalized_cols(df, SEARCH_COLS.get(key, []))
    return tables

def load_tables():
    return read_tables(ensure_zip())

# ─── MATCHING ─────────────────────────────────────────────────────────────────

_WORD_CHARS = r"\w\u00C0-\u024F"

def _keyword_pattern(keyword, exact):
    kw = normalize(keyword.strip())
    if exact:
        return rf"(?<![{_WORD_CHARS}]){re.escape(kw)}(?![{_WORD_CHARS}])"
    variants = {kw}
    if len(kw) > 6: variants.add(kw[:-1])
    if len(kw) > 7: variants.add(kw[:-2])
    return "|".join(re.escape(v) for v in sorted(variants, key=len, reverse=True))

def _normalized_series(df, col):
    nc = norm_col(col)
    if nc in df.columns:
        return df[nc]
    return df[col].fillna("").map(normalize)

def rule_groups(rules):
    """Découpe les règles en groupes séparés par OU.
    Retourne [[(mot_clé, négation), ...], ...] ; ET/SAUF sont prioritaires sur OU."""
    groups = [[]]
    for r in rules:
        kw = r["keyword"].strip()
        if not kw:
            continue
        if r["op"] == "OU" and groups[-1]:
            groups.append([])
        groups[-1].append((kw, r["op"] in ("SAUF", "Exclure")))
    return [g for g in groups if g]

def _df_matches_rules(df, cols, rules, exact):
    series = [_normalized_series(df, c) for c in cols]

    def kw_mask(kw):
        pattern = _keyword_pattern(kw, exact)
        mask = pd.Series(False, index=df.index)
        for s in series:
            mask |= s.str.contains(pattern, regex=True)
        return mask

    groups = rule_groups(rules)
    if not groups:
        return pd.Series(True, index=df.index)
    result = pd.Series(False, index=df.index)
    for group in groups:
        group_mask = pd.Series(True, index=df.index)
        for kw, negate in group:
            m = kw_mask(kw)
            group_mask &= ~m if negate else m
        result |= group_mask
    return result

def search_secteurs(rules, df_secteurs, exact=False):
    sectors = df_secteurs[[COL_SECTEUR]].dropna().drop_duplicates()
    mask = _df_matches_rules(sectors, [COL_SECTEUR], rules, exact)
    return sorted(sectors.loc[mask, COL_SECTEUR].unique().tolist())

def search_objets(rules, df_objets, exact=False):
    mask = _df_matches_rules(df_objets, [COL_OBJET], rules, exact)
    return df_objets.loc[mask]

def search_organisations(rules, df_infos, exact=False):
    search_cols = [c for c in ["denomination", "nom_usage_hatvp", "sigle_hatvp"]
                   if c in df_infos.columns]
    mask = _df_matches_rules(df_infos, search_cols, rules, exact)
    return df_infos.loc[mask]

def search_donneurs_ordre(rules, df_clients, exact=False):
    if "denomination_client" not in df_clients.columns or df_clients.empty:
        return pd.DataFrame()
    mask = _df_matches_rules(df_clients, ["denomination_client"], rules, exact)
    return df_clients.loc[mask]

def search_personnes(rules, df_dirigeants, df_collaborateurs, exact=False):
    frames = []
    for df_src, statut, name_cols in [
        (df_dirigeants, "Dirigeant",
         ["nom_dirigeant", "prenom_dirigeant", "nom_prenom_dirigeant"]),
        (df_collaborateurs, "Collaborateur",
         ["nom_collaborateur", "prenom_collaborateur", "nom_prenom_collaborateur"]),
    ]:
        if df_src.empty: continue
        cols = [c for c in name_cols if c in df_src.columns]
        if not cols: continue
        mask = _df_matches_rules(df_src, cols, rules, exact)
        matched = df_src.loc[mask].copy()
        matched["_statut_match"] = statut
        frames.append(matched)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)

# ─── ENRICHISSEMENT ───────────────────────────────────────────────────────────

def _agg(df, group_col, val_col, out_name):
    return (df[df[val_col].notna() & (df[val_col].astype(str) != "nan")]
            .groupby(group_col)[val_col]
            .apply(lambda x: " | ".join(sorted(set(str(v) for v in x))))
            .reset_index().rename(columns={val_col: out_name}))

def enrich_actions(df_objets_matched, df_exercices,
                   df_ari, df_actions, df_ministeres, df_domaines, df_beneficiaires=None):
    """Jointures via 14_observations.csv comme table pivot."""
    if df_beneficiaires is None:
        df_beneficiaires = pd.DataFrame()

    # 1. objets → exercices
    exo_cols = ["exercices_id", "representants_id"]
    for c in ["annee_debut", "annee_fin"]:
        if c in df_exercices.columns:
            exo_cols.append(c)
    df = df_objets_matched.merge(
        df_exercices[exo_cols].drop_duplicates(), on="exercices_id", how="left")

    # 2. objets → observations (activite_id → action_representation_interet_id)
    if not df_ari.empty and "activite_id" in df_ari.columns and COL_ARI_ID_APP in df_ari.columns:
        df = df.merge(df_ari[["activite_id", COL_ARI_ID_APP]].drop_duplicates(),
                      on="activite_id", how="left")
    else:
        df[COL_ARI_ID_APP] = None

    # 3. ari_id → types d'actions
    if not df_actions.empty and COL_ARI_ID_APP in df_actions.columns:
        acol = next((c for c in df_actions.columns if "action_menee" in c and "autre" not in c), None)
        if acol:
            df = df.merge(_agg(df_actions, COL_ARI_ID_APP, acol, "types_actions"),
                          on=COL_ARI_ID_APP, how="left")
        else:
            df["types_actions"] = None
    else:
        df["types_actions"] = None

    # 4. ari_id → responsables publics
    if not df_ministeres.empty and COL_ARI_ID_APP in df_ministeres.columns:
        rcol = next((c for c in df_ministeres.columns
                     if "responsable_public" in c and "autre" not in c), None)
        if rcol:
            df = df.merge(_agg(df_ministeres, COL_ARI_ID_APP, rcol, "responsables_publics"),
                          on=COL_ARI_ID_APP, how="left")
        else:
            df["responsables_publics"] = None
    else:
        df["responsables_publics"] = None

    # 5. activite_id → domaines (direct)
    if not df_domaines.empty and "activite_id" in df_domaines.columns:
        dcol = next((c for c in df_domaines.columns if "domaine" in c), None)
        if dcol:
            df = df.merge(_agg(df_domaines, "activite_id", dcol, "domaines_intervention"),
                          on="activite_id", how="left")
        else:
            df["domaines_intervention"] = None
    else:
        df["domaines_intervention"] = None

    # 6. ari_id → donneur d'ordre (bénéficiaire quand action faite pour un tiers)
    if (not df_beneficiaires.empty
            and COL_ARI_ID_APP in df_beneficiaires.columns
            and "action_menee_en_propre" in df_beneficiaires.columns):
        df_tiers = df_beneficiaires[
            df_beneficiaires["action_menee_en_propre"].astype(str) == "0"
        ]
        if not df_tiers.empty:
            df = df.merge(
                _agg(df_tiers, COL_ARI_ID_APP, "beneficiaire_action_menee", "donneur_ordre"),
                on=COL_ARI_ID_APP, how="left")
        else:
            df["donneur_ordre"] = None
    else:
        df["donneur_ordre"] = None

    return df

# ─── CONSTRUCTION DES ONGLETS ─────────────────────────────────────────────────

def build_actions_sheet(df_enriched, df_infos):
    org_cols = [c for c in [COL_REP_ID, COL_DENOM, "nom_usage_hatvp",
                             "label_categorie_organisation", "ville", "site_web"]
                if c in df_infos.columns]
    df = df_enriched.merge(
        df_infos[org_cols].drop_duplicates(subset=[COL_REP_ID]),
        on=COL_REP_ID, how="left")
    wanted = [COL_DENOM, "label_categorie_organisation", "ville",
              COL_OBJET, "annee_debut", "annee_fin",
              "donneur_ordre", "types_actions", "responsables_publics", "domaines_intervention"]
    cols = list(dict.fromkeys(c for c in wanted if c in df.columns))
    rename = {
        COL_DENOM:                    "Organisation",
        "nom_usage_hatvp":            "Nom HATVP",
        "label_categorie_organisation":"Catégorie",
        "ville":                      "Ville",
        COL_REP_ID:                   "ID Organisation",
        COL_OBJET:                    "Objet de l'action",
        "annee_debut":                "Période début",
        "annee_fin":                  "Période fin",
        "types_actions":              "Types d'actions",
        "responsables_publics":       "Responsables publics contactés",
        "donneur_ordre":              "Donneur d'ordre",
        "domaines_intervention":      "Domaines d'intervention",
        "identifiant_fiche":          "ID Fiche HATVP",
        COL_ACT_ID:                   "ID Activité",
    }
    df = df[cols].rename(columns={k: v for k, v in rename.items() if k in cols})
    sort_col = "Organisation" if "Organisation" in df.columns else df.columns[0]
    return df.sort_values([sort_col, "Période début"] if "Période début" in df.columns
                          else [sort_col], na_position="last")

def build_orgs_sheet(ids, df_infos, df_enriched):
    df_summary = (
        df_enriched.groupby(COL_REP_ID)[COL_OBJET]
        .apply(lambda x: " | ".join(sorted(set(str(v)[:120] for v in x.dropna()))))
        .reset_index().rename(columns={COL_OBJET: "objets_activite_matches"})
    )
    df = df_infos[df_infos[COL_REP_ID].isin(ids)].merge(df_summary, on=COL_REP_ID, how="left")
    wanted = [COL_REP_ID, COL_DENOM, "nom_usage_hatvp", "sigle_hatvp",
              "label_categorie_organisation", "adresse", "code_postal", "ville", "pays",
              "site_web", "page_linkedin", "page_twitter",
              "objets_activite_matches", "date_premiere_publication",
              "identifiant_national", "type_identifiant_national"]
    cols = list(dict.fromkeys(c for c in wanted if c in df.columns))
    return (df[cols].drop_duplicates(subset=[COL_REP_ID])
            .sort_values(COL_DENOM if COL_DENOM in cols else cols[0], na_position="last"))

def build_persons_sheet(ids, df_orgs, df_dirigeants, df_collaborateurs):
    org_ref_cols = [c for c in [COL_REP_ID, COL_DENOM, "nom_usage_hatvp",
                                 "label_categorie_organisation", "ville", "site_web",
                                 "objets_activite_matches"] if c in df_orgs.columns]
    df_org = df_orgs[org_ref_cols].copy()
    frames = []
    for df_src, statut, rmap in [
        (df_dirigeants, "Dirigeant", {
            "civilite_dirigeant": "civilite", "nom_dirigeant": "nom",
            "prenom_dirigeant": "prenom", "fonction_dirigeant": "fonction",
            "nom_prenom_dirigeant": "nom_prenom"}),
        (df_collaborateurs, "Collaborateur", {
            "civilite_collaborateur": "civilite", "nom_collaborateur": "nom",
            "prenom_collaborateur": "prenom", "fonction_collaborateur": "fonction",
            "nom_prenom_collaborateur": "nom_prenom"}),
    ]:
        if df_src.empty or COL_REP_ID not in df_src.columns:
            continue
        df_f = df_src[df_src[COL_REP_ID].isin(ids)].copy()
        if df_f.empty:
            continue
        df_f["statut"] = statut
        df_f = df_f.rename(columns={k: v for k, v in rmap.items() if k in df_f.columns})
        frames.append(df_f)
    if not frames:
        return pd.DataFrame()
    df_p = pd.concat(frames, ignore_index=True).merge(df_org, on=COL_REP_ID, how="left")
    wanted = ["statut", "civilite", "nom", "prenom", "fonction",
              COL_DENOM, "label_categorie_organisation", "ville", "site_web", "objets_activite_matches"]
    cols = list(dict.fromkeys(c for c in wanted if c in df_p.columns))
    return df_p[cols].sort_values([COL_DENOM, "statut", "nom"], na_position="last")

def build_clients_sheet(ids, df_clients, df_infos, ari_ids=None, df_beneficiaires=None,
                        allowed_clients=None):
    if df_clients.empty or "representants_id" not in df_clients.columns:
        return pd.DataFrame()
    df = df_clients[df_clients["representants_id"].isin(ids)].copy()
    if df.empty:
        return pd.DataFrame()
    if allowed_clients is not None and "denomination_client" in df.columns:
        df = df[df["denomination_client"].isin(allowed_clients)]
    if (ari_ids is not None and df_beneficiaires is not None
            and not df_beneficiaires.empty
            and COL_ARI_ID_APP in df_beneficiaires.columns
            and "action_menee_en_propre" in df_beneficiaires.columns
            and "beneficiaire_action_menee" in df_beneficiaires.columns):
        df_tiers = df_beneficiaires[
            (df_beneficiaires[COL_ARI_ID_APP].isin(ari_ids)) &
            (df_beneficiaires["action_menee_en_propre"].astype(str) == "0")
        ]
        if not df_tiers.empty:
            matched_clients = set(df_tiers["beneficiaire_action_menee"].dropna().unique())
            df = df[df["denomination_client"].isin(matched_clients)]
    if df.empty:
        return pd.DataFrame()
    cab_cols = [c for c in ["representants_id", "denomination", "nom_usage_hatvp",
                             "label_categorie_organisation"] if c in df_infos.columns]
    df = df.merge(df_infos[cab_cols].drop_duplicates("representants_id"),
                  on="representants_id", how="left")
    wanted = ["denomination", "nom_usage_hatvp", "label_categorie_organisation",
              "denomination_client", "identifiant_national_client",
              "ancienclient", "datecessation"]
    cols = [c for c in wanted if c in df.columns]
    rename = {
        "denomination":               "Cabinet mandataire",
        "nom_usage_hatvp":            "Nom HATVP cabinet",
        "label_categorie_organisation": "Catégorie cabinet",
        "representants_id":           "ID Cabinet",
        "denomination_client":        "Client (donneur d'ordre)",
        "identifiant_national_client": "SIREN client",
        "ancienclient":               "Ancien client",
        "datecessation":              "Date cessation",
    }
    df = df[cols].rename(columns={k: v for k, v in rename.items() if k in cols})
    sort_col = "Cabinet mandataire" if "Cabinet mandataire" in df.columns else df.columns[0]
    return df.sort_values(sort_col, na_position="last")

# ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────

def build_excel(sheets):
    """sheets = liste de (nom, df). Retourne bytes."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets:
            if df is None or df.empty:
                continue
            # Certaines déclarations contiennent des caractères de contrôle refusés par Excel
            text_cols = df.select_dtypes(include="object").columns
            df = df.copy()
            df[text_cols] = df[text_cols].apply(lambda s: s.map(
                lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v))
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            for col_idx, col_name in enumerate(df.columns, 1):
                max_len = max(len(str(col_name)),
                              int(df[col_name].astype(str).str.len().quantile(0.9))
                              if len(df) > 0 else 0)
                ws.column_dimensions[
                    ws.cell(row=1, column=col_idx).column_letter
                ].width = min(max_len + 2, 60)
    return buf.getvalue()

# ─── FEEDBACK ─────────────────────────────────────────────────────────────────

def send_feedback_email(fb_type, message, user_email=""):
    form_id = os.environ.get("FORMSPREE_FORM_ID", "")
    if not form_id:
        return False, "Formspree non configuré (variable FORMSPREE_FORM_ID manquante)"
    body = f"Type : {fb_type}\n"
    if user_email:
        body += f"Email utilisateur : {user_email}\n"
    body += f"\nMessage :\n{message}"
    payload = {
        "subject": f"[HATVP To Table] {fb_type}",
        "message": body,
    }
    if user_email:
        payload["_replyto"] = user_email
    try:
        r = requests.post(
            f"https://formspree.io/f/{form_id}",
            json=payload,
            headers={"Accept": "application/json"},
            timeout=10,
        )
        if r.status_code == 200:
            return True, ""
        return False, f"Erreur Formspree : {r.status_code}"
    except Exception as e:
        return False, str(e)
