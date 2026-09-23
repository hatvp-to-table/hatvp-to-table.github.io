"""
HATVP To Table — Générateur du site statique

Produit une page HTML par organisation et par secteur du répertoire HATVP,
plus l'accueil, les index, le sitemap et un export Excel par organisation.
La recherche avancée reste dans l'app Streamlit (liens profonds ?mode=…&q=…).

Usage :
    python static_site/build.py                    # télécharge les données HATVP
    python static_site/build.py --zip donnees.zip  # utilise un ZIP local
    python static_site/build.py --limit 100        # prototype rapide (100 organisations)
    python static_site/build.py --no-excel         # sans les exports Excel

Variables d'environnement :
    SITE_URL  URL publique du site statique (balises canonical, sitemap). Son chemin
              éventuel (https://orga.github.io/depot) préfixe tous les liens internes.
    APP_URL   URL de l'app Streamlit
"""

import argparse
import json
import os
import re
import shutil
import sys
import time
from collections import Counter
from datetime import date
from pathlib import Path
from urllib.parse import urlencode, urlparse

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent))
import hatvp  # noqa: E402

def env_url(name, default):
    return (os.environ.get(name, "").strip().strip("\"'") or default).rstrip("/")

SITE_URL = env_url("SITE_URL", "http://localhost:8000")
APP_URL  = env_url("APP_URL", "https://hatvp-to-table.onrender.com")
BASE_PATH = urlparse(SITE_URL).path.rstrip("/")
HATVP_FICHE_URL = "https://www.hatvp.fr/fiche-organisation/?organisation={}"

REP = hatvp.COL_REP_ID
LATEST_ACTIONS = 10
TOP_N = 10

# Colonnes supplémentaires par rapport à l'app : dépenses, dates, désinscription…
SITE_TABLE_COLS = {
    **hatvp.TABLE_COLS,
    "infos": hatvp._col_filter({
        "representants_id", "denomination", "nom_usage_hatvp", "sigle_hatvp",
        "label_categorie_organisation", "adresse", "code_postal", "ville", "pays",
        "site_web", "page_linkedin", "page_twitter", "page_facebook",
        "date_premiere_publication", "identifiant_national", "type_identifiant_national",
        "datecessation"}),
    "exercices": hatvp._col_filter({
        "exercices_id", "representants_id", "annee_debut", "annee_fin",
        "date_debut", "date_fin", "montant_depense", "montant_depense_inf",
        "nombre_salaries"}),
    "objets": hatvp._col_filter({
        "objet_activite", "exercices_id", "activite_id", "date_publication_activite"}),
}

# ─── UTILITAIRES ──────────────────────────────────────────────────────────────

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s or None

def clean_id(v):
    s = clean(v)
    return s[:-2] if s and s.endswith(".0") else s

def clean_url(v):
    s = clean(v)
    if s and not re.match(r"https?://", s, re.I):
        s = "https://" + s
    return s

def slugify(text, max_len=60):
    s = re.sub(r"[^a-z0-9]+", "-", hatvp.normalize(text or "")).strip("-")
    return s[:max_len].rstrip("-") or "sans-nom"

def fr_int(n):
    return f"{int(n):,}".replace(",", "\u00a0")

def fmt_date(v):
    """'24/01/2023 10:45:59' ou '2024-04-02' → '24/01/2023'."""
    s = clean(v)
    if not s:
        return None
    ts = pd.to_datetime(s, dayfirst="/" in s, errors="coerce")
    return None if pd.isna(ts) else f"{ts:%d/%m/%Y}"

def fmt_period(start, end):
    a = pd.to_datetime(clean(start), errors="coerce")
    b = pd.to_datetime(clean(end), errors="coerce")
    if pd.isna(a) or pd.isna(b):
        return None
    if a.year == b.year and (a.month, a.day, b.month, b.day) == (1, 1, 12, 31):
        return str(a.year)
    return f"{a:%m/%Y} – {b:%m/%Y}"

def fmt_etp(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(f) else (str(int(f)) if f.is_integer() else f"{f:g}".replace(".", ","))

def split_counts(series, top=None):
    c = Counter()
    for v in series.dropna():
        for part in str(v).split(" | "):
            if part and part != "nan":
                c[part] += 1
    return c.most_common(top)

def app_link(mode, query, exact=True):
    params = {"mode": mode, "q": query}
    if exact:
        params["exact"] = "1"
    return f"{APP_URL}/?{urlencode(params)}"

def letter_of(name):
    first = slugify(name)[:1]
    return first if first.isalpha() else "0-9"

# ─── PRÉPARATION DES DONNÉES ──────────────────────────────────────────────────

def load(zip_path):
    t = time.time()
    tables = hatvp.read_tables(zip_path or hatvp.ensure_zip(), SITE_TABLE_COLS)
    print(f"Données chargées en {time.time() - t:.1f}s")
    return tables

def prepare(tables, limit=None):
    infos = tables["infos"]
    exo = tables["exercices"].copy()
    exo["periode"] = [fmt_period(a, b) for a, b in zip(exo["date_debut"], exo["date_fin"])]

    enriched = hatvp.enrich_actions(
        tables["objets"], exo, tables["ari"], tables["actions"],
        tables["ministeres"], tables["domaines"], tables.get("beneficiaires"))
    enriched["periode"] = enriched["exercices_id"].map(exo.set_index("exercices_id")["periode"])
    objets_par_exo = tables["objets"].groupby("exercices_id").size()

    clients = tables.get("clients", pd.DataFrame())
    clients = clients.assign(id_client=clients["identifiant_national_client"].map(clean_id))

    groups = {
        "actions":  dict(tuple(enriched.groupby(REP))),
        "exo":      dict(tuple(exo.groupby(REP))),
        "dirig":    dict(tuple(tables["dirigeants"].groupby(REP))),
        "collab":   dict(tuple(tables["collaborateurs"].groupby(REP))),
        "clients":  dict(tuple(clients.groupby(REP))),
        "mandants": dict(tuple(clients.dropna(subset=["id_client"]).groupby("id_client"))),
    }
    secteurs_par_rep = (tables["secteurs"].dropna(subset=[hatvp.COL_SECTEUR])
                        .groupby(REP)[hatvp.COL_SECTEUR].apply(lambda s: sorted(set(s))))

    # Identité et URL de chaque organisation (nécessaire avant les liens croisés)
    orgs = {}
    used_slugs = set()
    for row in infos.itertuples(index=False):
        rid = getattr(row, REP)
        denom = clean(row.denomination) or "Sans nom"
        ident = clean_id(row.identifiant_national)
        slug = f"{slugify(clean(row.nom_usage_hatvp) or denom)}-{slugify(ident or f'id{rid}')}"
        if slug in used_slugs:
            slug = f"{slug}-{rid}"
        used_slugs.add(slug)
        actions = groups["actions"].get(rid)
        name = clean(row.nom_usage_hatvp) or denom
        sigle = clean(row.sigle_hatvp)
        if sigle and hatvp.normalize(sigle) in hatvp.normalize(name):
            sigle = None
        orgs[rid] = {
            "rid": rid, "slug": slug, "url": f"/organisation/{slug}/",
            "name": name, "denomination": denom, "sigle": sigle, "ident": ident,
            "category": clean(row.label_categorie_organisation),
            "n_actions": 0 if actions is None else actions["activite_id"].nunique(),
        }
    by_ident = {o["ident"]: o for o in orgs.values() if o["ident"]}

    selected = sorted(orgs.values(), key=lambda o: -o["n_actions"])
    if limit:
        selected = selected[:limit]
    selected_ids = {o["rid"] for o in selected}

    infos_by_rid = infos.set_index(REP)
    for o in selected:
        rid = o["rid"]
        info = infos_by_rid.loc[rid]
        actions = groups["actions"].get(rid, enriched.iloc[0:0])
        exos = groups["exo"].get(rid, exo.iloc[0:0])

        exercices = []
        for e in exos.sort_values("date_debut", ascending=False).itertuples(index=False):
            n = int(objets_par_exo.get(e.exercices_id, 0))
            depense = clean(e.montant_depense)
            if n == 0 and not depense:
                continue
            exercices.append({"periode": e.periode or "—", "n": n, "depense": depense,
                              "etp": fmt_etp(e.nombre_salaries)})
        max_n = max([e["n"] for e in exercices] or [1]) or 1
        for e in exercices:
            e["pct"] = round(100 * e["n"] / max_n)

        latest = (actions.sort_values(["date_publication_activite", "activite_id"],
                                      ascending=False, na_position="last")
                  .drop_duplicates("activite_id").head(LATEST_ACTIONS))

        people = []
        for df, statut, suffix in [(groups["dirig"].get(rid), "Dirigeant", "dirigeant"),
                                   (groups["collab"].get(rid), "Collaborateur", "collaborateur")]:
            if df is None:
                continue
            for p in df.itertuples(index=False):
                people.append({
                    "statut": statut,
                    "nom": " ".join(filter(None, [clean(getattr(p, f"prenom_{suffix}", None)),
                                                  clean(getattr(p, f"nom_{suffix}", None))])),
                    "fonction": clean(getattr(p, f"fonction_{suffix}", None)),
                })

        def client_entry(name, ident):
            target = by_ident.get(ident)
            return {"name": name, "url": target["url"] if target else None}

        cl = groups["clients"].get(rid)
        clients_list = [] if cl is None else sorted(
            {(clean(c.denomination_client), c.id_client) for c in cl.itertuples(index=False)
             if clean(c.denomination_client)}, key=lambda x: x[0])
        mandants = groups["mandants"].get(o["ident"]) if o["ident"] else None
        cabinets = [] if mandants is None else sorted(
            set(mandants[REP]) & orgs.keys(), key=lambda r: orgs[r]["name"])

        o.update({
            "adresse": " ".join(filter(None, [clean(info.get("adresse")), clean(info.get("code_postal"))])),
            "ville": clean(info.get("ville")), "pays": clean(info.get("pays")),
            "site_web": clean_url(info.get("site_web")), "linkedin": clean_url(info.get("page_linkedin")),
            "twitter": clean_url(info.get("page_twitter")), "facebook": clean_url(info.get("page_facebook")),
            "type_ident": clean(info.get("type_identifiant_national")),
            "hatvp_url": HATVP_FICHE_URL.format(o["ident"]) if o["ident"] else None,
            "first_pub": fmt_date(info.get("date_premiere_publication")),
            "cessation": fmt_date(info.get("datecessation")),
            "secteurs": [{"name": s, "url": f"/secteur/{slugify(s)}/"}
                         for s in secteurs_par_rep.get(rid, [])],
            "exercices": exercices,
            "last_depense": next((e for e in exercices if e["depense"]), None),
            "types": split_counts(actions["types_actions"], TOP_N),
            "responsables": split_counts(actions["responsables_publics"], TOP_N),
            "domaines": split_counts(actions.drop_duplicates("activite_id")["domaines_intervention"], TOP_N),
            "latest": [{"objet": clean(a.objet_activite), "periode": a.periode,
                        "types": clean(a.types_actions), "responsables": clean(a.responsables_publics),
                        "donneur": clean(a.donneur_ordre)}
                       for a in latest.itertuples(index=False)],
            "people": people,
            "clients": [client_entry(n, i) for n, i in clients_list],
            "cabinets": [{"name": orgs[r]["name"], "url": orgs[r]["url"]} for r in cabinets],
            "app_url": app_link("organisations", o["denomination"]),
            "excel": f"hatvp-{o['slug']}.xlsx",
        })
        sources = [o["linkedin"], o["twitter"], o["facebook"], o["hatvp_url"]]
        o["jsonld"] = {k: v for k, v in {
            "@context": "https://schema.org", "@type": "Organization",
            "name": o["name"], "legalName": o["denomination"], "alternateName": o["sigle"],
            "identifier": o["ident"], "url": o["site_web"],
            "sameAs": [s for s in sources if s],
        }.items() if v}

    return {"orgs": orgs, "selected": selected, "selected_ids": selected_ids,
            "enriched": enriched, "secteurs": tables["secteurs"], "infos": infos,
            "tables": tables}

def prepare_secteurs(data):
    sect = data["secteurs"].dropna(subset=[hatvp.COL_SECTEUR])
    enriched = data["enriched"]
    result = []
    for name, g in sect.groupby(hatvp.COL_SECTEUR):
        rids = set(g[REP])
        members = sorted((data["orgs"][r] for r in rids if r in data["orgs"]),
                         key=lambda o: (-o["n_actions"], o["name"]))
        acts = enriched[enriched[REP].isin(rids)]
        result.append({
            "name": name, "slug": slugify(name), "url": f"/secteur/{slugify(name)}/",
            "orgs": members, "n_orgs": len(members),
            "n_actions": acts["activite_id"].nunique(),
            "domaines": split_counts(acts.drop_duplicates("activite_id")["domaines_intervention"], TOP_N),
            "responsables": split_counts(acts["responsables_publics"], TOP_N),
            "app_url": app_link("secteurs", name),
        })
    return sorted(result, key=lambda s: -s["n_orgs"])

# ─── RENDU ────────────────────────────────────────────────────────────────────

class Site:
    def __init__(self, out):
        self.out = out
        self.paths = []
        self.env = Environment(loader=FileSystemLoader(ROOT / "templates"),
                               autoescape=select_autoescape(["html", "xml"]),
                               trim_blocks=True, lstrip_blocks=True)
        self.env.filters["fr"] = fr_int
        self.env.filters["u"] = lambda path: BASE_PATH + path
        self.env.globals.update(site_url=SITE_URL, app_url=APP_URL,
                                build_date=f"{date.today():%d/%m/%Y}")

    def page(self, path, template, **ctx):
        target = self.out / path.lstrip("/") / "index.html" if path.endswith("/") else self.out / path.lstrip("/")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(self.env.get_template(template).render(path=path, **ctx), encoding="utf-8")
        if path.endswith("/"):
            self.paths.append(path)

def build_excel(data, org):
    t = data["tables"]
    rid = org["rid"]
    enriched = data["enriched"]
    df = enriched[enriched[REP] == rid]
    s1 = hatvp.build_actions_sheet(df, data["infos"])
    s2 = hatvp.build_orgs_sheet([rid], data["infos"], df)
    s3 = hatvp.build_persons_sheet([rid], s2, t["dirigeants"], t["collaborateurs"])
    s4 = hatvp.build_clients_sheet([rid], t.get("clients", pd.DataFrame()), data["infos"])
    return hatvp.build_excel([
        ("Actions de lobbying", s1),
        ("Organisations", s2.drop(columns=[REP], errors="ignore")),
        ("Dirigeants & Collaborateurs", s3),
        ("Clients & Mandats", s4),
    ])

def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--zip", type=Path, help="ZIP HATVP local (sinon téléchargé)")
    parser.add_argument("--out", type=Path, default=ROOT / "dist")
    parser.add_argument("--limit", type=int, help="Nombre max d'organisations (prototype)")
    parser.add_argument("--no-excel", action="store_true")
    args = parser.parse_args()

    t0 = time.time()
    data = prepare(load(args.zip), args.limit)
    secteurs = prepare_secteurs(data)
    print(f"Données préparées en {time.time() - t0:.1f}s")

    if args.out.exists():
        shutil.rmtree(args.out)
    shutil.copytree(ROOT / "static", args.out / "static")
    site = Site(args.out)
    selected = data["selected"]
    orgs = [o for o in data["orgs"].values() if o["rid"] in data["selected_ids"]]

    # Organisations
    t = time.time()
    for o in selected:
        o["excel_available"] = not args.no_excel
        site.page(o["url"], "organisation.html", org=o)
        if not args.no_excel:
            (args.out / o["url"].lstrip("/") / o["excel"]).write_bytes(build_excel(data, o))
    print(f"{len(selected)} pages organisation en {time.time() - t:.1f}s")

    # Secteurs : seules les organisations générées sont listées
    for s in secteurs:
        s["orgs"] = [o for o in s["orgs"] if o["rid"] in data["selected_ids"]]
        site.page(s["url"], "secteur.html", secteur=s)
    site.page("/secteurs/", "secteurs.html", secteurs=secteurs)

    # Index alphabétique
    letters = {}
    for o in sorted(orgs, key=lambda o: slugify(o["name"])):
        letters.setdefault(letter_of(o["name"]), []).append(o)
    letter_keys = sorted(letters, key=lambda k: (k == "0-9", k))
    for k in letter_keys:
        site.page(f"/organisations/{k}/", "organisations_lettre.html",
                  letter=k, orgs=letters[k], letters=letter_keys)
    site.page("/organisations/", "organisations.html", letters=letter_keys,
              counts={k: len(v) for k, v in letters.items()},
              top=sorted(orgs, key=lambda o: -o["n_actions"])[:50])

    # Accueil
    site.page("/", "index.html",
              n_orgs=len(data["orgs"]), n_actions=data["enriched"]["activite_id"].nunique(),
              n_secteurs=len(secteurs), top=selected[:12], secteurs=secteurs)
    site.page("/404.html", "404.html")

    # Index de recherche client, sitemap, robots
    index = [{"n": o["name"], "d": o["denomination"] if o["denomination"] != o["name"] else "",
              "s": o["sigle"] or "", "u": BASE_PATH + o["url"], "a": o["n_actions"]} for o in orgs]
    (args.out / "search-index.json").write_text(json.dumps(index, ensure_ascii=False, separators=(",", ":")),
                                                encoding="utf-8")
    today = date.today().isoformat()
    urls = "\n".join(f"  <url><loc>{SITE_URL}{p}</loc><lastmod>{today}</lastmod></url>" for p in site.paths)
    (args.out / "sitemap.xml").write_text(
        f'<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n{urls}\n</urlset>\n', encoding="utf-8")
    (args.out / "robots.txt").write_text(f"User-agent: *\nAllow: /\n\nSitemap: {SITE_URL}/sitemap.xml\n")

    print(f"Site généré dans {args.out} : {len(site.paths)} pages en {time.time() - t0:.1f}s")

if __name__ == "__main__":
    main()
