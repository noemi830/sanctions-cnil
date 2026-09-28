"""
Veille CNIL : lit https://www.cnil.fr/fr/actualite, repère les articles
portant le tag #Sanction, et ajoute une ligne par nouvelle sanction
dans Sanctions_CNIL.xlsx (colonne G = lien de l'article, sert à éviter les doublons).

Usage :  python veille_cnil.py            (met à jour le fichier)
         python veille_cnil.py --dry-run  (affiche sans rien écrire)
"""
import re
import sys
import datetime as dt
from pathlib import Path
from urllib.parse import urljoin

import html as htmlmod
import shutil
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from copy import copy

# ---------------- Réglages ----------------
FICHIER_EXCEL = Path(__file__).with_name("Sanctions_CNIL.xlsx")
DOSSIER_SITE = Path(__file__).with_name("site")  # page web publiée par GitHub Pages
URL_ACTUS = "https://www.cnil.fr/fr/actualite"
NB_PAGES = 3                      # pages de la liste à parcourir (~6 articles/page)
TITRES_EXCLUS = ["clôture de l’injonction", "clôture de l'injonction"]  # tag #Sanction mais pas une sanction
SURLIGNER_NOUVEAU = PatternFill("solid", fgColor="FFF2CC")  # jaune pâle = ligne à relire
COL_LIEN, COL_AJOUT = 7, 8        # G = lien article, H = date d'ajout automatique
HEADERS = {"User-Agent": "Mozilla/5.0 (veille sanctions CNIL, usage personnel)"}
# ------------------------------------------

MOIS = {m: i for i, m in enumerate(
    ["janvier", "février", "mars", "avril", "mai", "juin", "juillet",
     "août", "septembre", "octobre", "novembre", "décembre"], 1)}
RE_DATE = re.compile(r"(\d{1,2})(?:er)?\s+(" + "|".join(MOIS) + r")\s+(\d{4})", re.I)


def parse_date(txt):
    m = RE_DATE.search(txt or "")
    if not m:
        return None
    return dt.datetime(int(m.group(3)), MOIS[m.group(2).lower()], int(m.group(1)))


def telecharger(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text


def extraire_articles(html, base=URL_ACTUS):
    """Renvoie la liste des articles de la page : titre, lien, résumé, tags, date de publication."""
    soup = BeautifulSoup(html, "html.parser")
    articles = {}
    for h3 in soup.find_all("h3"):
        a = h3.find("a", href=True)
        if not a:
            continue
        lien = urljoin(base, a["href"])
        # remonte jusqu'au bloc qui contient CET article uniquement (et sa date)
        bloc = h3
        for _ in range(8):
            parent = bloc.parent
            if parent is None:
                break
            liens_h3 = {urljoin(base, x["href"]) for h in parent.find_all("h3") for x in h.find_all("a", href=True)}
            if len(liens_h3) > 1:
                break
            bloc = parent
        texte = bloc.get_text(" ", strip=True)
        tags = [t.get_text(strip=True) for t in bloc.find_all("a", href=re.compile(r"/tag/"))]
        if not tags:
            tags = re.findall(r"#[\wÀ-ÿ’' -]+?(?=\s#|\s\d|$)", texte)
        titre = a.get_text(" ", strip=True)
        resume = texte.replace(titre, "", 1)
        resume = re.split(r"\s#", resume)[0].strip()
        dates = RE_DATE.findall(texte)
        date_pub = None
        if dates:
            j, m, y = dates[-1]
            date_pub = dt.datetime(int(y), MOIS[m.lower()], int(j))
        articles[lien] = dict(titre=titre, lien=lien, resume=resume,
                              tags=" ".join(tags), date_pub=date_pub)
    return list(articles.values())


def est_sanction(art):
    if "sanction" not in art["tags"].lower():
        return False
    return not any(x in art["titre"].lower() for x in TITRES_EXCLUS)


TAGS_IGNORES = {"sanction", "particulier", "professionnel"}


def _propre(t):
    return re.sub(r"\s+", " ", (t or "").replace("\u2019", "'")).strip()


def details_article(lien):
    """Lit la page de l'article CNIL et en tire, quand c'est possible :
    - la délibération (référence SAN + date), dans la rubrique « La délibération » en bas d'article ;
    - les manquements, à partir des intertitres « Un manquement à ... (article X du RGPD) » ;
    - l'entité (« concernant la société X ») et le montant (« amende de X euros ») ;
    - les mots clés thématiques que la CNIL attache à l'article (#Cookies, #Prospection...).
    """
    infos = dict(ref="", date=None, manquements="", entite="", montant="", mots_cles="", texte="")
    try:
        soup = BeautifulSoup(telecharger(lien), "html.parser")
    except Exception as e:
        print(f"  ! impossible de lire {lien} : {e}")
        return infos
    corps = soup.find("main") or soup
    texte = _propre(corps.get_text(" ", strip=True))

    # 1) Délibération(s) : "Délibération SAN-2025-017 du 30 décembre 2025"
    #    ou "Délibération de la formation restreinte n° SAN – 2026-010 du 21 juillet 2026 concernant la société EXTIA"
    refs, dates = [], []
    for m in re.finditer(r"Délibération[^.]{0,80}?(SAN\s*[-–]\s*\d{4}\s*[-–]\s*\d{3})\s*(?:du\s+(\d{1,2}(?:er)?\s+\w+\s+\d{4}))?"
                         r"(?:[^-]{0,60}?(?:concernant|à l'encontre de)\s+(?:la société |les sociétés |l'|le |la )?([^-|]+?)(?=\s+-\s+Légifrance|\s+Texte\b|\s+Délibération|$))?",
                         texte):
        ref = re.sub(r"\s", "", m.group(1)).replace("–", "-")
        if ref not in refs:
            refs.append(ref)
            if m.group(2):
                dates.append(parse_date(m.group(2)))
            if m.group(3) and not infos["entite"] and len(m.group(3)) < 80 and not m.group(3).strip().upper().startswith("X"):
                infos["entite"] = m.group(3).strip()
    infos["ref"] = " et ".join(f"Délibération {r}" for r in refs)
    infos["date"] = next((d for d in dates if d), None)

    # 2) Manquements : intertitres "Un manquement à l'obligation de ... (article 6 du RGPD)"
    #    (variantes rencontrées : h2 ou h3, "article 32 RGPD", "article L.34-5 du CPCE", "articles 5-1-a) et 5-2 du RGPD")
    lignes, vus = [], set()
    for h in corps.find_all(["h2", "h3", "h4"]):
        t = _propre(h.get_text(" "))
        m = re.search(r"^(?:un |des |le |les )?(manquements?\b.*?)\s*\((?:articles?|art\.)\s+(.+?)\s+((?:du |de la )?(?:RGPD|règlement[^)]*|loi[^)]*|CPCE|code des postes[^)]*|LIL))\s*\)\s*$", t, re.I)
        if not m:
            continue
        desc, nums, src = m.group(1), m.group(2), m.group(3).lower()
        if "loi" in src or "lil" in src:
            nums, suffixe = nums, " LIL"
        elif "cpce" in src or "postes" in src:
            suffixe = " CPCE"
        else:
            suffixe = ""
        ligne = f"Art {nums}{suffixe} ({desc[0].lower() + desc[1:]})"
        if ligne not in vus:
            vus.add(ligne)
            lignes.append(ligne)
    #    Autre présentation : liste à puces en gras sous « Les manquements sanctionnés »,
    #    avec l'article cité dans la phrase d'introduction (cas fréquent pour les cookies).
    if not lignes:
        titre = next((h for h in corps.find_all(["h2", "h3"]) if "manquements sanctionn" in _propre(h.get_text()).lower()), None)
        if titre:
            bloc, puces = [], []
            for el in titre.find_all_next():
                if el.name == "h2" and el is not titre:
                    break
                if el.name in ("p", "h3"):
                    bloc.append(_propre(el.get_text(" ")))
                if el.name == "strong":
                    puces.append(_propre(el.get_text(" ")).rstrip(" :"))
            intro = " ".join(bloc)
            m = re.search(r"articles?\s+([\dL][\w.\-]*(?:\s+et\s+[\w.\-]+)?)\s+(du RGPD|de la loi Informatique et Libertés|du CPCE)", intro, re.I)
            if m:
                num, src = m.group(1), m.group(2)
            else:
                m = re.search(r"(RGPD|loi Informatique et Libertés|CPCE)\s*\((?:articles?|art\.)\s+([\w.\-]+)\)", intro, re.I)
                num, src = (m.group(2), m.group(1)) if m else (None, None)
            if num and puces:
                suffixe = " LIL" if "loi" in src.lower() else (" CPCE" if "CPCE" in src else "")
                lignes.append(f"Art {num}{suffixe} (" + " ; ".join(p[0].lower() + p[1:] for p in puces) + ")")
    infos["manquements"] = "\n".join(lignes)

    # 3) Montant : "amende de 3,5 millions d'euros" / "amende administrative d'un montant de 5 000 euros"
    m = re.search(r"amendes?(?: administratives?)?(?: d'un montant)? de ((?:respectivement )?\d[\d\s.,]*(?:\s*(?:et|,)\s*\d[\d\s.,]*)?\s*(?:millions?|milliards?)?)\s*(?:d'euros|euros|€)", texte, re.I)
    if m:
        infos["montant"] = m.group(1).replace("respectivement ", "").strip() + " €"

    # 4) Mots clés : tags thématiques de la CNIL + phrase d'accroche de l'article
    tags = []
    for a in corps.find_all("a", href=re.compile(r"/tag/")):
        t = _propre(a.get_text()).lstrip("#")
        if t and t.lower() not in TAGS_IGNORES and t not in tags:
            tags.append(t)
    infos["mots_cles"] = "\n".join(tags)
    infos["texte"] = texte
    if re.search(r"sans qu'il soit[^.]*utile de nommer|ne peut pas divulguer le nom|anonymis", texte, re.I) and not infos["entite"]:
        infos["entite"] = "Identité de la société non disponible"
    return infos


def entite(titre):
    m = re.search(r"à l[’']encontre (?:de la société |des sociétés |de l[’']|du |de la |des |de |d[’'])(.+)$", titre, re.I)
    return m.group(1).strip() if m else ""


def montant(titre, resume):
    for t in (titre, resume):
        m = re.search(r"(\d[\d\s.,]*\s*(?:millions?|milliards?)?)\s*(?:d[’']euros|euros|€)", t, re.I)
        if m:
            return m.group(1).strip() + " €"
    return ""


def generer_page():
    """Construit site/index.html (le tableau que voient les boss) à partir de l'Excel."""
    ws = load_workbook(FICHIER_EXCEL, data_only=True).worksheets[0]
    entetes = [str(ws.cell(1, c).value or "").strip() for c in range(1, 7)]
    lignes = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, COL_LIEN + 1)]
        if all(v in (None, "") for v in vals[:6]):
            continue
        lignes.append(vals)
    lignes.sort(key=lambda v: v[1] if isinstance(v[1], dt.datetime) else dt.datetime.min, reverse=True)

    def cellule(v):
        if isinstance(v, dt.datetime):
            return f'<td data-sort="{v:%Y%m%d}">{v:%d/%m/%Y}</td>'
        t = htmlmod.escape(str(v or "").strip()).replace("\n", "<br>")
        cls = ' class="todo"' if "À compléter" in t else ""
        return f"<td{cls}>{t}</td>"

    corps = ""
    for v in lignes:
        lien = str(v[6] or "").strip()
        a = f'<a href="{htmlmod.escape(lien)}" target="_blank" rel="noopener">Voir</a>' if lien.startswith("http") else ""
        corps += "<tr>" + "".join(cellule(x) for x in v[:6]) + f"<td>{a}</td></tr>\n"
    maj = dt.datetime.now(ZoneInfo("Europe/Paris"))
    th = "".join(f"<th>{htmlmod.escape(e)}</th>" for e in entetes) + "<th>Article CNIL</th>"
    page = PAGE_HTML.format(th=th, corps=corps, n=len(lignes), maj=f"{maj:%d/%m/%Y à %H:%M}")
    DOSSIER_SITE.mkdir(exist_ok=True)
    (DOSSIER_SITE / "index.html").write_text(page, encoding="utf-8")
    shutil.copy(FICHIER_EXCEL, DOSSIER_SITE / FICHIER_EXCEL.name)
    print(f"Page générée : {DOSSIER_SITE / 'index.html'} ({len(lignes)} sanctions)")


PAGE_HTML = """<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Veille des sanctions CNIL</title>
<style>
 body{{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f5f6f8;color:#1d1d1f}}
 header{{background:#1f3a5f;color:#fff;padding:20px 28px}}
 header h1{{margin:0 0 6px;font-size:22px}} header p{{margin:0;opacity:.85;font-size:14px}}
 main{{padding:20px 28px}}
 .barre{{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin-bottom:14px}}
 input{{padding:8px 10px;border:1px solid #c7ccd4;border-radius:6px;min-width:260px;font-size:14px}}
 .btn{{background:#1f3a5f;color:#fff;text-decoration:none;padding:8px 14px;border-radius:6px;font-size:14px}}
 .tab{{overflow-x:auto;background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.08)}}
 table{{border-collapse:collapse;width:100%;font-size:13px}}
 th{{background:#e9edf3;text-align:left;padding:10px;position:sticky;top:0;cursor:pointer;white-space:nowrap}}
 td{{padding:10px;border-top:1px solid #eceef2;vertical-align:top;min-width:90px}}
 td:nth-child(3),td:nth-child(5){{min-width:260px}}
 tr:hover td{{background:#fafbfd}} .todo{{color:#b26a00;font-style:italic}}
 a{{color:#0563c1}} footer{{padding:10px 28px 30px;font-size:12px;color:#6b7280}}
</style></head><body>
<header><h1>Veille des sanctions CNIL</h1>
<p>{n} sanctions · mis à jour automatiquement chaque jour ouvré · dernière vérification le {maj}</p></header>
<main>
<div class="barre"><input id="q" placeholder="Rechercher (entité, article, mot clé…)" oninput="filtrer()">
<a class="btn" href="Sanctions_CNIL.xlsx" download>Télécharger l'Excel</a></div>
<div class="tab"><table id="t"><thead><tr>{th}</tr></thead><tbody>
{corps}</tbody></table></div></main>
<footer>Source : cnil.fr/fr/actualite. Les informations marquées « À compléter » ou « à vérifier » ont été extraites automatiquement.</footer>
<script>
function filtrer(){{var q=document.getElementById('q').value.toLowerCase();
 document.querySelectorAll('#t tbody tr').forEach(function(r){{r.style.display=r.innerText.toLowerCase().indexOf(q)>-1?'':'none'}});}}
document.querySelectorAll('#t th').forEach(function(th,i){{var asc=false;th.onclick=function(){{asc=!asc;
 var b=document.querySelector('#t tbody');var rs=Array.from(b.rows);
 rs.sort(function(x,y){{var a=x.cells[i].dataset.sort||x.cells[i].innerText,c=y.cells[i].dataset.sort||y.cells[i].innerText;
 return (asc?1:-1)*a.localeCompare(c,'fr',{{numeric:true}})}});rs.forEach(function(r){{b.appendChild(r)}});}}}});
</script></body></html>"""


def main(dry_run=False):
    wb = load_workbook(FICHIER_EXCEL)
    ws = wb.worksheets[0]
    if ws.cell(1, COL_LIEN).value in (None, ""):
        ws.cell(1, COL_LIEN, "Lien article CNIL")
        ws.cell(1, COL_AJOUT, "Ajouté automatiquement le")
        for c in (COL_LIEN, COL_AJOUT):
            ws.cell(1, c)._style = copy(ws.cell(1, 1)._style)
    deja = {str(ws.cell(r, COL_LIEN).value).strip() for r in range(2, ws.max_row + 1) if ws.cell(r, COL_LIEN).value}

    trouves, nb_articles = [], 0
    for p in range(NB_PAGES):
        html = telecharger(f"{URL_ACTUS}?page={p}")
        arts = extraire_articles(html)
        nb_articles += len(arts)
        print(f"Page {p + 1} : {len(arts)} articles lus")
        trouves += [a for a in arts if est_sanction(a)]

    # Garde-fou : si aucun article n'est lu, la page de la CNIL a sans doute changé.
    # On fait échouer le robot (croix rouge + e-mail automatique de GitHub) au lieu de se taire.
    if nb_articles == 0:
        sys.exit("ERREUR : aucun article lu sur cnil.fr/fr/actualite. La structure du site a peut-être changé.")

    trouves = list(dict((a["lien"], a) for a in trouves).values())
    print(f"\nSanctions repérées sur cnil.fr ({len(trouves)}) :")
    for a in trouves:
        etat = "déjà dans le tableau" if a["lien"] in deja else "NOUVELLE"
        print(f"  - [{etat}] {a['titre']}")
    print()

    nouveaux = [a for a in trouves if a["lien"] not in deja]
    nouveaux.sort(key=lambda a: a["date_pub"] or dt.datetime.min)
    if not nouveaux:
        print("Aucune nouvelle sanction.")
        if not dry_run:
            generer_page()
        return []

    derniere = max(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value not in (None, ""))
    modele = derniere
    ajoutes = []
    for art in nouveaux:
        # On ne garde que les sanctions dont la délibération est publiée (référence SAN-xxxx dans l'article).
        # Les récapitulatifs de sanctions en procédure simplifiée (anonymes, sans délibération) sont ignorés.
        if re.search(r"simplifi", art["titre"] + " " + art["resume"], re.I):
            print(f"  (ignorée, procédure simplifiée) {art['titre']}")
            continue
        d = details_article(art["lien"])
        if not d["ref"]:
            print(f"  (ignorée, pas de délibération publiée) {art['titre']}")
            continue
        ajoutes.append(art)
        date_decision = d["date"] or parse_date(art["resume"]) or art["date_pub"]
        chapo = art["resume"].rstrip("…. ")
        m = re.search(re.escape(_propre(chapo)[:60]) + r"[^.]*\.", d["texte"]) if chapo else None
        if m:
            chapo = m.group(0)
        mots = chapo + ("\n" + d["mots_cles"] if d["mots_cles"] else "")
        ligne = [
            d["ref"] or "À compléter",
            date_decision,
            mots,
            entite(art["titre"]) or d["entite"] or "À compléter",
            d["manquements"] or "À compléter",
            montant(art["titre"], art["resume"]) or d["montant"] or "À compléter",
            art["lien"],
            dt.datetime.now().replace(microsecond=0),
        ]
        print(f"+ {art['titre']}")
        if dry_run:
            continue
        derniere += 1
        for col, val in enumerate(ligne, 1):
            c = ws.cell(derniere, col, val)
            c._style = copy(ws.cell(modele, min(col, 6))._style)
            c.fill = SURLIGNER_NOUVEAU
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(derniere, COL_LIEN).hyperlink = art["lien"]
        ws.cell(derniere, COL_LIEN).font = Font(name=ws.cell(modele, 1).font.name, color="0563C1", underline="single")
        ws.cell(derniere, COL_AJOUT).number_format = "dd/mm/yyyy hh:mm"

    if not dry_run:
        if ajoutes:
            wb.save(FICHIER_EXCEL)
        print(f"{len(ajoutes)} ligne(s) ajoutée(s) dans {FICHIER_EXCEL.name}")
        generer_page()
    return ajoutes
    return nouveaux


if __name__ == "__main__":
    if "--page-seule" in sys.argv:
        generer_page()
    else:
        main(dry_run="--dry-run" in sys.argv)
