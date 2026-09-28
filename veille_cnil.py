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


def details_article(lien):
    """Va lire la page de l'article pour trouver la référence SAN-xxxx et les articles cités."""
    try:
        soup = BeautifulSoup(telecharger(lien), "html.parser")
    except Exception as e:
        print(f"  ! impossible de lire {lien} : {e}")
        return "", ""
    corps = soup.find("main") or soup
    texte = corps.get_text(" ", strip=True)
    refs = sorted(set(re.findall(r"SAN\s?[-–]\s?\d{4}\s?[-–]\s?\d{3}", texte)))
    refs = [re.sub(r"\s", "", r).replace("–", "-") for r in refs]
    arts = re.findall(r"articles?\s+((?:\d+(?:\.\d+)*(?:\.[a-z])?(?:,\s*|\s+et\s+)?)+)\s*(du RGPD|de la loi Informatique et Libertés|RGPD|LIL)?", texte, re.I)
    vus, cites = set(), []
    for nums, source in arts:
        src = "LIL" if source and "loi" in source.lower() or source == "LIL" else "RGPD"
        for n in re.findall(r"\d+(?:\.\d+)*(?:\.[a-z])?", nums):
            k = f"Art {n} {src}"
            if k not in vus:
                vus.add(k)
                cites.append(k)
    return " et ".join(f"Délibération {r}" for r in refs), "\n".join(cites)


def entite(titre):
    m = re.search(r"à l[’']encontre (?:de la société |de l[’']|du |de la |des |de |d[’'])(.+)$", titre, re.I)
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

    trouves = []
    for p in range(NB_PAGES):
        html = telecharger(f"{URL_ACTUS}?page={p}")
        trouves += [a for a in extraire_articles(html) if est_sanction(a)]

    nouveaux = [a for a in dict((a["lien"], a) for a in trouves).values() if a["lien"] not in deja]
    nouveaux.sort(key=lambda a: a["date_pub"] or dt.datetime.min)
    if not nouveaux:
        print("Aucune nouvelle sanction.")
        if not dry_run:
            generer_page()
        return []

    derniere = max(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value not in (None, ""))
    modele = derniere
    for art in nouveaux:
        ref, articles_cites = details_article(art["lien"])
        date_decision = parse_date(art["resume"]) or art["date_pub"]
        ligne = [
            ref or "À compléter",
            date_decision,
            f"{art['titre']}\n{art['resume']}",
            entite(art["titre"]) or "À compléter",
            (articles_cites + "\n(extraction auto, à vérifier)") if articles_cites else "À compléter",
            montant(art["titre"], art["resume"]) or "À compléter",
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
        wb.save(FICHIER_EXCEL)
        print(f"{len(nouveaux)} ligne(s) ajoutée(s) dans {FICHIER_EXCEL.name}")
        generer_page()
    return nouveaux


if __name__ == "__main__":
    if "--page-seule" in sys.argv:
        generer_page()
    else:
        main(dry_run="--dry-run" in sys.argv)
