import streamlit as st
import anthropic
import re
import json
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import _Cell, _Row
from copy import deepcopy
import io
import os
from datetime import date, timedelta

# ── Config page ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Assistant CR — MH2O",
    page_icon="🏗️",
    layout="wide"
)

# ── Style ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f4f7fb; }
    .stApp { font-family: 'DM Sans', sans-serif; }
    .header-box {
        background: #0d1f3c;
        color: white;
        padding: 24px 32px;
        border-radius: 12px;
        margin-bottom: 24px;
    }
    .header-box h1 { color: white; margin: 0; font-size: 28px; }
    .header-box p { color: rgba(255,255,255,0.7); margin: 8px 0 0; }
    .accent { color: #e8a020; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-box">
    <h1>🏗️ Assistant CR <span class="accent">MH2O</span></h1>
    <p>Reformulation de notes de chantier et insertion automatique dans le Word du CR précédent</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    # Clé lue dans .streamlit/secrets.toml ou la variable d'environnement, sinon saisie manuelle
    try:
        cle_defaut = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        cle_defaut = ""
    cle_defaut = cle_defaut or os.environ.get("ANTHROPIC_API_KEY", "")
    api_key = st.text_input("Clé API Claude", value=cle_defaut, type="password", placeholder="sk-ant-...")
    if cle_defaut:
        st.caption("🔑 Clé chargée automatiquement")
    st.markdown("---")
    st.markdown("""
### 📖 Comment utiliser
1. Entre ta clé API Claude (ou mets-la dans `.streamlit/secrets.toml`)
2. Upload le Word du CR précédent
3. Vérifie les infos du nouveau CR (pré-remplies)
4. Saisis toutes tes notes brutes, toutes sections confondues
5. Vérifie le tableau des points générés
6. Génère et télécharge le nouveau CR
    """)

# ── Constantes ────────────────────────────────────────────────────────────────
BLEU = RGBColor(0x00, 0x00, 0xFF)
BLEU_TITRE = RGBColor(0x98, 0xB5, 0xD8)
NOIR = RGBColor(0x00, 0x00, 0x00)
ROUGE = RGBColor(0xFF, 0x00, 0x00)

# ── Fonctions ─────────────────────────────────────────────────────────────────

def ecrire_dans_cellule(cell, texte, gras_majuscules=False, couleur=BLEU, ajouter=False):
    """Écrit en bleu (par défaut) en reprenant la police de la cellule ; entreprises (MAJUSCULES) en gras si demandé.
    ajouter=True complète le texte existant au lieu de le remplacer."""
    paras = cell.paragraphs
    if ajouter:
        para = next((p for p in reversed(paras) if p.text.strip()), paras[-1])
        runs_texte = [r for r in para.runs if r.text.strip()]
        run_modele = runs_texte[-1] if runs_texte else None
        texte = " " + texte
    else:
        for para in paras[1:]:
            para._p.getparent().remove(para._p)
        para = paras[0]
        # Mise en forme du premier run existant (police, taille…) à réutiliser
        run_modele = next((r for r in para.runs if r.text.strip()), None)
    rpr = run_modele._r.rPr if run_modele is not None else None
    if rpr is None and para._p.pPr is not None:
        # Cellule vide : Word garde la mise en forme sur la marque de paragraphe
        rpr = para._p.pPr.find(qn("w:rPr"))
    rpr = deepcopy(rpr) if rpr is not None else None
    if not ajouter:
        for enfant in list(para._p):
            if enfant.tag != qn("w:pPr"):
                para._p.remove(enfant)
    tokens = re.split(r'(\b[A-Z]{2}[A-Z0-9\-&/]*\b)', texte) if gras_majuscules else [texte]
    for token in tokens:
        if not token:
            continue
        run = para.add_run(token)
        if rpr is not None:
            run._r.insert(0, deepcopy(rpr))
        run.font.color.rgb = couleur
        if gras_majuscules:
            run.bold = bool(re.match(r'^[A-Z]{2}[A-Z0-9\-&/]*$', token))


def ligne_modele(table):
    """Dernière ligne remplie du tableau, ou None."""
    return next((r for r in reversed(table.rows[1:]) if r.cells[0].text.strip()), None)


def ajouter_ligne(table, modele):
    """Nouvelle ligne copiée sur le modèle : mêmes largeurs, bordures et alignements."""
    if modele is None:
        row = table.add_row()
    else:
        tr = deepcopy(modele._tr)
        table.rows[-1]._tr.addnext(tr)
        row = _Row(tr, table)
    # Cellules réelles (une cellule fusionnée n'apparaît qu'une fois)
    return [_Cell(tc, table) for tc in row._tr.tc_lst]


def colonnes_du_tableau(table, nb_cellules):
    """Repère les colonnes d'après l'en-tête (ACTION PAR, NOTE LE, POUR LE, FAIT LE)."""
    roles = {"texte": 0}
    entetes = [_Cell(tc, table).text.upper() for tc in table.rows[0]._tr.tc_lst]
    if len(entetes) == nb_cellules:
        for i, nom in enumerate(entetes[1:], start=1):
            if "ACTION" in nom or "RESP" in nom:
                roles.setdefault("responsable", i)
            elif "NOTE" in nom:
                roles.setdefault("note_le", i)
            elif "POUR" in nom or "CHÉANCE" in nom or "CHEANCE" in nom:
                roles.setdefault("pour_le", i)
            elif "FAIT" in nom:
                roles.setdefault("fait_le", i)
    if len(roles) == 1:
        # En-tête non reconnu : ordre habituel
        for i, role in enumerate(["responsable", "note_le", "pour_le", "fait_le"], start=1):
            if i < nb_cellules:
                roles[role] = i
    return roles


JOURS_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]
MOIS_FR = ["JANVIER", "FÉVRIER", "MARS", "AVRIL", "MAI", "JUIN",
           "JUILLET", "AOÛT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DÉCEMBRE"]

# Mots en majuscules qui ne sont pas des entreprises
MOTS_EXCLUS = {
    "PA", "PM", "EN", "ATT", "ATTENTE", "FAIT", "CR", "RDV", "BAT", "NB", "OK", "NON", "ACTION PAR",
    "DESIGNATION", "REUNION", "RÉUNION", "DATE", "STATUT", "OBS", "RESP", "LE", "LA",
    "LES", "DE", "DU", "DES", "ET", "AU", "AUX", "SUR", "POUR", "PAR",
} | set(JOURS_FR) | set(MOIS_FR)


def tous_les_paragraphes(doc):
    """Tous les paragraphes : corps, tableaux, zones de texte, en-têtes et pieds de page."""
    racines = [doc.element.body]
    for sec in doc.sections:
        for hf in (sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer):
            # Un en-tête lié au précédent n'a pas de contenu propre
            if not hf.is_linked_to_previous:
                racines.append(hf._element)
    paragraphes, vus = [], set()
    for racine in racines:
        if id(racine) in vus:
            continue
        vus.add(id(racine))
        paragraphes.extend(Paragraph(p, None) for p in racine.iter(qn("w:p")))
    return paragraphes


def remplacer_texte_paragraphe(para, nouveau):
    """Remplace le texte d'un paragraphe découpé en plusieurs runs, en gardant la mise en forme du 1er."""
    runs = [r for r in para.runs if r.text]
    if not runs:
        return
    runs[0].text = nouveau
    for run in runs[1:]:
        run.text = ""


def infos_cr_precedent(doc):
    """Extrait (numéro, heure du RDV) du CR précédent, None si introuvable."""
    numero, heure = None, None
    for para in tous_les_paragraphes(doc):
        txt = para.text.upper()
        if numero is None and ("CR DE REUNION" in txt or "CR DE RÉUNION" in txt):
            m = re.search(r'N°\s*(\d+)', txt)
            if m:
                numero = int(m.group(1))
        if heure is None and "PROCHAIN RDV" in txt:
            m = re.search(r'(\d{1,2})\s*[H:]\s*(\d{2})?', txt)
            if m:
                heure = f"{int(m.group(1)):02d}h{m.group(2) or '00'}"
    return numero, heure


def formater_rdv(jour, heure):
    """Ex : MARDI 03 NOVEMBRE 2026 à 09h00"""
    return f"{JOURS_FR[jour.weekday()]} {jour.day:02d} {MOIS_FR[jour.month - 1]} {jour.year} à {heure}"


def entreprises_connues(doc, index_tables):
    """Noms courts utilisés dans la colonne ACTION PAR (ROC, 3G, MOA, MOEX…), du plus fréquent au moins fréquent."""
    valeurs = []
    for i in index_tables:
        table = doc.tables[i]
        for row in table.rows[1:]:
            tcs = row._tr.tc_lst
            roles = colonnes_du_tableau(table, len(tcs))
            if "responsable" in roles:
                texte = " ".join(_Cell(tcs[roles["responsable"]], table).text.split())
                # « ROC/ MOEX - MOA » → ROC, MOEX, MOA
                valeurs += [v.strip() for v in re.split(r'\s*[/,]\s*|\s+-\s+', texte) if v.strip()]
    simples = set(valeurs)
    noms = []
    for v in valeurs:
        mots = v.split()
        # « MOA MOEX » = deux noms déjà connus ; « ROC FACADE » reste un nom
        morceaux = mots if len(mots) > 1 and all(m in simples for m in mots) else [v]
        noms += [m for m in morceaux if m.upper() not in MOTS_EXCLUS and not m.isdigit()]
    return sorted(set(noms), key=lambda n: -noms.count(n))


def apercu_sections(doc, sections, nb=4):
    """Derniers points de chaque section, pour que l'IA voie de quoi parle chacune."""
    apercu = {}
    for nom, i in sections.items():
        textes = [" ".join(r.cells[0].text.split()) for r in doc.tables[i].rows[1:]]
        apercu[nom] = [t[:150] for t in textes if t][-nb:]
    return apercu


def intervenants_du_chantier(doc):
    """Tableau des intervenants (en-tête DESIGNATION) : [{"lot": ..., "entreprise": ...}].
    Entreprise vide = pas encore désignée."""
    lots = {}
    for table in doc.tables:
        if not table.rows or "DESIGNATION" not in table.rows[0].cells[0].text.upper():
            continue
        for row in table.rows[1:]:
            cells = row.cells
            if len(cells) < 2:
                continue
            lot = " ".join(cells[0].text.split())
            if lot and lot not in lots:
                # 1re ligne du lot = nom de l'entreprise (les suivantes : adresse, contacts)
                lots[lot] = " ".join(cells[1].text.split())
    return [{"lot": lot, "entreprise": ent} for lot, ent in lots.items()]


def repasser_en_noir(doc):
    """Les ajouts bleus du CR précédent deviennent de l'historique (noir)."""
    nb = 0
    for para in tous_les_paragraphes(doc):
        for run in para.runs:
            if run.font.color.type is not None and run.font.color.rgb == BLEU:
                run.font.color.rgb = NOIR
                nb += 1
    return nb


def lire_date(texte):
    """« 31-03-26 », « 31/03/2026 »… en début de texte → date, sinon None."""
    m = re.match(r'\s*(\d{1,2})[-/](\d{1,2})[-/](\d{4}|\d{2})\b', texte or "")
    if not m:
        return None
    jour, mois, annee = (int(x) for x in m.groups())
    try:
        return date(annee if annee > 99 else 2000 + annee, mois, jour)
    except ValueError:
        return None


def cellules_du_point(doc, pid):
    """« T12-L5 » → (tableau, cellules réelles de la ligne, rôles des colonnes)."""
    i, j = (int(x) for x in re.findall(r'\d+', pid))
    table = doc.tables[i]
    tcs = table.rows[j]._tr.tc_lst
    return table, tcs, colonnes_du_tableau(table, len(tcs))


def points_ouverts(doc, sections):
    """Actions du CR précédent pas encore soldées (FAIT LE vide, RAPPEL, EN ATT, EN COURS…).
    Les points PM (pour mémoire) ne sont pas suivis, sauf s'ils sont déjà en rappel."""
    ouverts = []
    for nom, i in sections.items():
        table = doc.tables[i]
        for j in range(1, len(table.rows)):
            pid = f"T{i}-L{j}"
            _, tcs, roles = cellules_du_point(doc, pid)
            if not {"texte", "pour_le", "fait_le"} <= roles.keys() or max(roles.values()) >= len(tcs):
                continue
            val = {r: " ".join(_Cell(tcs[idx], table).text.split()) for r, idx in roles.items()}
            fait, pour = val["fait_le"].upper(), val["pour_le"].upper()
            if not val["texte"] or fait == "FAIT LE":  # ligne vide ou en-tête répété
                continue
            if lire_date(fait) or fait in ("FAIT", "NON FAIT", "PM") or fait.startswith("SANS OBJET"):
                continue
            if pour == "PM" and not fait.startswith(("RAPPEL", "EN ")):
                continue
            ouverts.append({"id": pid, "section": nom, "texte": val["texte"],
                            "responsable": val.get("responsable", ""), "note_le": val.get("note_le", ""),
                            "pour_le": val["pour_le"], "fait_le": val["fait_le"]})
    return ouverts


def retard_jours(point, jour):
    """Jours de retard à la date du CR si l'échéance est dépassée et le point sans suite (vide ou RAPPEL)."""
    fait = point["fait_le"].upper()
    if fait and not fait.startswith("RAPPEL"):
        return None
    debut = lire_date(point["pour_le"])
    if debut is None and point["pour_le"].upper() == "SANS DELAI" and "JOURS" in fait:
        debut = lire_date(point["note_le"])  # « SANS DELAI » : retard compté depuis la note
    if debut is None or debut >= jour:
        return None
    return (jour - debut).days


def maj_cellule(doc, pid, role, texte, couleur=BLEU, ajouter=False):
    """Écrit dans une colonne d'une ligne existante du CR."""
    table, tcs, roles = cellules_du_point(doc, pid)
    if role in roles and roles[role] < len(tcs):
        ecrire_dans_cellule(_Cell(tcs[roles[role]], table), texte,
                            gras_majuscules=(role == "texte"), couleur=couleur, ajouter=ajouter)


# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Étape 1 — Upload du CR précédent")
uploaded_file = st.file_uploader(
    "Dépose le fichier Word du CR précédent (.docx)",
    type=["docx"]
)

if uploaded_file:
    doc_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(doc_bytes))

    # Détecter les sections (tables ≥5 colonnes, hors DESIGNATION)
    sections = {}
    for i, table in enumerate(doc.tables):
        if len(table.columns) >= 5 and len(table.rows) >= 1:
            en_tete = table.rows[0].cells[0].text.strip()
            if en_tete and len(en_tete) > 2 and "DESIGNATION" not in en_tete.upper():
                label = f"{en_tete[:60]} (tableau {i+1})"
                sections[label] = i

    num_prec, heure_prec = infos_cr_precedent(doc)
    intervenants = intervenants_du_chantier(doc)
    apercu = apercu_sections(doc, sections)
    ouverts = points_ouverts(doc, sections)
    entreprises = entreprises_connues(doc, sections.values())

    st.success(f"✅ {len(doc.tables)} tableaux trouvés · {len(sections)} sections détectées")

    # ── Infos CR ─────────────────────────────────────────────────────────────
    st.markdown("### 📋 Étape 2 — Informations du nouveau CR")
    c1, c2, c3 = st.columns(3)
    with c1:
        num_cr = st.text_input(
            "Numéro du CR",
            value=str(num_prec + 1) if num_prec is not None else "",
            placeholder="Ex : 29",
            help=f"CR précédent : N°{num_prec}" if num_prec is not None else None,
        )
    with c2:
        date_cr = st.date_input("Date de la réunion", value=date.today(), format="DD/MM/YYYY")
    with c3:
        # Par défaut : une semaine plus tard, même heure que le RDV précédent
        prochain = st.text_input(
            "Prochain RDV",
            value=formater_rdv(date_cr + timedelta(days=7), heure_prec or "09h00"),
        )

    with st.expander(f"👷 Qui fait quoi — {len(intervenants)} lots / intervenants lus dans le CR"):
        st.caption("Entreprise vide = pas encore désignée. Tu peux corriger ou compléter.")
        intervenants_edites = st.data_editor(
            intervenants,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "lot": st.column_config.TextColumn("Lot / rôle", width="medium"),
                "entreprise": st.column_config.TextColumn("Entreprise", width="large"),
            },
        )
    entreprises_txt = st.text_input(
        "Noms courts utilisés dans la colonne ACTION PAR",
        value=", ".join(entreprises),
        help="L'IA utilise ces noms courts (ex : ROC plutôt que ROC - ROCHEFOLLE CONSTRUCTIONS)",
    )
    repasser_noir = st.checkbox(
        "Repasser en noir les ajouts bleus du CR précédent", value=True
    )
    retards = [(p, n) for p in ouverts if (n := retard_jours(p, date_cr))]
    maj_retards = st.checkbox(
        f"⏰ Mettre à jour les {len(retards)} échéances dépassées (FAIT LE → « RAPPEL n JOURS DE RETARD » en rouge)",
        value=True,
    )
    with st.expander(f"Voir les {len(ouverts)} points ouverts du CR précédent, dont {len(retards)} en retard"):
        st.dataframe(
            [{"Section": p["section"], "Texte": p["texte"], "Action par": p["responsable"],
              "Pour le": p["pour_le"], "Fait le": p["fait_le"],
              "Retard": f"{n} j" if (n := retard_jours(p, date_cr)) else ""} for p in ouverts],
            use_container_width=True,
        )

    # ── Section ───────────────────────────────────────────────────────────────
    st.markdown("### 🎯 Étape 3 — Sections du CR")
    if sections:
        st.caption("L'IA range chaque point dans la bonne section : " + " · ".join(sections))
    else:
        st.warning("⚠️ Aucune section détectée.")

    # ── Notes ─────────────────────────────────────────────────────────────────
    st.markdown("### ✏️ Étape 4 — Notes brutes")
    notes = st.text_area(
        "Tes notes de réunion (toutes sections confondues) :",
        height=250,
        placeholder="- serbois retard bat3\n- bazzi peinture 80% bat1\n- enedis raccordement prévu 3 avril"
    )

    # ── Génération ────────────────────────────────────────────────────────────
    if st.button("🤖 Générer les points CR avec l'IA", type="primary", use_container_width=True):
        if not api_key:
            st.error("⚠️ Entre ta clé API dans la barre latérale")
        elif not notes.strip():
            st.error("⚠️ Saisis tes notes")
        elif not sections:
            st.error("⚠️ Aucune section où insérer les points")
        else:
            with st.spinner("L'IA reformule tes notes..."):
                try:
                    client = anthropic.Anthropic(api_key=api_key)
                    prompt = f"""Tu es assistant MOE chez MH2O, ingénierie du bâtiment.
Reformule ces notes brutes en points de CR professionnels et range chacun dans la bonne section du CR.

RÈGLES :
- Chaque point = une phrase courte et factuelle
- Noms d'entreprises en MAJUSCULES (SERBOIS, BAZZI, ENEDIS, ROC...)
- Action par : l'entreprise ou l'intervenant qui doit agir (ex : SERBOIS, MH2O)
- Pour le : échéance au format JJ-MM-AA si les notes en donnent une (« pour le 6/11 », « avant vendredi »…), sinon vide
- Fait le : vide si l'action n'est pas entièrement terminée (« reste la façade » = pas fait) ;
  si elle est terminée, la date de réalisation au format JJ-MM-AA quand les notes la précisent (« fait le 20/10 », « fait lundi »…), sinon FAIT
- Le texte ne contient pas le responsable entre parenthèses ; la date d'un événement (réunion, livraison, visite…) reste dans le texte,
  « pour le » est réservé à l'échéance d'une action demandée
- Section : traite chaque note indépendamment, sans reprendre l'entreprise ou la section de la note précédente
  · une section « Lot … – Entreprise X » seulement si la note concerne cette entreprise ou ce corps d'état
  · une note qui ne vise pas une entreprise précise (réunion échantillons, réunion de coordination, visite, planning général…)
    va dans la section transversale où ce sujet est déjà traité (MOEX, OPC, PLANNING, MOA…) : regarde les derniers points de chaque section
- Entreprises : retrouve-les dans la liste « QUI FAIT QUOI » ci-dessous, même si la note est approximative
- Écris-les avec le nom court déjà utilisé dans le CR (texte et action par) : {entreprises_txt or "aucun"}
  (ex : ROC plutôt que ROC - ROCHEFOLLE CONSTRUCTIONS, 3G plutôt que 3G CONSTRUCTION) ; sinon le nom de la liste
- Si une note désigne un corps d'état (« le peintre », « l'électricien », « le plombier »…), retrouve l'entreprise du lot correspondant
- Si un LOT n'a pas d'entreprise, elle n'est pas encore désignée : n'invente aucun nom, écris « l'entreprise du LOT XX (non désignée) » et mets le lot en action par (ex : LOT 19)
- Pour les lignes sans numéro de lot (concessionnaires : ENEDIS, FRANCE TELECOM, GAZ…), le nom de la ligne désigne l'intervenant (ex : ENEDIS)
- POINTS OUVERTS : si une note fait avancer ou solde un point ouvert listé plus bas, ne crée PAS de nouveau point ;
  ajoute une entrée dans « mises_a_jour » avec l'id du point (seulement pour les points clairement concernés) :
  · fait_le : FAIT, ou la date JJ-MM-AA si les notes la précisent, quand c'est terminé ; EN COURS ou EN ATT pour un simple avancement ; sinon vide (inchangé)
  · pour_le : nouvelle échéance JJ-MM-AA si les notes en donnent une, sinon vide (inchangé)
  · complement : courte phrase factuelle ajoutée à la fin du texte existant (ex : « ROC indique une livraison le 12-11-26. »), sinon vide

EXEMPLES (texte → action par, pour le, fait le) :
MH2O demande à SERBOIS de transmettre un planning actualisé pour le Bâtiment 3. → SERBOIS, 06-11-26, (vide)
Raccordement ENEDIS prévu le 03-04-26, confirmation attendue. → MH2O, (vide), (vide)
Retouches peinture cage d'escalier A effectuées par BAZZI. → BAZZI, (vide), FAIT
Reprise des enduits Bâtiment 2 réalisée par SERBOIS le 20-10-26. → SERBOIS, (vide), 20-10-26

QUI FAIT QUOI (lot / rôle → entreprise) :
{chr(10).join("- " + (i.get("lot") or "") + " → " + ((i.get("entreprise") or "").strip() or "NON DÉSIGNÉE") for i in intervenants_edites if i.get("lot"))}

SECTIONS DISPONIBLES (avec leurs derniers points, pour savoir de quoi parle chacune) :
{chr(10).join("- " + nom + "".join(chr(10) + "    · " + pt for pt in apercu[nom]) for nom in sections)}

POINTS OUVERTS DU CR PRÉCÉDENT ([id] section : texte | action par | pour le | fait le) :
{chr(10).join(f"[{p['id']}] {p['section']} : {p['texte'][:200]} | {p['responsable']} | {p['pour_le']} | {p['fait_le']}" for p in ouverts) or "aucun"}

Notes :
{notes}

CR n° : {num_cr} — Date du jour : {JOURS_FR[date_cr.weekday()]} {date_cr.strftime('%d-%m-%y')}"""

                    # L'IA remplit un « formulaire » : chaque colonne du tableau séparément.
                    # Le format JSON est garanti par l'API (sortie structurée).
                    schema = {
                        "type": "object",
                        "properties": {
                            "points": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "section": {"type": "string", "enum": list(sections)},
                                        "texte": {"type": "string"},
                                        "responsable": {"type": "string"},
                                        "pour_le": {"type": "string", "description": "Échéance JJ-MM-AA ou vide"},
                                        "fait_le": {"type": "string", "description": "Vide si pas terminé, date JJ-MM-AA si précisée, sinon FAIT"},
                                    },
                                    "required": ["section", "texte", "responsable", "pour_le", "fait_le"],
                                    "additionalProperties": False,
                                },
                            },
                            "mises_a_jour": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "id": {"type": "string", "enum": [p["id"] for p in ouverts]} if ouverts else {"type": "string"},
                                        "fait_le": {"type": "string", "description": "FAIT, JJ-MM-AA, EN COURS, EN ATT ou vide"},
                                        "pour_le": {"type": "string", "description": "Nouvelle échéance JJ-MM-AA ou vide"},
                                        "complement": {"type": "string", "description": "Phrase ajoutée au texte existant, ou vide"},
                                    },
                                    "required": ["id", "fait_le", "pour_le", "complement"],
                                    "additionalProperties": False,
                                },
                            },
                        },
                        "required": ["points", "mises_a_jour"],
                        "additionalProperties": False,
                    }
                    msg = client.beta.messages.create(
                        model="claude-sonnet-5-5",
                        max_tokens=16000,
                        # Effort bas : tâche de mise en forme, on limite la réflexion (et le coût)
                        output_config={"effort": "low", "format": {"type": "json_schema", "schema": schema}},
                        # Si le modèle refuse la demande, l'API la relance sur un modèle de secours
                        betas=["server-side-fallback-2026-07-01"],
                        fallbacks="default",
                        messages=[{"role": "user", "content": prompt}]
                    )
                    if msg.stop_reason == "refusal":
                        st.error("⚠️ L'IA a refusé de traiter ces notes. Reformule-les puis réessaie.")
                    elif msg.stop_reason == "max_tokens":
                        st.error("⚠️ Réponse de l'IA tronquée : découpe tes notes en deux et réessaie.")
                    else:
                        texte_json = next(b.text for b in msg.content if b.type == "text")
                        reponse = json.loads(texte_json)
                        st.session_state.pop("points_edites", None)
                        st.session_state.pop("maj_editees", None)
                        par_id = {p["id"]: p for p in ouverts}
                        st.session_state.update({
                            "points": reponse["points"],
                            "ouverts": ouverts,
                            # Texte d'origine affiché à côté de chaque mise à jour proposée
                            "mises_a_jour": [
                                {"id": m["id"], "section": par_id[m["id"]]["section"], "texte": par_id[m["id"]]["texte"],
                                 "fait_le": m["fait_le"], "pour_le": m["pour_le"], "complement": m["complement"]}
                                for m in reponse["mises_a_jour"] if m["id"] in par_id
                            ],
                            "sections": sections,
                            "num_cr": num_cr,
                            "date_cr": date_cr,
                            "prochain": prochain,
                            "doc_bytes": doc_bytes,
                        })
                except Exception as e:
                    st.error(f"Erreur API : {e}")

    # ── Résultat & export ─────────────────────────────────────────────────────
    if "points" in st.session_state:
        st.markdown("### ✅ Étape 5 — Points générés")
        st.caption("Vérifie et corrige directement dans le tableau (ajout / suppression de lignes possibles).")
        points_edites = st.data_editor(
            st.session_state["points"],
            key="points_edites",
            num_rows="dynamic",
            use_container_width=True,
            column_order=["section", "texte", "responsable", "pour_le", "fait_le"],
            column_config={
                "section": st.column_config.SelectboxColumn(
                    "Section", options=list(st.session_state["sections"]), required=True, width="medium"
                ),
                "texte": st.column_config.TextColumn("Texte", width="large"),
                "responsable": st.column_config.TextColumn("Action par", width="small"),
                "pour_le": st.column_config.TextColumn("Pour le", width="small", help="Échéance JJ-MM-AA"),
                "fait_le": st.column_config.TextColumn("Fait le", width="small", help="Vide, FAIT ou date JJ-MM-AA"),
            },
        )

        maj_editees = []
        if st.session_state.get("mises_a_jour"):
            st.markdown("#### 🔄 Mises à jour des points du CR précédent")
            st.caption("Ces points existants seront complétés en bleu. Supprime une ligne pour ignorer la mise à jour.")
            maj_editees = st.data_editor(
                st.session_state["mises_a_jour"],
                key="maj_editees",
                num_rows="dynamic",
                use_container_width=True,
                column_order=["section", "texte", "fait_le", "pour_le", "complement"],
                disabled=["section", "texte"],
                column_config={
                    "section": st.column_config.TextColumn("Section", width="small"),
                    "texte": st.column_config.TextColumn("Point existant", width="large"),
                    "fait_le": st.column_config.TextColumn("Fait le", width="small", help="FAIT, JJ-MM-AA, EN COURS, EN ATT ou vide"),
                    "pour_le": st.column_config.TextColumn("Nouvelle échéance", width="small"),
                    "complement": st.column_config.TextColumn("Complément ajouté au texte", width="large"),
                },
            )

        st.markdown("### 📥 Étape 6 — Générer le Word")
        if st.button("📄 Créer le nouveau CR Word", type="primary", use_container_width=True):
            with st.spinner("Insertion dans le document..."):
                try:
                    doc2 = Document(io.BytesIO(st.session_state["doc_bytes"]))
                    num = st.session_state["num_cr"]
                    dt = st.session_state["date_cr"]
                    rdv = st.session_state["prochain"]
                    smap = st.session_state["sections"]
                    date_cr_fmt = dt.strftime("%d-%m-%y")  # format du CR : 28-10-26

                    # Jours en français pour le titre
                    date_titre = f"{JOURS_FR[dt.weekday()]} {dt.day} {MOIS_FR[dt.month - 1]} {dt.year}"

                    # Avant d'ajouter le nouveau bleu, l'ancien passe en noir
                    if repasser_noir:
                        nb_noir = repasser_en_noir(doc2)
                        st.info(f"🖤 {nb_noir} passages bleus du CR précédent repassés en noir")

                    # Points ouverts : retards recalculés, puis mises à jour validées (prioritaires)
                    ouverts_par_id = {p["id"]: p for p in st.session_state.get("ouverts", [])}
                    nb_retards = 0
                    if maj_retards:
                        for p in ouverts_par_id.values():
                            if n := retard_jours(p, dt):
                                maj_cellule(doc2, p["id"], "fait_le", f"RAPPEL {n} JOURS DE RETARD", couleur=ROUGE)
                                nb_retards += 1
                    nb_maj = 0
                    for m in maj_editees:
                        if m.get("id") not in ouverts_par_id:
                            continue
                        modifie = False
                        for role in ("fait_le", "pour_le"):
                            valeur = (m.get(role) or "").strip().replace("/", "-")
                            if valeur:
                                maj_cellule(doc2, m["id"], role, valeur)
                                modifie = True
                        complement = (m.get("complement") or "").strip()
                        if complement:
                            maj_cellule(doc2, m["id"], "texte", complement, ajouter=True)
                            modifie = True
                        nb_maj += modifie

                    for para in tous_les_paragraphes(doc2):
                        txt = para.text.upper()

                        # Mise à jour numéro + date dans le titre principal
                        # Travail sur le texte complet : Word découpe souvent le titre en plusieurs runs
                        if "CR DE REUNION" in txt or "CR DE RÉUNION" in txt:
                            full_text = para.text
                            new_text = full_text
                            # Remplacer le numéro (N°28 → N°{num})
                            if num:
                                new_text = re.sub(r'N°\s*\d+', f'N°{num}', new_text)
                            # Remplacer "MARDI 20 OCTOBRE 2026" (avec ou sans DU devant)
                            new_text = re.sub(
                                r'(LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE)\s+\d{1,2}\s+\w+\s+\d{4}',
                                date_titre, new_text, flags=re.IGNORECASE
                            )
                            if new_text != full_text:
                                remplacer_texte_paragraphe(para, new_text)

                        # Mise à jour prochain RDV
                        if "PROCHAIN RDV" in txt and rdv:
                            full_text = para.text
                            # Remplacer tout ce qui suit "le "
                            new_text = re.sub(
                                r'(?<=\ble\s).*',
                                rdv, full_text, flags=re.IGNORECASE | re.DOTALL
                            )
                            if new_text == full_text:
                                # Fallback : depuis le premier jour de semaine
                                new_text = re.sub(
                                    r'(LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI|SAMEDI|DIMANCHE).*',
                                    rdv, full_text, flags=re.IGNORECASE | re.DOTALL
                                )
                            if new_text != full_text:
                                remplacer_texte_paragraphe(para, new_text)

                    # Insertion dans le tableau
                    nb_ajouts = 0
                    modeles = {}  # une ligne d'origine par tableau, choisie avant tout ajout
                    for point in points_edites:
                        texte = (point.get("texte") or "").strip()
                        if not texte or point.get("section") not in smap:
                            continue
                        table = doc2.tables[smap[point["section"]]]
                        modele = modeles.setdefault(smap[point["section"]], ligne_modele(table))
                        cells = ajouter_ligne(table, modele)
                        roles = colonnes_du_tableau(table, len(cells))
                        valeurs = {
                            "texte": texte,
                            "responsable": (point.get("responsable") or "").strip(),
                            "note_le": date_cr_fmt,
                            "pour_le": (point.get("pour_le") or "").strip().replace("/", "-"),
                            "fait_le": (point.get("fait_le") or "").strip().replace("/", "-"),
                        }
                        # Toutes les cellules copiées sont vidées, puis remplies selon leur colonne
                        for i, cell in enumerate(cells):
                            role = next((r for r, idx in roles.items() if idx == i), None)
                            ecrire_dans_cellule(cell, valeurs.get(role, ""), gras_majuscules=(role == "texte"))
                        nb_ajouts += 1

                    output = io.BytesIO()
                    doc2.save(output)
                    output.seek(0)
                    nom = f"CR_CHANTIER_N{num}_{dt.strftime('%d-%m-%Y')}.docx"
                    st.download_button(
                        f"⬇️ Télécharger {nom}",
                        data=output,
                        file_name=nom,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    st.success(
                        f"✅ CR généré : {nb_ajouts} nouveaux points, {nb_maj} points mis à jour, "
                        f"{nb_retards} retards recalculés. Télécharge le fichier ci-dessus."
                    )

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())

else:
    st.info("👆 Upload le fichier Word du CR précédent pour commencer")
