import streamlit as st
import anthropic
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
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

# ── Fonctions ─────────────────────────────────────────────────────────────────

def ecrire_cellule_texte(cell, texte, nouveau=True):
    """Calibri 8pt, bleu si nouveau, entreprises (MAJUSCULES) en gras."""
    para = cell.paragraphs[0]
    # Vider les runs existants
    for run in para.runs:
        run.text = ""
    # Découper : mots tout en majuscules vs reste
    tokens = re.split(r'(\b[A-Z]{2}[A-Z0-9\-&/]*\b)', texte)
    couleur = BLEU if nouveau else NOIR
    for token in tokens:
        if not token:
            continue
        run = para.add_run(token)
        run.font.name = "Calibri"
        run.font.size = Pt(8)
        run.font.color.rgb = couleur
        run.bold = bool(re.match(r'^[A-Z]{2}[A-Z0-9\-&/]*$', token))

def ecrire_cellule_simple(cell, texte, nouveau=True):
    """Calibri 8pt, bleu si nouveau."""
    para = cell.paragraphs[0]
    for run in para.runs:
        run.text = ""
    run = para.add_run(texte)
    run.font.name = "Calibri"
    run.font.size = Pt(8)
    run.font.color.rgb = BLEU if nouveau else NOIR

JOURS_FR = ["LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI", "DIMANCHE"]
MOIS_FR = ["JANVIER", "FÉVRIER", "MARS", "AVRIL", "MAI", "JUIN",
           "JUILLET", "AOÛT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DÉCEMBRE"]

# Mots en majuscules qui ne sont pas des entreprises
MOTS_EXCLUS = {
    "PA", "PM", "EN", "ATT", "ATTENTE", "FAIT", "CR", "RDV", "BAT", "NB", "OK",
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
    """Noms en MAJUSCULES déjà présents dans les tableaux de sections."""
    noms = set()
    for i in index_tables:
        for row in doc.tables[i].rows[1:]:
            cells = row.cells
            # Colonne responsable : tout est pris ; colonne texte : seulement le gras
            if len(cells) > 1:
                noms.update(re.findall(r'\b[A-Z][A-Z0-9\-&]{1,}\b', cells[1].text))
            for para in cells[0].paragraphs:
                for run in para.runs:
                    if run.bold:
                        noms.update(re.findall(r'\b[A-Z][A-Z0-9\-&]{1,}\b', run.text))
    return sorted(n for n in noms if n not in MOTS_EXCLUS and not n.isdigit())


def repasser_en_noir(doc):
    """Les ajouts bleus du CR précédent deviennent de l'historique (noir)."""
    nb = 0
    for para in tous_les_paragraphes(doc):
        for run in para.runs:
            if run.font.color.type is not None and run.font.color.rgb == BLEU:
                run.font.color.rgb = NOIR
                nb += 1
    return nb


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

    entreprises_txt = st.text_input(
        "Entreprises connues (orthographe imposée à l'IA)",
        value=", ".join(entreprises),
        help="Détectées dans le CR précédent — ajoute ou corrige si besoin",
    )
    repasser_noir = st.checkbox(
        "Repasser en noir les ajouts bleus du CR précédent", value=True
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
- Statut : PA (Pour Action), PM (Pour Mémoire), EN ATT (En Attente), FAIT, ou date DD/MM/YY si échéance
- Responsable : l'entreprise ou l'intervenant concerné (ex : SERBOIS, MH2O)
- Le texte ne contient ni le statut ni le responsable entre parenthèses : ils vont dans leurs propres champs
- Section : choisis celle qui correspond le mieux au sujet du point
- Entreprises déjà connues sur ce chantier (respecte exactement cette orthographe, même si la note est approximative) : {entreprises_txt or "aucune"}

EXEMPLES (texte → statut, responsable) :
MH2O demande à SERBOIS de transmettre un planning actualisé pour le Bâtiment 3. → PA, SERBOIS
Raccordement ENEDIS prévu le 03/04/26, confirmation attendue. → EN ATT, MH2O
Avancement peinture BAZZI : 80% Bâtiment 1, retard de 2 semaines. → PA, BAZZI

SECTIONS DISPONIBLES :
{chr(10).join("- " + nom for nom in sections)}

Notes :
{notes}

CR n° : {num_cr} — Date : {date_cr.strftime('%d/%m/%Y')}"""

                    # L'IA remplit un « formulaire » : chaque colonne du tableau séparément
                    outil = {
                        "name": "enregistrer_points",
                        "description": "Enregistre les points du CR, un par ligne de tableau.",
                        "input_schema": {
                            "type": "object",
                            "properties": {
                                "points": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "section": {"type": "string", "enum": list(sections)},
                                            "texte": {"type": "string"},
                                            "statut": {"type": "string", "description": "PA, PM, EN ATT, FAIT ou date DD/MM/YY"},
                                            "responsable": {"type": "string"},
                                        },
                                        "required": ["section", "texte", "statut", "responsable"],
                                    },
                                }
                            },
                            "required": ["points"],
                        },
                    }
                    msg = client.messages.create(
                        model="claude-sonnet-4-6",
                        max_tokens=4000,
                        tools=[outil],
                        tool_choice={"type": "tool", "name": "enregistrer_points"},
                        messages=[{"role": "user", "content": prompt}]
                    )
                    reponse = next(b.input for b in msg.content if b.type == "tool_use")
                    st.session_state.pop("points_edites", None)
                    st.session_state.update({
                        "points": reponse["points"],
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
            column_order=["section", "texte", "statut", "responsable"],
            column_config={
                "section": st.column_config.SelectboxColumn(
                    "Section", options=list(st.session_state["sections"]), required=True, width="medium"
                ),
                "texte": st.column_config.TextColumn("Texte", width="large"),
                "statut": st.column_config.TextColumn("Statut", width="small"),
                "responsable": st.column_config.TextColumn("Responsable", width="small"),
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
                    date_fr = dt.strftime("%d/%m/%Y")

                    # Jours en français pour le titre
                    date_titre = f"{JOURS_FR[dt.weekday()]} {dt.day} {MOIS_FR[dt.month - 1]} {dt.year}"

                    # Avant d'ajouter le nouveau bleu, l'ancien passe en noir
                    if repasser_noir:
                        nb_noir = repasser_en_noir(doc2)
                        st.info(f"🖤 {nb_noir} passages bleus du CR précédent repassés en noir")

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
                    for point in points_edites:
                        texte = (point.get("texte") or "").strip()
                        if not texte or point.get("section") not in smap:
                            continue
                        statut = (point.get("statut") or "").strip()
                        resp = (point.get("responsable") or "").strip()
                        table = doc2.tables[smap[point["section"]]]
                        nb_cols = len(table.columns)
                        row = table.add_row()
                        cells = row.cells

                        ecrire_cellule_texte(cells[0], texte, nouveau=True)
                        if nb_cols > 1:
                            ecrire_cellule_simple(cells[1], resp, nouveau=True)
                        if nb_cols > 2:
                            ecrire_cellule_simple(cells[2], date_fr, nouveau=True)
                        if nb_cols > 3:
                            ecrire_cellule_simple(cells[3], statut, nouveau=True)
                        if nb_cols > 4:
                            ecrire_cellule_simple(cells[4], "", nouveau=False)
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
                    st.success(f"✅ CR généré ({nb_ajouts} points insérés) ! Télécharge le fichier ci-dessus.")

                except Exception as e:
                    st.error(f"Erreur : {e}")
                    import traceback
                    st.code(traceback.format_exc())

else:
    st.info("👆 Upload le fichier Word du CR précédent pour commencer")
