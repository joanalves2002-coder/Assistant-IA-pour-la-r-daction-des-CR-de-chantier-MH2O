import streamlit as st
import anthropic
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import copy
import io
from datetime import date
import locale

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
    .section-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #dde3ef;
        margin-bottom: 16px;
    }
    .result-box {
        background: #f0f4ff;
        border-left: 4px solid #1a4480;
        padding: 16px;
        border-radius: 8px;
        font-family: monospace;
        white-space: pre-wrap;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
    <h1>🏗️ Assistant CR <span class="accent">MH2O</span></h1>
    <p>Reformulation de notes de chantier et insertion automatique dans le Word du CR précédent</p>
</div>
""", unsafe_allow_html=True)

# ── Clé API ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    api_key = st.text_input("Clé API Claude", type="password", placeholder="sk-ant-...")
    st.markdown("---")
    st.markdown("### 📖 Comment utiliser")
    st.markdown("""
1. Entre ta clé API Claude
2. Upload le Word du CR précédent
3. Renseigne les infos du nouveau CR
4. Choisis la section à mettre à jour
5. Saisis tes notes brutes
6. Génère et télécharge le nouveau CR
    """)

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Étape 1 — Upload du CR précédent")
uploaded_file = st.file_uploader(
    "Dépose ici le fichier Word du CR précédent (.docx)",
    type=["docx"],
    help="Le fichier sera modifié pour créer le nouveau CR"
)

if uploaded_file:
    doc_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(doc_bytes))

    # Détecter les tables de contenu (5 colonnes = tables OPC/Entreprises)
    sections = {}
    for i, table in enumerate(doc.tables):
        if len(table.columns) >= 5 and len(table.rows) >= 2:
            header = table.rows[0].cells[0].text.strip()
            if header and len(header) > 3 and "DESIGNATION" not in header:
                sections[header[:80]] = i

    st.success(f"✅ Fichier chargé — {len(doc.paragraphs)} paragraphes, {len(doc.tables)} tableaux détectés")

    # ── Infos nouveau CR ──────────────────────────────────────────────────────
    st.markdown("### 📋 Étape 2 — Informations du nouveau CR")
    col1, col2, col3 = st.columns(3)
    with col1:
        nouveau_numero = st.text_input("Numéro du CR", placeholder="Ex : 29")
    with col2:
        nouvelle_date = st.date_input("Date de la réunion", value=date.today())
    with col3:
        prochain_rdv = st.text_input("Prochain RDV", placeholder="Ex : MARDI 03 NOVEMBRE 2026 à 09h00")

    # ── Sélection section ─────────────────────────────────────────────────────
    st.markdown("### 🎯 Étape 3 — Section à mettre à jour")

    if sections:
        section_choisie = st.selectbox(
            "Quelle section veux-tu mettre à jour ?",
            options=list(sections.keys()),
            help="Sélectionne la section correspondant à tes notes"
        )
    else:
        st.warning("Aucune section détectée automatiquement.")
        section_choisie = st.text_input("Entre le nom de la section manuellement")

    # ── Notes brutes ──────────────────────────────────────────────────────────
    st.markdown("### ✏️ Étape 4 — Notes brutes de la réunion")
    notes_brutes = st.text_area(
        f"Tes notes pour la section : **{section_choisie}**",
        height=200,
        placeholder="""Exemple :
- serbois retard menuiserie bat3 - relancé - attente planning
- bazzi peinture 80% bat1 - 2 semaines de retard
- raccordement enedis prévu 3 avril - à confirmer
- demander à ROC le planning actualisé GO avant vendredi"""
    )

    # ── Génération ────────────────────────────────────────────────────────────
    if st.button("🤖 Générer les points CR", type="primary", use_container_width=True):
        if not api_key:
            st.error("⚠️ Entre ta clé API Claude dans la barre latérale")
        elif not notes_brutes.strip():
            st.error("⚠️ Saisis des notes à reformuler")
        else:
            with st.spinner("L'IA reformule tes notes..."):
                try:
                    client = anthropic.Anthropic(api_key=api_key)

                    prompt = f"""Tu es un assistant de maîtrise d'œuvre chez MH2O, entreprise d'ingénierie du bâtiment basée à Boulogne-Billancourt.
Tu dois reformuler des notes brutes de réunion de chantier en points de CR professionnels.

Règles de formatage strictes :
- Chaque point est une phrase courte et factuelle
- Utilise les abréviations standards : PA (Pour Action), PM (Pour Mémoire), EN ATT (En Attente), FAIT
- Le responsable est indiqué en fin de ligne entre parenthèses
- Format de chaque ligne : texte du point. STATUT (RESPONSABLE)
- Ton neutre et professionnel, style MH2O
- Ne jamais inventer d'informations non présentes dans les notes
- Commence directement par les points, sans introduction

Exemples de style attendu :
"MH2O demande à SERBOIS de transmettre un planning d'exécution actualisé pour le Bâtiment 3. PA (SERBOIS)"
"Raccordement ENEDIS programmé le 03/04/2026 — confirmation définitive attendue. EN ATT (MH2O)"
"Avancement peinture BAZZI : 80% Bâtiment 1 — retard de 2 semaines constaté, accélération demandée. PA (BAZZI)"

Notes brutes à reformuler :
{notes_brutes}

Section du CR : {section_choisie}
Numéro du CR : {nouveau_numero}
Date : {nouvelle_date.strftime('%d/%m/%Y')}"""

                    message = client.messages.create(
                        model="claude-sonnet-4-6",
                        max_tokens=1000,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    points_generes = message.content[0].text
                    st.session_state["points_generes"] = points_generes
                    st.session_state["section_choisie"] = section_choisie
                    st.session_state["sections"] = sections
                    st.session_state["nouveau_numero"] = nouveau_numero
                    st.session_state["nouvelle_date"] = nouvelle_date
                    st.session_state["prochain_rdv"] = prochain_rdv
                    st.session_state["doc_bytes"] = doc_bytes

                except Exception as e:
                    st.error(f"Erreur API : {e}")

    # ── Résultat ──────────────────────────────────────────────────────────────
    if "points_generes" in st.session_state:
        st.markdown("### ✅ Points générés par l'IA")
        points = st.session_state["points_generes"]

        edited_points = st.text_area(
            "Vérifie et modifie si nécessaire avant insertion :",
            value=points,
            height=250,
            key="edited_points"
        )

        st.markdown("### 📥 Étape 5 — Insérer dans le Word et télécharger")
        if st.button("📄 Générer le nouveau CR Word", type="primary", use_container_width=True):
            with st.spinner("Insertion dans le Word..."):
                try:
                    doc2 = Document(io.BytesIO(st.session_state["doc_bytes"]))
                    num = st.session_state["nouveau_numero"]
                    dt = st.session_state["nouvelle_date"]
                    rdv = st.session_state["prochain_rdv"]
                    sections_map = st.session_state["sections"]
                    section_sel = st.session_state["section_choisie"]

                    # ── Mise à jour date et numéro ────────────────────────────
                    date_fr = dt.strftime("%d/%m/%Y")
                    for para in doc2.paragraphs:
                        if "CR DE REUNION DE CHANTIER" in para.text or "N°" in para.text:
                            # Remplacer le numéro
                            old_num_match = re.search(r'N°(\d+)', para.text)
                            if old_num_match and num:
                                for run in para.runs:
                                    run.text = re.sub(r'N°\d+', f'N°{num}', run.text)
                        if "PROCHAIN RDV" in para.text and rdv:
                            for run in para.runs:
                                if "MARDI" in run.text or "LUNDI" in run.text or "MERCREDI" in run.text or "JEUDI" in run.text or "VENDREDI" in run.text:
                                    run.text = re.sub(
                                        r'(LUNDI|MARDI|MERCREDI|JEUDI|VENDREDI)\s+\d+\s+\w+\s+\d{4}.*',
                                        rdv, run.text
                                    )

                    # ── Insertion dans la bonne table ─────────────────────────
                    if section_sel in sections_map:
                        table_idx = sections_map[section_sel]
                        table = doc2.tables[table_idx]

                        # Ajouter une nouvelle ligne par point
                        lignes = [l.strip() for l in edited_points.strip().split('\n') if l.strip() and not l.strip().startswith('#')]

                        for ligne in lignes:
                            # Extraire action_par et statut
                            action_par = ""
                            statut = ""
                            match_resp = re.search(r'\(([^)]+)\)\s*$', ligne)
                            if match_resp:
                                action_par = match_resp.group(1)
                                ligne_texte = ligne[:match_resp.start()].strip()
                            else:
                                ligne_texte = ligne

                            match_statut = re.search(r'\b(PA|PM|EN ATT|FAIT|EN ATTENTE)\b', ligne_texte, re.IGNORECASE)
                            if match_statut:
                                statut = match_statut.group(1).upper()

                            # Ajouter une ligne au tableau
                            new_row = table.add_row()
                            # Col 0 : texte du point
                            new_row.cells[0].text = ligne_texte
                            # Col 1 : action par
                            if len(new_row.cells) > 1:
                                new_row.cells[1].text = action_par
                            # Col 2 : noté le (date)
                            if len(new_row.cells) > 2:
                                new_row.cells[2].text = date_fr
                            # Col 3 : statut
                            if len(new_row.cells) > 3:
                                new_row.cells[3].text = statut
                            # Col 4 : fait le
                            if len(new_row.cells) > 4:
                                new_row.cells[4].text = ""

                    # ── Export ────────────────────────────────────────────────
                    output = io.BytesIO()
                    doc2.save(output)
                    output.seek(0)

                    nom_fichier = f"CR_CHANTIER_N{num}_{dt.strftime('%d-%m-%Y')}.docx"
                    st.download_button(
                        label=f"⬇️ Télécharger {nom_fichier}",
                        data=output,
                        file_name=nom_fichier,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    st.success("✅ Nouveau CR généré avec succès !")

                except Exception as e:
                    st.error(f"Erreur lors de la génération : {e}")

else:
    st.info("👆 Upload le fichier Word du CR précédent pour commencer")
