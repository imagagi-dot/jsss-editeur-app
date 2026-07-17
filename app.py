import streamlit as st
import os
import json
import importlib
import re
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from google import genai
import remplir_article
importlib.reload(remplir_article)

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_content_and_images(uploaded_file, temp_dir):
    doc = Document(uploaded_file)
    content = []
    img_count = 0
    
    NS = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            for run in block.runs:
                drawing_elements = run._element.findall('.//w:drawing', NS)
                for drawing in drawing_elements:
                    blips = drawing.findall('.//a:blip', NS)
                    for blip in blips:
                        rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        if rId:
                            image_part = doc.part.related_parts[rId]
                            img_ext = image_part.content_type.split('/')[-1]
                            if img_ext == 'jpeg': img_ext = 'jpg'
                            img_name = f"image_{img_count}.{img_ext}"
                            img_path = os.path.join(temp_dir, img_name)
                            with open(img_path, "wb") as f:
                                f.write(image_part.blob)
                            content.append(f"\n[IMAGE_PLACEHOLDER: {img_name}]\n")
                            img_count += 1
            if text:
                content.append(text)
                
        elif isinstance(block, Table):
            table_md = []
            for i, row in enumerate(block.rows):
                row_data = [cell.text.replace('\\n', ' ').strip() for cell in row.cells]
                table_md.append("| " + " | ".join(row_data) + " |")
                if i == 0:
                    table_md.append("|" + "|".join(["---"] * len(row.cells)) + "|")
            content.append("\n[TABLE_PLACEHOLDER_START]\n" + "\n".join(table_md) + "\n[TABLE_PLACEHOLDER_END]\n")
            
    return "\n".join(content)

def process_manuscript(text, api_key):
    client = genai.Client(api_key=api_key)
    
    with open("GABARIT_article.json", "r", encoding="utf-8") as f:
        gabarit = f.read()

    prompt = f"""Tu es un éditeur scientifique et médical expert, spécialisé dans la préparation de manuscrits pour le "Journal Sahélien des Sciences de la Santé (JSSS)".

Tâche : Transforme et corrige le texte brut fourni pour qu'il soit conforme, puis génère STRICTEMENT un objet JSON (pas de markdown autour).

L'objet JSON doit avoir EXACTEMENT la même structure que l'exemple suivant :
{gabarit}

Directives de traitement CRITIQUES :
1. 'header_citation' : Nom du premier auteur (ex: Harouna Amadou ML.) suivi de 'et al. J Sah Sci Santé (2026), vol 06 (1): [pages]'.
2. 'authors' & 'affiliations' : Extrais les auteurs et relie-les à leurs affiliations via des numéros en EXPOSANT encadrés par des accents circonflexes (ex: Nom Prénom^1,2^*). L'auteur correspondant prend un astérisque en plus.
3. TRADUCTION BILINGUE : Si le résumé français est présent mais pas l'anglais, tu DOIS générer 'abstract' et 'keywords_en' avec une traduction scientifique et médicale anglaise parfaite.
4. BIBLIOGRAPHIE (references) : Tu DOIS reformater CHAQUE référence bibliographique à la toute fin du document selon la NORME DE VANCOUVER stricte. Corrige les fautes, abréviations, année, etc.
5. 'body' : Le texte principal. Utilise 'type': 'h2' pour les grands titres, 'type': 'h3' pour les sous-titres, et 'type': 'p' pour les paragraphes.
6. TABLEAUX : Remplace [TABLE_PLACEHOLDER_START] par {{"type": "table", "data": [["colonne1", "colonne2"], ["valeur1", "valeur2"]]}}
7. IMAGES : Remplace [IMAGE_PLACEHOLDER: nom_du_fichier.png] par {{"type": "figure", "image": "nom_du_fichier.png"}} avec le titre de l'image (s'il y en a un juste en dessous) dans la balise "caption".

Texte brut du manuscrit à traiter avec placeholders :
{text}
"""
    
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=prompt
    )
    
    out_text = response.text.strip()
    if out_text.startswith("```json"): out_text = out_text[7:]
    if out_text.startswith("```"): out_text = out_text[3:]
    if out_text.endswith("```"): out_text = out_text[:-3]
        
    return json.loads(out_text.strip())

def compute_health_score(spec_json):
    checks = []
    
    # 1. Mots-clés
    kw_str = spec_json.get("keywords", "")
    kws = [k.strip() for k in kw_str.split(",") if k.strip()]
    if 3 <= len(kws) <= 6:
        checks.append(("✅ Mots-clés", f"{len(kws)} trouvés (Conforme)"))
    else:
        checks.append(("⚠️ Mots-clés", f"{len(kws)} trouvés (Le standard est de 3 à 6)"))
        
    # 2. Résumé Word count
    res_list = spec_json.get("resume", [])
    res_text = " ".join([item.get("t", "") for item in res_list if isinstance(item, dict)])
    words = len(res_text.split())
    if words <= 250:
        checks.append(("✅ Résumé", f"{words} mots (Conforme < 250)"))
    else:
        checks.append(("⚠️ Résumé", f"{words} mots (Trop long ! Max 250 mots)"))
        
    # 3. Structure IMRAD
    res_lower = res_text.lower()
    has_intro = "introduction" in res_lower
    has_meth = "matériel" in res_lower or "méthode" in res_lower
    has_res = "résultat" in res_lower
    has_conc = "conclusion" in res_lower
    if has_intro and has_meth and has_res and has_conc:
        checks.append(("✅ Structure IMRAD", "Présente dans le résumé"))
    else:
        checks.append(("⚠️ Structure IMRAD", "Il manque au moins une section IMRAD dans le résumé"))
        
    return checks

st.set_page_config(page_title="JSSS Auto-Éditeur (V4)", page_icon="📝", layout="wide")
st.title("📝 Éditeur Automatisé - Journal JSSS")
st.markdown("Version 4 - Assistants avancés (Traduction, Vancouver, Anonymisation, Pre-flight)")

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = {}

api_key = st.text_input("Clé API Google Gemini", type="password")
st.markdown("---")

uploaded_files = st.file_uploader("Téléversez vos manuscrits (.docx)", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("1️⃣ Analyser les documents (Étape 1)", type="primary"):
        if not api_key:
            st.error("Veuillez saisir votre clé API Google Gemini.")
        else:
            st.session_state.processed_files = {} # Reset
            for uploaded_file in uploaded_files:
                st.markdown(f"### Analyse de : `{uploaded_file.name}`")
                temp_dir = f"temp_images_{uploaded_file.name.replace('.docx', '')}"
                os.makedirs(temp_dir, exist_ok=True)
                
                with st.spinner("Extraction et IA en cours..."):
                    try:
                        raw_text = extract_content_and_images(uploaded_file, temp_dir)
                        spec_json = process_manuscript(raw_text, api_key)
                        spec_json["images_dir"] = temp_dir
                        st.session_state.processed_files[uploaded_file.name] = spec_json
                        st.success(f"Analyse terminée pour {uploaded_file.name} !")
                    except Exception as e:
                        st.error(f"Erreur avec {uploaded_file.name} : {e}")

if st.session_state.processed_files:
    st.markdown("---")
    st.header("2️⃣ Vérification & Édition (Étape 2)")
    
    updated_jsons = {}
    
    for filename, spec_json in st.session_state.processed_files.items():
        st.subheader(f"Document : {filename}")
        
        # Pre-flight checks
        st.markdown("**Contrôle Qualité (Pre-flight) :**")
        checks = compute_health_score(spec_json)
        cols = st.columns(len(checks))
        for col, check in zip(cols, checks):
            col.info(f"**{check[0]}**\n\n{check[1]}")
            
        # JSON Editor
        st.markdown("**Éditeur manuel de la structure :**")
        st.caption("Vous pouvez corriger librement la structure ci-dessous avant de générer le fichier Word final.")
        new_json_str = st.text_area(f"Structure JSON ({filename})", value=json.dumps(spec_json, ensure_ascii=False, indent=2), height=400, key=f"json_{filename}")
        try:
            updated_jsons[filename] = json.loads(new_json_str)
        except json.JSONDecodeError:
            st.error("⚠️ Le format JSON n'est plus valide. Veuillez corriger les erreurs de syntaxe (guillemets, virgules...).")
            updated_jsons[filename] = spec_json

    st.markdown("---")
    st.header("3️⃣ Génération des Fichiers (Étape 3)")
    
    generate_anon = st.checkbox("Générer une version Anonymisée (Peer-Review)", value=False)
    
    if st.button("🚀 Générer les fichiers Word", type="primary"):
        for filename, final_json in updated_jsons.items():
            try:
                out_filename = f"JSSS_Formate_{filename}"
                tmp_path = f"tmp_{out_filename}"
                remplir_article.fill(final_json, "TEMPLATE_article_JSSS.docx", tmp_path)
                with open(tmp_path, "rb") as f: docx_bytes = f.read()
                
                st.download_button(label=f"⬇️ Télécharger {out_filename}", data=docx_bytes, file_name=out_filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_{filename}")
                
                if generate_anon:
                    anon_json = json.loads(json.dumps(final_json)) # Deep copy
                    anon_json["authors"] = "VERSION ANONYMISÉE POUR RELECTURE"
                    anon_json["affiliations"] = []
                    anon_json["corresponding"] = ""
                    
                    anon_filename = f"JSSS_Anonyme_{filename}"
                    anon_tmp_path = f"tmp_{anon_filename}"
                    remplir_article.fill(anon_json, "TEMPLATE_article_JSSS.docx", anon_tmp_path)
                    with open(anon_tmp_path, "rb") as f: anon_docx_bytes = f.read()
                    
                    st.download_button(label=f"⬇️ Télécharger {anon_filename}", data=anon_docx_bytes, file_name=anon_filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_anon_{filename}")
                    
            except Exception as e:
                st.error(f"Une erreur est survenue lors de la génération de {filename} : {e}")
