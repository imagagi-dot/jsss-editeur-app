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

def extract_drawings_from_runs(runs, doc, temp_dir, NS, img_count):
    extracted_placeholders = []
    for run in runs:
        drawing_elements = run._element.findall('.//w:drawing', NS)
        for drawing in drawing_elements:
            blips = drawing.findall('.//a:blip', NS)
            for blip in blips:
                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if rId and rId in doc.part.related_parts:
                    image_part = doc.part.related_parts[rId]
                    img_ext = image_part.content_type.split('/')[-1]
                    if img_ext == 'jpeg': img_ext = 'jpg'
                    img_name = f"image_{img_count[0]}.{img_ext}"
                    img_path = os.path.join(temp_dir, img_name)
                    with open(img_path, "wb") as f:
                        f.write(image_part.blob)
                    extracted_placeholders.append(f"\n[IMAGE_PLACEHOLDER: {img_name}]\n")
                    img_count[0] += 1
    return extracted_placeholders

def extract_content_and_images(uploaded_file, temp_dir):
    doc = Document(uploaded_file)
    content = []
    img_count = [0]
    
    NS = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            imgs = extract_drawings_from_runs(block.runs, doc, temp_dir, NS, img_count)
            content.extend(imgs)
            text = block.text.strip()
            if text:
                content.append(text)
                
        elif isinstance(block, Table):
            table_md = []
            has_table_text = False
            for i, row in enumerate(block.rows):
                row_data = []
                for cell in row.cells:
                    for p in cell.paragraphs:
                        imgs = extract_drawings_from_runs(p.runs, doc, temp_dir, NS, img_count)
                        content.extend(imgs)
                    c_text = cell.text.replace('\n', ' ').strip()
                    if c_text:
                        has_table_text = True
                    row_data.append(c_text)
                table_md.append("| " + " | ".join(row_data) + " |")
                if i == 0:
                    table_md.append("|" + "|".join(["---"] * len(row.cells)) + "|")
            if has_table_text:
                content.append("\n[TABLE_PLACEHOLDER_START]\n" + "\n".join(table_md) + "\n[TABLE_PLACEHOLDER_END]\n")
            
    return "\n".join(content)

def process_manuscript(text, api_key):
    client = genai.Client(api_key=api_key)
    
    with open("GABARIT_article.json", "r", encoding="utf-8") as f:
        gabarit = f.read()

    prompt = f"""Tu es un éditeur scientifique et médical expert, spécialisé dans la préparation et la structuration de manuscrits pour le "Journal Sahélien des Sciences de la Santé (JSSS)".

RÈGLE D'OR ABSOLUE : Tu es un FORMATEUR DE TEXTE, PAS UN RÉDACTEUR NI UN RÉSUMEUR.
LE TEXTE ORIGINAL SOUMIS DOIT SORTIR À 100% INTÉGRAL, SANS AUCUNE COUPE, SANS AUCUN RÉSUMÉ, SANS REFORMULATION.

Tâche : Structure l'intégralité du texte brut fourni en un objet JSON valide (sans markdown ```json autour).

Structure JSON type requise :
{gabarit}

Directives STRICTES ET NON NÉGOCIABLES :
1. 'header_citation' : Nom du premier auteur (ex: Nom Initiale.) suivi de 'et al. J Sah Sci Santé (2026), vol 06 (1): [pages]'.
2. 'authors' & 'affiliations' :
   - Extrais fidèlement toutes les affiliations (1, 2, 3, 4, 5, etc.) dans 'affiliations' sans décalage.
   - Relie chaque auteur à ses numéros d'affiliation exacts avec des exposants : ex: "Nom P^1,2^*" (astérisque pour l'auteur correspondant).
3. RÉSUMÉ & ABSTRACT ('resume' et 'abstract') :
   - DANS 'resume' : Découpe le résumé français en 4 sections : "Introduction", "Matériels et méthodes", "Résultats", "Conclusion". TU DOIS RECOPIER L'INTÉGRALITÉ ABSOLUE DU TEXTE DE CHAQUE SECTION SANS ENLEVER UNE SEULE PHRASE. IL EST STRICTEMENT INTERDIT DE RÉSUMER L'INTRODUCTION OU LES RÉSULTATS DU RÉSUMÉ.
   - DANS 'abstract' : Si le résumé en anglais (Summary ou Abstract) est présent dans le texte original, RECOPIE-LE INTÉGRALEMENT mot pour mot avec ses 4 sections. Ne le résume jamais. S'il n'existe pas dans le texte original, traduis fidèlement le résumé français complet en anglais médical.
4. CORPS DU MANUSCRIT ('body') :
   - COPIE L'INTÉGRALITÉ ABSOLUE DU MANUSCRIT. CHAQUE PARAGRAPHE, CHAQUE PHRASE, CHAQUE TERME DOIT ÊTRE PRÉSENT.
   - Titres principaux (Introduction, Matériel et Méthodes, Résultats, Discussion, Conclusion) -> {{"type": "h2", "text": "..."}}
   - Sous-titres -> {{"type": "h3", "text": "..."}}
   - Paragraphes de texte -> {{"type": "p", "text": "..."}}
5. FIGURES & SOUS-LÉGENDES DÉTAILLÉES :
   - Chaque [IMAGE_PLACEHOLDER: nom_du_fichier.png] DOIT être converti en {{"type": "figure", "image": "nom_du_fichier.png", "caption": "..."}}.
   - ATTENTION AUX LÉGENDES ET SOUS-LÉGENDES : Si une figure comporte des explications détaillées ou des sous-descriptions (ex: "Figure 2: Pièces chirurgicales", "A-Pancréatectomie caudale : ...", "B-Spléno-pancréatectomie : ..."), TU DOIS IMPÉRATIVEMENT TOUT CONSERVER en plaçant la légende complète dans "caption" ou en insérant des paragraphes {{"type": "p", "text": "..."}} immédiatement sous la figure. IL EST STRICTEMENT INTERDIT D'OUBLIER UNE LÉGENDE.
6. TABLEAUX :
   - Transforme chaque bloc [TABLE_PLACEHOLDER_START]...[TABLE_PLACEHOLDER_END] en {{"type": "table", "data": [[...], [...]]}}.
   - Si le tableau possède un titre au-dessus (ex: "Tableau N° 1 : ..."), place-le dans un bloc {{"type": "caption", "text": "Tableau N° 1 : ..."}} juste AVANT l'objet table.
   - Conserve toutes les lignes et colonnes du tableau à 100%.
7. BIBLIOGRAPHIE ('references') :
   - Reformate chaque référence selon la norme de Vancouver stricte (Auteurs. Titre. Revue abrégée. Année;volume(numéro):pages). Ne supprime aucune référence.
8. 'corresponding' & 'conflict' :
   - Extrais les coordonnées complètes de l'auteur correspondant.

Texte brut du manuscrit soumis :
{text}
"""
    
    try:
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=prompt,
            config=genai.types.GenerateContentConfig(
                temperature=0.0
            )
        )
    except Exception as e:
        try:
            # En cas d'échec, on liste les modèles disponibles pour diagnostiquer
            models = [m.name for m in client.models.list()]
            raise Exception(f"Le modèle demandé n'est pas accessible. Modèles disponibles pour cette clé : {', '.join(models)}. Erreur d'origine: {str(e)}")
        except Exception:
            raise e
    
    out_text = response.text.strip()
    if out_text.startswith("```json"): out_text = out_text[7:]
    if out_text.startswith("```"): out_text = out_text[3:]
    if out_text.endswith("```"): out_text = out_text[:-3]
        
    parsed_json = json.loads(out_text.strip())
    
    # --- SECURITE ANTI-OUBLI ---
    # Récupération forcée des images
    original_images = re.findall(r'\[IMAGE_PLACEHOLDER:\s*(.*?)\]', text)
    json_body = parsed_json.get("body", [])
    
    # Nettoyer si l'IA a laissé le placeholder en texte au lieu d'en faire un objet
    cleaned_body = []
    for block in json_body:
        if block.get("type") == "p" and "text" in block and "[IMAGE_PLACEHOLDER:" in block["text"]:
            imgs = re.findall(r'\[IMAGE_PLACEHOLDER:\s*(.*?)\]', block["text"])
            for img in imgs:
                cleaned_body.append({"type": "figure", "image": img.strip(), "caption": "⚠️ FIGURE MAL FORMATÉE PAR L'IA"})
            block["text"] = re.sub(r'\[IMAGE_PLACEHOLDER:\s*.*?\]', '', block["text"]).strip()
            if block["text"]: cleaned_body.append(block)
        else:
            cleaned_body.append(block)
    
    json_body = cleaned_body
    
    # Trouver les images déjà présentes dans le JSON
    json_images = [b.get("image") for b in json_body if b.get("type") == "figure" and "image" in b]
    for img in original_images:
        if img.strip() not in json_images:
            json_body.append({
                "type": "figure",
                "image": img.strip(),
                "caption": "⚠️ FIGURE OUBLIÉE RÉCUPÉRÉE AUTOMATIQUEMENT"
            })
            
    # Récupération forcée des tableaux
    original_tables = re.findall(r'\[TABLE_PLACEHOLDER_START\](.*?)\[TABLE_PLACEHOLDER_END\]', text, re.DOTALL)
    json_tables_count = sum(1 for b in json_body if b.get("type") == "table")
    
    if json_tables_count < len(original_tables):
        for i in range(json_tables_count, len(original_tables)):
            md_table = original_tables[i].strip()
            data = []
            for line in md_table.split('\n'):
                line = line.strip()
                if line.startswith('|') and '---' not in line:
                    cols = [c.strip() for c in line.strip('|').split('|')]
                    if any(cols): data.append(cols)
            if data:
                json_body.append({"type": "table", "data": data})
                json_body.append({"type": "p", "text": "⚠️ TABLEAU OUBLIÉ RÉCUPÉRÉ AUTOMATIQUEMENT (Titre manquant)"})
                
    parsed_json["body"] = json_body
    return parsed_json

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

st.set_page_config(page_title="JSSS Auto-Éditeur (V4)", page_icon="🔬", layout="wide")

col_head1, col_head2 = st.columns([0.8, 0.2])
with col_head1:
    st.title("🔬 Éditeur Scientifique - Journal JSSS")
    st.markdown("Version 4 - Assistants avancés (Traduction, Vancouver, Anonymisation, Pre-flight)")
with col_head2:
    st.write("") # Espacement pour aligner
    if st.button("🔄 Rafraîchir l'application", type="secondary", use_container_width=True):
        st.session_state.clear()
        st.rerun()

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
    
    st.markdown("**Paramètres de publication :**")
    col_pub1, col_pub2, col_pub3, col_pub4 = st.columns(4)
    with col_pub1:
        start_page_num = st.number_input("1ère page", min_value=1, value=1, step=1)
    with col_pub2:
        year_num = st.number_input("Année", min_value=2020, value=2026, step=1)
    with col_pub3:
        vol_num = st.number_input("Volume", min_value=1, value=6, step=1)
    with col_pub4:
        issue_num = st.number_input("Numéro", min_value=1, value=1, step=1)
        
    st.markdown("**Options :**")
    generate_anon = st.checkbox("Générer une version Anonymisée (Peer-Review)", value=False)
    
    if st.button("🚀 Générer les fichiers Word", type="primary"):
        import re
        for filename, final_json in updated_jsons.items():
            try:
                final_json["start_page"] = start_page_num
                
                # Mise à jour dynamique de l'année, du volume et du numéro
                if "header_citation" in final_json:
                    cit = final_json["header_citation"]
                    cit = re.sub(r'\(\d{4}\)', f'({year_num})', cit)
                    cit = re.sub(r'vol\s+\d+\s*\(\d+\)', f'vol {vol_num:02d} ({issue_num})', cit, flags=re.IGNORECASE)
                    final_json["header_citation"] = cit
                
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
