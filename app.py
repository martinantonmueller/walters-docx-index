import streamlit as st
from docx import Document
import re
import requests
from lxml import etree

NAMESPACES = {
    'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'xml': 'http://www.w3.org/XML/1998/namespace'
}

@st.cache_data(show_spinner=False)
def get_person_data(pid):
    url = f"https://pmb.acdh.oeaw.ac.at/apis/api/entities/person/{pid}"
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            return f"{data.get('name')} ({data.get('first_name', '')}, {data.get('start_date_written', '')}–{data.get('end_date_written', '')})"
    except:
        pass
    return None

def extract_id(comment_text):
    id_match = re.search(r'/person/(\d+)', comment_text)
    if id_match:
        return id_match.group(1)
    num_match = re.search(r'\b(\d{3,})\b', comment_text)
    if num_match:
        return num_match.group(1)
    return None

def extract_comments_extensible(docx_file):
    doc = Document(docx_file)
    comments_part = None

    # Suche Part mit Kommentaren (egal ob klassisch oder extensible)
    for rel in doc.part.rels.values():
        if "comments" in rel.reltype:
            comments_part = rel.target_part
            break

    if not comments_part:
        return []

    # Lade XML
    comments_xml = etree.fromstring(comments_part.blob)

    # Prüfe Root-Tag, ob extensible Kommentare vorliegen
    if comments_xml.tag.endswith('commentsExtensible'):
        # Kommentare im neuen extensible Format
        comments = comments_xml.findall(".//w16cex:commentExtensible", namespaces=NAMESPACES)
        comment_map = {}
        for comment in comments:
            cid = comment.get("{http://schemas.microsoft.com/office/word/2018/wordml/cex}durableId")
            # Der Text ist oft in einem "w:t" Tag verschachtelt, wir holen alles an Text darunter
            text = ''.join(comment.xpath(".//w:t//text()", namespaces=NAMESPACES)).strip()
            comment_map[cid] = text
        return comment_map

    elif comments_xml.tag.endswith('comments'):
        # Klassisches Kommentarformat (Fallback)
        comments = comments_xml.findall(".//w:comment", namespaces=NAMESPACES)
        comment_map = {}
        for comment in comments:
            cid = comment.get('{http://www.w3.org/XML/1998/namespace}id')
            text = ''.join(comment.itertext()).strip()
            comment_map[cid] = text
        return comment_map

    else:
        # Unbekanntes Format
        return {}


# --- STREAMLIT UI ---

st.title("DOCX-Kommentare + PMB-Link-Parser")
st.write("Lade eine `.docx`-Datei hoch, um Kommentare zu extrahieren.")

uploaded = st.file_uploader("DOCX-Datei wählen", type=["docx"])

if uploaded:
    result = extract_comments(uploaded)
    if result:
        text_output = "\n".join(result)
        st.download_button(
            "Ergebnis herunterladen",
            text_output.encode("utf-8"),
            file_name=uploaded.name.replace(".docx", "_index.txt"),
            mime="text/plain"
        )
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
