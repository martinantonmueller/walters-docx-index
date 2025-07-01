import streamlit as st
from docx import Document
import re
import requests
from io import BytesIO
from lxml import etree

def extract_id(comment_text):
    id_match = re.search(r'/person/(\d+)', comment_text)
    if id_match:
        return id_match.group(1)
    num_match = re.search(r'\b(\d{3,})\b', comment_text)
    if num_match:
        return num_match.group(1)
    return None

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

def extract_comments(docx_file):
    doc = Document(docx_file)
    rels = doc.part.rels
    comments_part = None

    for rel in rels:
        if "comments" in rels[rel].reltype:
            comments_part = rels[rel]._target
            break

    if not comments_part:
        return []

    
    comments_xml = etree.fromstring(comments_part.blob)
    comments = comments_xml.findall(".//w:comment", namespaces=doc.part.package.xmlns)

    comment_map = {}
f   or comment in comments:
        cid = comment.get("{http://www.w3.org/XML/1998/namespace}id")
        text = ''.join(comment.itertext()).strip()
        comment_map[cid] = text

    output_lines = []
    for para in doc.paragraphs:
        for run in para.runs:
            comment_ref = run._element.xpath(".//w:commentRangeStart")
            if comment_ref:
                cid = comment_ref[0].get("{http://www.w3.org/XML/1998/namespace}id")
                word = run.text
                comment = comment_map.get(cid, "")
                pmb_id = extract_id(comment)
                extra = ""
                if pmb_id:
                    person = get_person_data(pmb_id)
                    if person:
                        extra = f" → {person}"
                output_lines.append(f"{word}] {comment}{extra}")
    return output_lines

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
