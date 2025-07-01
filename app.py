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

def extract_comments(docx_file):
    doc = Document(docx_file)
    comments_part = None

    for rel in doc.part.rels.values():
        if "comments" in rel.reltype:
            comments_part = rel.target_part
            break

    if not comments_part:
        st.warning("Keine Kommentar-Parts im Dokument gefunden!")
        return []

    comments_xml_str = comments_part.blob.decode('utf-8')
    st.text_area("Kommentare XML komplett", comments_xml_str, height=300)

    comments_xml = etree.fromstring(comments_part.blob)

    # Zunächst klassische Kommentare suchen:
    comments = comments_xml.findall(".//w16cex:commentExtensible", namespaces={
        'w16cex': 'http://schemas.microsoft.com/office/word/2018/wordml/cex'
    })
    if len(comments) == 0:
        # Wenn keine klassischen Kommentare, dann in extensible Comments suchen:
        comments = comments_xml.findall(".//w16cex:commentExtensible", namespaces=NAMESPACES)

        st.write(f"Gefundene Kommentare: {len(comments)}")

    if len(comments) == 0:
        st.warning("Keine Kommentare im XML gefunden!")
        return []

    comment_map = {}
    for comment in comments:
        cid = comment.get('{http://www.w3.org/XML/1998/namespace}id')
        if cid:
            # Raw XML als Debug
            raw_xml = etree.tostring(comment, pretty_print=True, encoding='unicode')
            st.write(f"Kommentar ID: {cid}")
            st.write(raw_xml)
            # Versuche Text zusammenzuziehen
            text = ''.join(comment.itertext()).strip()
            st.write(f"Extrahierter Text: '{text}'")
            comment_map[cid] = text

    st.write(f"Gefundene Kommentare im XML: {len(comments)}")
    st.write("Kommentar-Map:", comment_map)

    output_lines = []
    for para in doc.paragraphs:
        for run in para.runs:
            lxml_elem = etree.fromstring(run._element.xml.encode('utf-8'))
            comment_refs = lxml_elem.xpath(".//w:commentRangeStart", namespaces=NAMESPACES)
            if comment_refs:
                cid = comment_refs[0].get("{http://www.w3.org/XML/1998/namespace}id")
                word = run.text or ""
                comment = comment_map.get(cid, "")
                pmb_id = extract_id(comment)
                extra = ""
                if pmb_id:
                    person = get_person_data(pmb_id)
                    if person:
                        extra = f" → {person}"
                line = f"{word}] {comment}{extra}"
                st.write("Gefundener Kommentar-Text:", line)
                output_lines.append(line)

    st.write(f"Output lines count: {len(output_lines)}")
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
