import streamlit as st
import zipfile
from lxml import etree

def extract_comments_from_odt_bytesio(uploaded_file):
    ns = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0'
    }

    comments = []

    try:
        with zipfile.ZipFile(uploaded_file) as z:
            with z.open('content.xml') as f:
                parser = etree.XMLParser(recover=True)  # Wichtig: recover=True erlaubt "kaputtes" XML
                tree = etree.parse(f, parser)
    except Exception as e:
        st.error(f"Fehler beim Lesen der ODT-Datei: {e}")
        return []

    annotations = tree.xpath('//office:annotation', namespaces=ns)
    if not annotations:
        return []

    for ann in annotations:
        paragraphs = ann.xpath('./text:p', namespaces=ns)
        comment_text = "\n".join(''.join(p.itertext()) for p in paragraphs).strip()
        author_el = ann.find('dc:creator', namespaces=ns)
        author = author_el.text if author_el is not None else "Unbekannt"
        comments.append(f"{author}: {comment_text}")

    return comments

st.title("Kommentare aus ODT-Dateien extrahieren")

uploaded = st.file_uploader("Bitte ODT-Datei auswählen", type=["odt"])

if uploaded:
    comments = extract_comments_from_odt_bytesio(uploaded)
    if comments:
        st.write(f"{len(comments)} Kommentar(e) gefunden:")
        for c in comments:
            st.markdown(f"- {c}")
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
