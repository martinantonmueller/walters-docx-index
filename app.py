import streamlit as st
import zipfile
from lxml import etree
import requests

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
                parser = etree.XMLParser(recover=True)
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
        comments.append((author, comment_text))

    return comments

def fetch_person_data(person_id):
    url = f"https://pmb.acdh.oeaw.ac.at/apis/api/entities/person/{person_id}/detail"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            # Hier je nach API-Struktur anpassen:
            name = data.get('name', 'Name nicht gefunden')
            first_name = data.get('first_name', 'Vorname nicht gefunden')
            return name, first_name
        else:
            return None, None
    except Exception as e:
        st.error(f"API-Abfrage fehlgeschlagen: {e}")
        return None, None

st.title("Kommentare aus ODT-Dateien mit Personen-Infos")

uploaded = st.file_uploader("Bitte ODT-Datei auswählen", type=["odt"])

if uploaded:
    comments = extract_comments_from_odt_bytesio(uploaded)
    if comments:
        st.write(f"{len(comments)} Kommentar(e) gefunden:")
        for author, comment in comments:
            st.write(f"**{author}** schrieb:")
            st.write(f"> {comment}")
            # Prüfen, ob Kommentar nur eine Ziffer ist:
            if comment.isdigit():
                name, first_name = fetch_person_data(comment)
                if name and first_name:
                    st.success(f"Person gefunden: {first_name} {name}")
                else:
                    st.warning("Personendaten konnten nicht geladen werden.")
            else:
                st.info("Kommentar enthält keine reine Zahl.")
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
