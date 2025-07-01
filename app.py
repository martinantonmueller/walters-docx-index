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

    # Alle Annotations finden
    annotations = tree.xpath('//office:annotation', namespaces=ns)
    if not annotations:
        return []

    # Alle <span> Elemente mit office:annotation-Attribut finden (dein XPath)
    annotated_spans = tree.xpath('//*[local-name()="span" and @office:annotation]', namespaces=ns)

    # Wir gehen davon aus, dass annotations und annotated_spans in der gleichen Reihenfolge passen
    for ann, span in zip(annotations, annotated_spans):
        # Kommentartext zusammenbauen
        paragraphs = ann.xpath('./text:p', namespaces=ns)
        comment_text = "\n".join(''.join(p.itertext()) for p in paragraphs).strip()

        # Autor des Kommentars
        author_el = ann.find('dc:creator', namespaces=ns)
        author = author_el.text if author_el is not None else "Unbekannt"

        # Text, an dem der Kommentar hängt (annotated span Text)
        annotated_text = ''.join(span.itertext()).strip()

        comments.append((author, comment_text, annotated_text))

    return comments

def fetch_person_data(person_id):
    url = f"https://pmb.acdh.oeaw.ac.at/apis/api/entities/person/{person_id}/"
    try:
        response = requests.get(url, timeout=5)
        st.write(f"API URL: {url}")
        st.write(f"Status Code: {response.status_code}")
        if response.status_code == 200:
            data = response.json()
            name = data.get('name')
            first_name = data.get('first_name')
            return name, first_name
        else:
            st.error(f"Fehler bei API-Abfrage: Status {response.status_code}")
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
        for author, comment, annotated_text in comments:
            st.write(f"**{author}** schrieb:")
            st.write(f"> {comment}")
            st.write(f"*Kommentierter Text:* {annotated_text}")

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
