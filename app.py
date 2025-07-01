import streamlit as st
import zipfile
from lxml import etree
import requests

def extract_comments_with_context_from_odt_bytesio(uploaded_file):
    ns = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0'
    }

    results = []

    try:
        with zipfile.ZipFile(uploaded_file) as z:
            with z.open('content.xml') as f:
                parser = etree.XMLParser(recover=True)
                tree = etree.parse(f, parser)
    except Exception as e:
        st.error(f"Fehler beim Lesen der ODT-Datei: {e}")
        return []

    spans_with_annotation = tree.xpath('//text:span[office:annotation]', namespaces=ns)
    if not spans_with_annotation:
        return []

    for span in spans_with_annotation:
        annotation = span.find('office:annotation', namespaces=ns)
        paragraphs = annotation.xpath('./text:p', namespaces=ns)
        comment_text = "\n".join(''.join(p.itertext()) for p in paragraphs).strip()
        author_el = annotation.find('dc:creator', namespaces=ns)
        author = author_el.text if author_el is not None else "Unbekannt"
        
        # Text außerhalb der Annotation im selben <text:span>
        text_outside_annotation = ""
        for node in span.iterchildren():
            if node.tag == '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}annotation':
                continue
            text_outside_annotation += ''.join(node.itertext())
        if annotation.tail:
            text_outside_annotation += annotation.tail.strip()

        results.append({
            'author': author,
            'comment': comment_text,
            'context_text': text_outside_annotation.strip()
        })

    return results

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


import re

if uploaded:
    comments = extract_comments_with_context_from_odt_bytesio(uploaded)
    if comments:
        for c in comments:
            st.write(f"> Kommentar: {c['comment']}")
            st.write(f"> Kontext: {c['context_text']}")

            # Nummer aus Kommentar extrahieren (Zahl oder aus URL):
            match = re.search(r'(\d+)', c['comment'])
            if match:
                person_id = match.group(1)
                name, first_name = fetch_person_data(person_id)
                if name and first_name:
                    st.success(f"Person gefunden: {first_name} {name}")
                else:
                    st.warning("Personendaten konnten nicht geladen werden.")
            else:
                st.info("Kommentar enthält keine Zahl oder gültige ID.")
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
