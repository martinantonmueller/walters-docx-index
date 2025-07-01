import streamlit as st
import zipfile
import re
from lxml import etree
import requests

def extract_comments_with_context_from_odt_bytesio(uploaded_file):
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

    # Suche alle <text:span>, die eine Annotation haben
    spans_with_annotations = tree.xpath('//text:span[office:annotation]', namespaces=ns)

    for span in spans_with_annotations:
        annotation = span.find('office:annotation', namespaces=ns)
        if annotation is None:
            continue
        
        # Kommentartext aus Annotation extrahieren
        paragraphs = annotation.findall('text:p', namespaces=ns)
        comment_text = "\n".join(''.join(p.itertext()) for p in paragraphs).strip()
        
        # Kontext-Text: alles im span außerhalb der Annotation-Elemente (also z.B. nach </office:annotation>)
        # Wir nehmen den Text des span nach der Annotation (tail) + den restlichen Text im span
        # Die Annotation-Elemente haben i.d.R. keinen eigenen Text, deshalb:
        
        # Der Kontexttext ist der Text im span nach dem Annotation-Element:
        # Etwa span.text, oder span.tail – man muss vorsichtig mit XML-Text und Tail umgehen.
        # Einfach alle Texte in span außer Annotation zusammenfassen:
        
        context_parts = []
        # Text vor erster Annotation (falls vorhanden)
        if span.text:
            context_parts.append(span.text)
        # Nach Annotationen jeweils das .tail
        for elem in span:
            if elem.tail:
                context_parts.append(elem.tail)
        
        context_text = ''.join(context_parts).strip()
        
        # Autor aus Annotation
        author_el = annotation.find('dc:creator', namespaces=ns)
        author = author_el.text if author_el is not None else "Unbekannt"

        comments.append({
            'author': author,
            'comment': comment_text,
            'context_text': context_text
        })

    return comments

def fetch_person_data(person_id):
    api_url = f"https://pmb.acdh.oeaw.ac.at/apis/api/entities/person/{person_id}/"
    try:
        response = requests.get(api_url, timeout=5)
        st.write(f"API URL: {api_url}")
        if response.status_code == 200:
            data = response.json()
            name = data.get('name') or ""
            first_name = data.get('first_name') or ""
            start_date_written = data.get('start_date_written') or ""
            end_date_written = data.get('end_date_written') or ""
            return name, first_name, start_date_written, end_date_written
        else:
            st.error(f"Fehler bei API-Abfrage: Status {response.status_code}")
            return None, None, None, None
    except Exception as e:
        st.error(f"API-Abfrage fehlgeschlagen: {e}")
        return None, None, None, None

st.title("Kommentare aus ODT-Dateien mit Personen-Infos")

uploaded = st.file_uploader("Bitte ODT-Datei auswählen", type=["odt"])

if uploaded:
    comments = extract_comments_with_context_from_odt_bytesio(uploaded)
    if comments:
        for i, c in enumerate(comments, start=1):
            comment_text = c['comment']
            context_text = c['context_text']

            match = re.search(r'(\d+)', comment_text)
            if match:
                person_id = match.group(1)
                name, first_name, start_date, end_date = fetch_person_data(person_id)

                if name or first_name:
                    # Lebensdaten nur anzeigen, wenn vorhanden
                    life_dates = ""
                    if start_date and end_date:
                        life_dates = f" ({start_date} – {end_date})"
                    elif start_date:
                        life_dates = f" (geb. {start_date})"
                    elif end_date:
                        life_dates = f" (gest. {end_date})"

                    person_info = f"{first_name} {name}".strip() + life_dates
                else:
                    person_info = "Personendaten konnten nicht geladen werden."

                display_url = f"https://pmb.acdh.oeaw.ac.at/apis/entities/entity/person/{person_id}/detail"
            else:
                person_info = "Keine Zahl oder gültige ID im Kommentar gefunden."
                display_url = "Keine gültige URL"

            st.markdown(f"""
            <div style="background-color:#d4edda; border:1px solid #c3e6cb; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
                <b>Kommentar #{i}</b><br>
                <b>Kontext:</b> {context_text}<br>
                <b>URL / Kommentar:</b> <a href="{display_url}" target="_blank">{display_url}</a><br>
                <b>Name:</b> {person_info}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")


