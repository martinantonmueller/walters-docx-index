import zipfile
import xml.etree.ElementTree as ET
import requests
import streamlit as st

NS = {
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
}

def extract_comments_with_context(odt_file):
    comments = []
    try:
        with zipfile.ZipFile(odt_file) as odt_zip:
            content_xml = odt_zip.read('content.xml')
            root = ET.fromstring(content_xml)
            
            # Alle office:annotation-Elemente suchen
            for annotation in root.findall('.//office:annotation', NS):
                # Kommentarinformationen
                creator = annotation.find('dc:creator', NS)
                creator_name = creator.text if creator is not None else "Unbekannt"
                date = annotation.find('dc:date', NS)
                date_str = date.text if date is not None else "Kein Datum"
                
                # Kommentartext (kann in mehreren text:p sein, hier vereinfachend erstes p)
                p = annotation.find('text:p', NS)
                comment_text = p.text if p is not None else ""
                
                # Der Text, in dem der Kommentar steht, ist das Elternelement von annotation
                # Wir gehen ein Level höher zum <text:span> (oder <text:p>), dann lesen wir den Text dieses Elements.
                parent = annotation.getparent() if hasattr(annotation, 'getparent') else None
                # Da xml.etree.ElementTree hat keine getparent(), wir brauchen workaround:
                # Wir iterieren im content.xml-Text und merken uns den Text in Nähe der Annotation
                
                # Einfacher Workaround: Im ODT sind office:annotation Elemente immer in einem text:span,
                # Wir suchen den nächsten Textknoten nach office:annotation, der der kommentierte Text ist.
                
                # Da ET keine getparent hat, hier die Methode über XPath nach text:span, das annotation enthält:
                # Wir können eine Map bauen oder einfach den text unmittelbar nach annotation ausgeben.
                
                # Alternativ, wir extrahieren den Tail-Text des annotation-Elements:
                tail_text = annotation.tail or ""
                
                comments.append({
                    'creator': creator_name,
                    'date': date_str,
                    'comment_text': comment_text.strip(),
                    'context_text': tail_text.strip()
                })
                
        if not comments:
            return None
        return comments
    except Exception as e:
        st.error(f"Fehler beim Lesen der ODT-Datei: {e}")
        return None

def fetch_person_data(person_id):
    url = f"https://pmb.acdh.oeaw.ac.at/apis/api/entities/person/{person_id}/"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            # Beispielhafte Felder, ggf. anpassen
            name = data.get('name')
            first_name = data.get('first_name')
            return name, first_name
        else:
            return None, None
    except Exception:
        return None, None


# Beispiel Integration in Streamlit

import streamlit as st

uploaded = st.file_uploader("ODT-Datei wählen", type=["odt"])

if uploaded:
    comments = extract_comments_with_context(uploaded)
    if comments:
        for c in comments:
            st.write(f"Kommentar von: {c['creator']} am {c['date']}")
            st.write(f"Kommentartext: {c['comment_text']}")
            st.write(f"Kontext/Text zur Markierung: {c['context_text']}")
            
            # Prüfe ob Kommentar nur eine Zahl enthält
            if c['comment_text'].isdigit():
                name, first_name = fetch_person_data(c['comment_text'])
                if name or first_name:
                    st.write(f"Personendaten: Name = {name}, Vorname = {first_name}")
                else:
                    st.write("Personendaten konnten nicht geladen werden.")
            st.write("---")
    else:
        st.write("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
