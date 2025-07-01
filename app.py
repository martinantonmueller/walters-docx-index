from odf.opendocument import load
from odf.text import P, Note
import streamlit as st
import re
import requests

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

def extract_comments_odt(odt_file):
    doc = load(odt_file)

    # Alle Kommentare (Notes) im Dokument finden
    notes = doc.getElementsByType(Note)

    output_lines = []
    for note in notes:
        # Text im Kommentar extrahieren
        texts = []
        for node in note.childNodes:
            if node.nodeType == node.TEXT_NODE:
                texts.append(node.data)
            else:
                # Textknoten rekursiv extrahieren
                texts.append(node.firstChild.data if node.firstChild else '')
        comment_text = ''.join(texts).strip()

        # ID extrahieren und ggf. Personendaten holen
        pmb_id = extract_id(comment_text)
        extra = ""
        if pmb_id:
            person = get_person_data(pmb_id)
            if person:
                extra = f" → {person}"

        output_lines.append(f"Kommentar: {comment_text}{extra}")

    return output_lines

# --- Streamlit UI ---

st.title("ODT-Kommentare + PMB-Link-Parser")
st.write("Lade eine `.odt`-Datei hoch, um Kommentare zu extrahieren.")

uploaded = st.file_uploader("ODT-Datei wählen", type=["odt"])

if uploaded:
    result = extract_comments_odt(uploaded)
    if result:
        text_output = "\n".join(result)
        st.download_button(
            "Ergebnis herunterladen",
            text_output.encode("utf-8"),
            file_name=uploaded.name.replace(".odt", "_index.txt"),
            mime="text/plain"
        )
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
