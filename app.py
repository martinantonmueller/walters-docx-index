import streamlit as st
from odf.opendocument import load
from odf.namespaces import OFFICE, TEXT
from odf.element import Element

def extract_odt_comments(odt_file):
    doc = load(odt_file)
    annotations = doc.getElementsByType(Element)
    comments = []
    for elem in annotations:
        if elem.qname[1] == 'annotation' and elem.qname[0] == OFFICE:
            comment_texts = []
            for child in elem.childNodes:
                if child.qname[1] == 'p' and child.qname[0] == TEXT:
                    text_content = ''.join(node.data for node in child.childNodes if node.nodeType == node.TEXT_NODE)
                    comment_texts.append(text_content)
            comment = '\n'.join(comment_texts).strip()
            if comment:
                comments.append(comment)
    return comments

st.title("Kommentare aus ODT extrahieren")

uploaded = st.file_uploader("ODT-Datei auswählen", type=["odt"])

if uploaded:
    comments = extract_odt_comments(uploaded)
    if comments:
        st.write(f"{len(comments)} Kommentare gefunden:")
        for i, c in enumerate(comments, 1):
            st.markdown(f"**Kommentar {i}:** {c}")
    else:
        st.info("Keine Kommentare gefunden oder Dokument nicht kompatibel.")
