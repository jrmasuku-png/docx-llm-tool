import streamlit as st
import io
import json
import zipfile

st.set_page_config(page_title="DOCX to LLM JSON Converter", layout="wide")

st.title("DOCX → LLM-Ready Comment JSON Converter")
st.write("Upload a Word document with comments. The app will extract comments, anchor text, paragraph text, and export JSON for LLM processing.")

try:
    from lxml import etree
except Exception as e:
    st.error("Missing dependency: lxml. Make sure requirements.txt contains lxml.")
    st.exception(e)
    st.stop()


W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w_attr(element, name):
    return element.get(f"{{{W_URI}}}{name}")


def local_name(tag):
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def read_docx_xml(docx_bytes, xml_path):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        if xml_path not in z.namelist():
            return None
        return etree.fromstring(z.read(xml_path))


def extract_comments(docx_bytes):
    root = read_docx_xml(docx_bytes, "word/comments.xml")

    if root is None:
        return {}

    comments = {}

    for comment in root.findall(".//w:comment", namespaces=W_NS):
        cid = w_attr(comment, "id")
        author = w_attr(comment, "author")
        date = w_attr(comment, "date")

        paragraphs = []
        for p in comment.findall(".//w:p", namespaces=W_NS):
            texts = p.findall(".//w:t", namespaces=W_NS)
            paragraph_text = "".join([t.text or "" for t in texts]).strip()
            if paragraph_text:
                paragraphs.append(paragraph_text)

        comments[cid] = {
            "comment_id": cid,
            "author": author,
            "date": date,
            "comment_text": "\n".join(paragraphs).strip()
        }

    return comments


def walk_word_xml(element, state):
    for child in element:
        name = local_name(child.tag)

        if name == "commentRangeStart":
            cid = w_attr(child, "id")
            if cid is not None:
                state["active_comments"][cid] = len(state["text"])

        elif name == "commentRangeEnd":
            cid = w_attr(child, "id")
            if cid in state["active_comments"]:
                start = state["active_comments"][cid]
                end = len(state["text"])

                state["spans"].append({
                    "comment_id": cid,
                    "paragraph_id": state["paragraph_id"],
                    "anchor_text": state["text"][start:end],
                    "char_span": {
                        "start": start,
                        "end": end
                    }
                })

                del state["active_comments"][cid]

        elif name == "commentReference":
            cid = w_attr(child, "id")
            if cid is not None:
                state["references"].append({
                    "comment_id": cid,
                    "paragraph_id": state["paragraph_id"],
                    "position": len(state["text"])
                })

        elif name == "t":
            state["text"] += child.text or ""

        elif name == "tab":
            state["text"] += "\t"

        elif name in ["br", "cr"]:
            state["text"] += "\n"

        else:
            walk_word_xml(child, state)


def extract_document_and_spans(docx_bytes):
    root = read_docx_xml(docx_bytes, "word/document.xml")

    if root is None:
        raise ValueError("Could not find word/document.xml inside the DOCX file.")

    paragraphs = []
    all_spans = []
    all_references = []

    word_paragraphs = root.findall(".//w:body//w:p", namespaces=W_NS)

    for p_idx, paragraph in enumerate(word_paragraphs):
        state = {
            "paragraph_id": p_idx,
            "text": "",
            "active_comments": {},
            "spans": [],
            "references": []
        }

        walk_word_xml(paragraph, state)

        para_text = state["text"].strip()

        if para_text or state["spans"] or state["references"]:
            paragraphs.append({
                "paragraph_id": p_idx,
                "text": para_text
            })

        all_spans.extend(state["spans"])
        all_references.extend(state["references"])

    return paragraphs, all_spans, all_references


def build_llm_json(docx_bytes, filename):
    comments = extract_comments(docx_bytes)
    paragraphs, spans, references = extract_document_and_spans(docx_bytes)

    paragraph_lookup = {
        p["paragraph_id"]: p["text"]
        for p in paragraphs
    }

    structured_comments = []

    for span in spans:
        cid = span["comment_id"]
        comment_info = comments.get(cid, {})

        structured_comments.append({
            "comment_id": cid,
            "author": comment_info.get("author"),
            "date": comment_info.get("date"),
            "comment_text": comment_info.get("comment_text", ""),
            "paragraph_id": span["paragraph_id"],
            "anchor_text": span["anchor_text"],
            "char_span": span["char_span"],
            "paragraph_text": paragraph_lookup.get(span["paragraph_id"], "")
        })

    anchored_ids = {c["comment_id"] for c in structured_comments}

    unanchored_comments = []
    for cid, comment_info in comments.items():
        if cid not in anchored_ids:
            unanchored_comments.append(comment_info)

    return {
        "metadata": {
            "source_file": filename,
            "paragraph_count": len(paragraphs),
            "comment_count": len(comments),
            "anchored_comment_count": len(structured_comments),
            "unanchored_comment_count": len(unanchored_comments)
        },
        "document": paragraphs,
        "comments": structured_comments,
        "unanchored_comments": unanchored_comments
    }


uploaded_file = st.file_uploader("Upload your commented Word document", type=["docx"])

if uploaded_file is not None:
    try:
        docx_bytes = uploaded_file.getvalue()

        with st.spinner("Extracting comments and preparing JSON..."):
            data = build_llm_json(docx_bytes, uploaded_file.name)

        st.success("Extraction complete.")

        col1, col2, col3 = st.columns(3)
        col1.metric("Paragraphs", data["metadata"]["paragraph_count"])
        col2.metric("Anchored comments", data["metadata"]["anchored_comment_count"])
        col3.metric("Unanchored comments", data["metadata"]["unanchored_comment_count"])

        if data["metadata"]["comment_count"] == 0:
            st.warning("No Word comments were found in this DOCX file. Make sure the file contains actual Word comments, not only tracked changes.")

        if data["metadata"]["unanchored_comment_count"] > 0:
            st.warning("Some comments were found but could not be linked to an exact text span. They are included under unanchored_comments.")

        st.subheader("Comment Preview")

        for comment in data["comments"][:10]:
            with st.expander(f"Comment {comment['comment_id']}"):
                st.write("**Comment text:**")
                st.write(comment["comment_text"])

                st.write("**Anchor text:**")
                st.code(comment["anchor_text"])

                st.write("**Paragraph text:**")
                st.write(comment["paragraph_text"])

        st.subheader("JSON Preview")
        st.json(data)

        json_text = json.dumps(data, indent=2, ensure_ascii=False)

        st.download_button(
            label="Download LLM-ready JSON",
            data=json_text,
            file_name="llm_ready_comments.json",
            mime="application/json"
        )

    except Exception as e:
        st.error("The app failed while processing the DOCX file.")
        st.exception(e)