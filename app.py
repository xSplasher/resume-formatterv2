import streamlit as st
from formatter import extract_text_from_bytes, call_llm, generate_docx_bytes, generate_html_preview, docx_bytes_to_pdf_bytes, count_pdf_pages

st.set_page_config(page_title="Resume Formatter", page_icon="📄", layout="wide")

API_KEY = st.secrets.get("OPENROUTER_KEY", "")

st.markdown("## Resume Formatter")
st.markdown("Upload a resume and get a professionally formatted version in seconds.")

uploaded = st.file_uploader("Upload a resume", type=["pdf", "docx", "doc"])
compact = st.checkbox("Compact mode", help="Shorten and tighten the resume to fit on one page. Content may be condensed.")

if uploaded is not None:
    file_bytes = uploaded.read()
    filename = uploaded.name

    is_processing = st.session_state.get("processing", False)
    if st.button("Format Resume", type="primary", disabled=is_processing):
        st.session_state.processing = True
        with st.spinner("Extracting text..."):
            try:
                text = extract_text_from_bytes(file_bytes, filename)
            except Exception as e:
                st.error(f"Failed to extract text: {e}")
                st.stop()

        if not text.strip():
            st.error("Could not extract any text from this file.")
            st.stop()

        st.info(f"Extracted {len(text)} characters. Sending to AI for processing...")

        with st.spinner("AI is analyzing and structuring the resume... this may take up to 60 seconds."):
            try:
                data = call_llm(text, API_KEY, compact=compact)
            except Exception as e:
                st.error(f"AI processing failed: {e}")
                st.stop()

        st.session_state.result_data = data
        st.session_state.result_filename = filename

        docx_bytes = generate_docx_bytes(data)
        st.session_state.docx_bytes = docx_bytes

        with st.spinner("Generating PDF preview..."):
            try:
                pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes)
                st.session_state.pdf_bytes = pdf_bytes
                st.session_state.page_count = count_pdf_pages(pdf_bytes)
            except Exception:
                st.session_state.pdf_bytes = None
                st.session_state.page_count = None

        st.session_state.processing = False
        st.success("Done!")
        st.rerun()

if "result_data" in st.session_state:
    data = st.session_state.result_data
    filename = st.session_state.result_filename
    docx_bytes = st.session_state.get("docx_bytes") or generate_docx_bytes(data)
    stem = filename.rsplit(".", 1)[0]

    page_count = st.session_state.get("page_count")
    is_compact = data.get("compact_mode", False)

    info_parts = []
    if page_count:
        info_parts.append(f"**{page_count} page{'s' if page_count > 1 else ''}**")
    info_parts.append(f"**{data.get('name', 'Unknown')}**")
    st.markdown(" — ".join(info_parts))

    if is_compact:
        st.info("Compact mode was used. Content may have been shortened to fit one page.")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Download DOCX",
            data=docx_bytes,
            file_name=f"{stem}_formatted.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
    with col2:
        pdf_bytes = st.session_state.get("pdf_bytes")
        if pdf_bytes:
            st.download_button(
                label="Download PDF",
                data=pdf_bytes,
                file_name=f"{stem}_formatted.pdf",
                mime="application/pdf",
            )

    warnings = data.get("warnings", [])
    if warnings:
        with st.expander("Warnings", expanded=False):
            for w in warnings:
                st.warning(w)

    html = generate_html_preview(data)
    st.markdown(html, unsafe_allow_html=True)
