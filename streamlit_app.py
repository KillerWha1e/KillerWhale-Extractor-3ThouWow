from pathlib import Path
import tempfile

import streamlit as st

import killerwhale_backend as kw

st.set_page_config(page_title="KillerWhale Extractor 3ThouWow", page_icon="🐋", layout="centered")

st.markdown(
    """
    <h1 style="font-size: 2.3rem; white-space: nowrap;">
        🐋 KillerWhale Extractor 3ThouWow 🐋
    </h1>
    """,
    unsafe_allow_html=True,
)

st.subheader("RE")
re_files = st.file_uploader(
    "RE PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key="re_files",
)

st.subheader("CE")
ce_pdf_files = st.file_uploader(
    "CE PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key="ce_pdf_files",
)
ce_excel_files = st.file_uploader(
    "CE Excel file(s)",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    key="ce_excel_files",
    help="Each CE Excel filename must match its CE PDF filename exactly, except for the extension.",
)

st.subheader("OATS")
oats_files = st.file_uploader(
    "OATS PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key="oats_files",
)
oats_class = st.radio(
    "OATS Class",
    options=["A", "B"],
    horizontal=True,
    format_func=lambda x: f"Class {x}",
)

st.divider()


def save_uploads(uploaded_files, folder: Path):
    paths = []
    for uploaded in uploaded_files or []:
        path = folder / uploaded.name
        path.write_bytes(uploaded.getbuffer())
        paths.append(path)
    return paths


if st.button("EXTRACT", type="primary", use_container_width=True):
    if not (re_files or ce_pdf_files or ce_excel_files or oats_files):
        st.warning("Select RE PDF(s), CE PDF + Excel pair(s), or OATS PDF(s) first.")
    elif bool(ce_pdf_files) != bool(ce_excel_files):
        st.error("For CE, select both CE PDF file(s) and CE Excel file(s).")
    else:
        progress = st.progress(0)
        status = st.empty()

        try:
            with tempfile.TemporaryDirectory(prefix="killerwhale_web_") as tmp:
                tmpdir = Path(tmp)
                re_dir = tmpdir / "RE"
                ce_pdf_dir = tmpdir / "CE_PDF"
                ce_excel_dir = tmpdir / "CE_Excel"
                oats_dir = tmpdir / "OATS"
                for d in (re_dir, ce_pdf_dir, ce_excel_dir, oats_dir):
                    d.mkdir(parents=True, exist_ok=True)

                re_paths = save_uploads(re_files, re_dir)
                ce_pdf_paths = save_uploads(ce_pdf_files, ce_pdf_dir)
                ce_excel_paths = save_uploads(ce_excel_files, ce_excel_dir)
                oats_paths = save_uploads(oats_files, oats_dir)

                def update_progress(percent, message):
                    progress.progress(max(0, min(100, int(percent))))
                    status.write(message)

                out_dir, output_xlsx, total_rows = kw.extract_multiple_pdfs(
                    re_paths,
                    ce_pdf_paths=ce_pdf_paths,
                    ce_excel_paths=ce_excel_paths,
                    oats_pdf_paths=oats_paths,
                    oats_class=oats_class,
                    progress_callback=update_progress,
                )

                output_bytes = Path(output_xlsx).read_bytes()

            progress.progress(100)
            status.success(f"Finished — {total_rows} data row(s) extracted.")
            st.download_button(
                "Download Excel",
                data=output_bytes,
                file_name="EMC_All_Results_RE_CE_OATS.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            progress.empty()
            status.empty()
            st.error(str(exc))

st.caption("Files are processed during the current app session. For confidential customer data, use only hosting approved by your company.")
