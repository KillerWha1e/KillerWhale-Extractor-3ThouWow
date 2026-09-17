from pathlib import Path
import tempfile
import hashlib
import streamlit as st
from streamlit_image_coordinates import streamlit_image_coordinates
from PIL import Image, ImageDraw
import io
import killerwhale_backend as kw

st.set_page_config(page_title="KillerWhale Extractor 3ThouWow", page_icon="🐋", layout="centered")
st.markdown("""<h1 style="font-size:2.3rem;white-space:nowrap;">🐋 KillerWhale Extractor 3ThouWow 🐋</h1>""", unsafe_allow_html=True)

st.markdown("#### Legend Masking")
legend_mode = st.radio("Masking method", ["KillerWhale Way", "Manual Way"],
                       horizontal=True, label_visibility="collapsed")
if legend_mode == "Manual Way":
    st.caption("Old-style Manual Way: prepare the previews, then click directly on each graph to move the red cut bar.")

st.subheader("RE")
re_files = st.file_uploader("RE PDF(s)", type=["pdf"], accept_multiple_files=True, key="re_files")

st.subheader("CE")
ce_files = st.file_uploader(
    "CE PDF + Excel file(s)",
    type=["pdf", "xlsx", "xlsm"],
    accept_multiple_files=True,
    key="ce_files",
    help="Select the matching CE PDF and Excel files together in this one row."
)

st.subheader("Recalculation")
oats_files = st.file_uploader("Recalculation PDF(s)", type=["pdf"], accept_multiple_files=True, key="oats_files")
oats_class = st.radio("Recalculation Class", ["A", "B"], horizontal=True, format_func=lambda x: f"Class {x}")
st.divider()

def split_ce(files):
    pdfs, excels = [], []
    for f in files or []:
        ext = Path(f.name).suffix.lower()
        if ext == ".pdf":
            pdfs.append(f)
        elif ext in (".xlsx", ".xlsm"):
            excels.append(f)
    return pdfs, excels

def save_uploads(files, folder):
    paths = []
    for f in files or []:
        p = folder / f.name
        p.write_bytes(f.getbuffer())
        paths.append(p)
    return paths

def run_backend(preview_only=False, callback=None):
    with tempfile.TemporaryDirectory(prefix="killerwhale_web_") as tmp:
        root = Path(tmp)
        folders = {n: root/n for n in ("RE", "CE_PDF", "CE_Excel", "Recalculation")}
        for d in folders.values():
            d.mkdir()
        re_paths = save_uploads(re_files, folders["RE"])
        ce_pdf_paths = save_uploads(ce_pdfs, folders["CE_PDF"])
        ce_excel_paths = save_uploads(ce_excels, folders["CE_Excel"])
        oats_paths = save_uploads(oats_files, folders["Recalculation"])
        result = kw.extract_multiple_pdfs(
            re_paths, ce_pdf_paths=ce_pdf_paths, ce_excel_paths=ce_excel_paths,
            oats_pdf_paths=oats_paths, oats_class=oats_class, progress_callback=callback
        )
        if preview_only:
            return None
        _, xlsx, rows = result
        return Path(xlsx).read_bytes(), rows

ce_pdfs, ce_excels = split_ce(ce_files)

signature = (
    tuple((f.name, f.size) for f in (re_files or [])),
    tuple((f.name, f.size) for f in (ce_files or [])),
    tuple((f.name, f.size) for f in (oats_files or [])),
)
if st.session_state.get("upload_signature") != signature:
    st.session_state.upload_signature = signature
    st.session_state.pop("mask_previews", None)
    st.session_state.pop("mask_clicks", None)

if legend_mode == "Manual Way":
    if st.button("PREPARE MANUAL SNIPS", use_container_width=True):
        if not (re_files or ce_files or oats_files):
            st.warning("Select files first.")
        elif bool(ce_pdfs) != bool(ce_excels):
            st.error("For CE, select the PDF(s) and matching Excel file(s) together.")
        else:
            try:
                kw.MANUAL_LEGEND_MASK = True
                kw.WEB_MASK_PREVIEW_MODE = True
                kw.WEB_MASK_PREVIEWS = []
                kw.WEB_MANUAL_MASK_INDEX = 0
                with st.spinner("Preparing the actual graph previews..."):
                    run_backend(preview_only=True)
                st.session_state.mask_previews = list(kw.WEB_MASK_PREVIEWS)
                st.session_state.mask_clicks = [p["default_fraction"] for p in kw.WEB_MASK_PREVIEWS]
                kw.WEB_MASK_PREVIEW_MODE = False
                st.success(f"{len(kw.WEB_MASK_PREVIEWS)} graph preview(s) ready.")
            except Exception as exc:
                kw.WEB_MASK_PREVIEW_MODE = False
                st.error(str(exc))

    previews = st.session_state.get("mask_previews", [])
    clicks = st.session_state.get("mask_clicks", [])
    if previews:
        st.markdown("### Manual Snips")
        st.caption("A red vertical bar shows the cut position. Click anywhere on the graph to move the red bar. Everything to the right of the bar on the detected legend row will be removed.")
        for i, p in enumerate(previews):
            st.markdown(f"**Graph {i+1}: {p['title']}**")
            # Build a preview with a visible red vertical cut bar.
            base_image = Image.open(io.BytesIO(p["png"])).convert("RGB")

            selected_fraction = clicks[i]
            if selected_fraction is None:
                selected_fraction = p["default_fraction"]

            bar_x = int(round(base_image.width * selected_fraction))
            draw = ImageDraw.Draw(base_image)
            draw.line(
                [(bar_x, 0), (bar_x, base_image.height)],
                fill="red",
                width=max(3, base_image.width // 250),
            )

            preview_hash = hashlib.sha1(
                p["png"] + str(selected_fraction).encode("utf-8")
            ).hexdigest()[:12]
            preview_path = Path(tempfile.gettempdir()) / f"killerwhale_mask_{i}_{preview_hash}.png"
            base_image.save(preview_path)

            # The image itself is clickable. Click anywhere to move the red bar.
            point = streamlit_image_coordinates(
                str(preview_path),
                width=p["display_width"],
                key=f"mask_graph_{i}_{preview_hash}",
            )
            if point and "x" in point:
                new_fraction = max(
                    0.0,
                    min(1.0, point["x"] / max(1, p["display_width"]))
                )
                if clicks[i] != new_fraction:
                    clicks[i] = new_fraction
                    st.session_state.mask_clicks = clicks
                    st.rerun()

            if clicks[i] is None:
                st.caption("Mask skipped.")
            else:
                st.caption("Click the graph to move the red bar.")

            c1, c2 = st.columns(2)
            if c1.button("Reset", key=f"reset_{i}"):
                clicks[i] = p["default_fraction"]
                st.session_state.mask_clicks = clicks
                st.rerun()
            if c2.button("Skip Mask", key=f"skip_{i}"):
                clicks[i] = None
                st.session_state.mask_clicks = clicks
                st.rerun()
        st.divider()

manual_ready = bool(st.session_state.get("mask_previews"))
if st.button("EXTRACT", type="primary", use_container_width=True,
             disabled=(legend_mode == "Manual Way" and not manual_ready)):
    if not (re_files or ce_files or oats_files):
        st.warning("Select files first.")
    elif bool(ce_pdfs) != bool(ce_excels):
        st.error("For CE, select both the CE PDF(s) and matching Excel file(s) together.")
    else:
        progress = st.progress(0)
        status = st.empty()
        try:
            def update(percent, message):
                progress.progress(max(0, min(100, int(percent))))
                status.write(message)

            kw.MANUAL_LEGEND_MASK = legend_mode == "Manual Way"
            kw.WEB_MASK_PREVIEW_MODE = False
            kw.WEB_MANUAL_MASK_INDEX = 0
            kw.WEB_MANUAL_MASK_FRACTIONS = (
                list(st.session_state.get("mask_clicks", []))
                if legend_mode == "Manual Way" else []
            )

            output, rows = run_backend(preview_only=False, callback=update)
            progress.progress(100)
            status.success(f"Finished — {rows} data row(s) extracted.")
            st.download_button(
                "Download Excel", data=output,
                file_name="EMC_All_Results_RE_CE_Recalculation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as exc:
            progress.empty()
            status.empty()
            st.error(str(exc))
