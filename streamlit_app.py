from pathlib import Path
import tempfile
import hashlib
import streamlit as st
from PIL import Image, ImageDraw
import io
import killerwhale_backend as kw

st.set_page_config(
    page_title="KillerWhale Extractor 3ThouWow",
    page_icon="🐋",
    layout="centered"
)


# Left-side animated GIF.
# Requires:
#   static/spiderman_side.gif
#   .streamlit/config.toml with enableStaticServing = true
st.markdown(
    """
<style>
[data-testid="stAppViewContainer"]::before {
    content: "";
    position: fixed;
    left: 300px;
    top: 50%;
    transform: translateY(-50%);
    width: 400px;
    height: 100vh;
    background-image: url("./app/static/spiderman_side.gif");
    background-position: center;
    background-repeat: no-repeat;
    background-size: 100% 100%;
    z-index: 1000;
    pointer-events: none;
}

[data-testid="stMainBlockContainer"] {
    position: relative;
    z-index: 10;
}

@media (max-width: 1350px) {
    [data-testid="stAppViewContainer"]::before {
        display: none;
    }
}
</style>
""",
    unsafe_allow_html=True,
)



st.markdown(
    """<h1 style="font-size:2.3rem;white-space:nowrap;">
    🐋 KillerWhale Extractor 3ThouWow 🐋
    </h1>""",
    unsafe_allow_html=True
)

legend_mode = st.radio(
    "Masking method",
    ["KillerWhale Way", "Manual Way"],
    horizontal=True,
    label_visibility="collapsed"
)

st.subheader("RE")
re_files = st.file_uploader(
    "RE PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key="re_files"
)

st.subheader("CE")
ce_files = st.file_uploader(
    "CE PDF + Excel file(s)",
    type=["pdf", "xlsx", "xlsm"],
    accept_multiple_files=True,
    key="ce_files"
)

st.subheader("Recalculation")
oats_files = st.file_uploader(
    "Recalculation PDF(s)",
    type=["pdf"],
    accept_multiple_files=True,
    key="oats_files"
)

oats_class = st.radio(
    "Recalculation Class",
    ["A", "B"],
    horizontal=True,
    format_func=lambda x: f"Class {x}",
    label_visibility="collapsed"
)

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
    with tempfile.TemporaryDirectory(
        prefix="killerwhale_web_"
    ) as tmp:

        root = Path(tmp)

        folders = {
            n: root / n
            for n in (
                "RE",
                "CE_PDF",
                "CE_Excel",
                "Recalculation"
            )
        }

        for d in folders.values():
            d.mkdir()

        re_paths = save_uploads(
            re_files,
            folders["RE"]
        )

        ce_pdf_paths = save_uploads(
            ce_pdfs,
            folders["CE_PDF"]
        )

        ce_excel_paths = save_uploads(
            ce_excels,
            folders["CE_Excel"]
        )

        oats_paths = save_uploads(
            oats_files,
            folders["Recalculation"]
        )

        result = kw.extract_multiple_pdfs(
            re_paths,
            ce_pdf_paths=ce_pdf_paths,
            ce_excel_paths=ce_excel_paths,
            oats_pdf_paths=oats_paths,
            oats_class=oats_class,
            progress_callback=callback
        )

        if preview_only:
            return None

        _, xlsx, rows = result

        return Path(xlsx).read_bytes(), rows


ce_pdfs, ce_excels = split_ce(ce_files)


signature = (
    tuple(
        (f.name, f.size)
        for f in (re_files or [])
    ),
    tuple(
        (f.name, f.size)
        for f in (ce_files or [])
    ),
    tuple(
        (f.name, f.size)
        for f in (oats_files or [])
    ),
)


if st.session_state.get(
    "upload_signature"
) != signature:

    st.session_state.upload_signature = signature

    st.session_state.pop(
        "mask_previews",
        None
    )

    st.session_state.pop(
        "mask_clicks",
        None
    )

    for _k in list(
        st.session_state.keys()
    ):
        if (
            str(_k).startswith("mask_drag_")
            or
            str(_k).startswith("mask_number_")
        ):
            del st.session_state[_k]


# ============================================================
# MANUAL WAY
# ============================================================

if legend_mode == "Manual Way":

    if st.button(
        "Manual Way",
        type="primary",
        use_container_width=True
    ):

        if not (
            re_files
            or ce_files
            or oats_files
        ):
            st.warning(
                "Select files first."
            )

        elif bool(ce_pdfs) != bool(ce_excels):

            st.error(
                "For CE, select the PDF(s) and matching Excel file(s) together."
            )

        else:

            try:

                kw.MANUAL_LEGEND_MASK = True
                kw.WEB_MASK_PREVIEW_MODE = True
                kw.WEB_MASK_PREVIEWS = []
                kw.WEB_MANUAL_MASK_INDEX = 0

                with st.spinner(
                    "Preparing the actual graph previews..."
                ):
                    run_backend(
                        preview_only=True
                    )

                st.session_state.mask_previews = list(
                    kw.WEB_MASK_PREVIEWS
                )

                st.session_state.mask_clicks = [
                    p["default_fraction"]
                    for p in kw.WEB_MASK_PREVIEWS
                ]

                kw.WEB_MASK_PREVIEW_MODE = False

                st.success(
                    f"{len(kw.WEB_MASK_PREVIEWS)} graph preview(s) ready."
                )

            except Exception as exc:

                kw.WEB_MASK_PREVIEW_MODE = False
                st.error(str(exc))


    previews = st.session_state.get(
        "mask_previews",
        []
    )

    clicks = st.session_state.get(
        "mask_clicks",
        []
    )


    if previews:

        st.markdown(
            "### Manual Snips"
        )

        for i, p in enumerate(previews):

            preview_title = str(
                p.get(
                    "title",
                    f"Graph {i+1}"
                )
            )

            pdf_label = preview_title.replace(
                "RE Legend Mask — drag red bar, then Apply",
                ""
            ).strip(" —-")


            if not pdf_label:

                all_pdf_names = (
                    [f.name for f in (re_files or [])]
                    +
                    [f.name for f in (ce_pdfs or [])]
                )

                pdf_label = (
                    all_pdf_names[i]
                    if i < len(all_pdf_names)
                    else f"Graph {i+1}"
                )


            st.markdown(
                f"**{pdf_label}**"
            )


            selected_fraction = clicks[i]

            if selected_fraction is None:
                selected_fraction = p[
                    "default_fraction"
                ]


            slider_key = (
                f"mask_drag_{i}"
            )

            number_key = (
                f"mask_number_{i}"
            )

            default_percent = float(
                selected_fraction * 100.0
            )


            if slider_key not in st.session_state:
                st.session_state[
                    slider_key
                ] = default_percent


            if number_key not in st.session_state:
                st.session_state[
                    number_key
                ] = default_percent


            def _slider_changed(
                sk=slider_key,
                nk=number_key
            ):
                st.session_state[nk] = float(
                    st.session_state[sk]
                )


            def _number_changed(
                sk=slider_key,
                nk=number_key
            ):
                st.session_state[sk] = float(
                    st.session_state[nk]
                )


            selected_percent = float(
                st.session_state[
                    slider_key
                ]
            )


            clicks[i] = (
                selected_percent / 100.0
            )

            st.session_state.mask_clicks = (
                clicks
            )


            base_image = Image.open(
                io.BytesIO(p["png"])
            ).convert("RGB")


            thumb_inset_px = 15.0

            display_w = float(
                p.get(
                    "display_width",
                    base_image.width
                )
            )

            inset_fraction = (
                thumb_inset_px
                /
                max(
                    1.0,
                    display_w
                )
            )


            aligned_fraction = (
                inset_fraction
                +
                (selected_percent / 100.0)
                *
                (
                    1.0
                    -
                    2.0 * inset_fraction
                )
            )


            bar_x = (
                int(
                    round(
                        base_image.width
                        *
                        aligned_fraction
                    )
                )
                + 3
            )


            draw = ImageDraw.Draw(
                base_image
            )


            draw.line(
                [
                    (bar_x, 0),
                    (
                        bar_x,
                        base_image.height
                    )
                ],
                fill="red",
                width=max(
                    4,
                    base_image.width // 220
                ),
            )


            st.image(
                base_image,
                width="stretch"
            )


            st.slider(
                "Red bar position",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                key=slider_key,
                on_change=_slider_changed,
                label_visibility="collapsed",
            )


            st.markdown(
                "Position"
            )


            st.number_input(
                "Position value",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.1f",
                key=number_key,
                on_change=_number_changed,
                label_visibility="collapsed",
                width=190,
            )


        st.divider()


# ============================================================
# EXTRACT
# ============================================================

manual_ready = bool(
    st.session_state.get(
        "mask_previews"
    )
)


extract_clicked = st.button(
    "EXTRACT",
    type="primary",
    use_container_width=True,
    disabled=(
        legend_mode == "Manual Way"
        and not manual_ready
    )
)


# ============================================================
# AUTO SNIP NOTE
# Only show this when KillerWhale Way is selected
# ============================================================

if legend_mode == "KillerWhale Way":

    st.markdown(
        """
**Note: Auto Snipping Order**  
**RE:** 30M-1G → 1G-6G → 1G-18G  
**CE:** AC Line 1-20 → AC Line → IPMI0-20 → IPMI → ETH0-20 → Tel Line
        """
    )


# ============================================================
# RUN EXTRACTION
# ============================================================

if extract_clicked:

    if not (
        re_files
        or ce_files
        or oats_files
    ):

        st.warning(
            "Select files first."
        )


    elif bool(ce_pdfs) != bool(ce_excels):

        st.error(
            "For CE, select both the CE PDF(s) and matching Excel file(s) together."
        )


    else:

        progress = st.progress(0)
        status = st.empty()

        try:

            def update(
                percent,
                message
            ):

                progress.progress(
                    max(
                        0,
                        min(
                            100,
                            int(percent)
                        )
                    )
                )

                status.write(
                    message
                )


            kw.MANUAL_LEGEND_MASK = (
                legend_mode
                ==
                "Manual Way"
            )

            kw.WEB_MASK_PREVIEW_MODE = False

            kw.WEB_MANUAL_MASK_INDEX = 0


            kw.WEB_MANUAL_MASK_FRACTIONS = (
                list(
                    st.session_state.get(
                        "mask_clicks",
                        []
                    )
                )
                if legend_mode
                ==
                "Manual Way"
                else []
            )


            output, rows = run_backend(
                preview_only=False,
                callback=update
            )


            progress.progress(100)


            status.success(
                f"Finished — {rows} data row(s) extracted."
            )


            st.download_button(
                "Download Excel",
                data=output,
                file_name=(
                    "EMC_All_Results_"
                    "RE_CE_Recalculation.xlsx"
                ),
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )


        except Exception as exc:

            progress.empty()
            status.empty()
            st.error(
                str(exc)
            )
