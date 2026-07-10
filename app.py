import streamlit as st
from pathlib import Path
import tempfile

# 1. This must be called first with real, supported arguments
st.set_page_config(
    page_title="CIVSOL BBS Extraction",
    layout="wide"
)

# 2. Injected CSS to hide the Top Header Bar, Hamburger Menu, and Footer
hide_elements_css = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stAppDeployButton {display: none;}
    </style>
"""
st.markdown(hide_elements_css, unsafe_allow_html=True)

# 3. Rest of your imports
from dxf_bbs_extractor import process_dxf

st.title("CIVSOL BBS Automation")

st.write(
    "Upload a DXF drawing and automatically generate a Bar Bending Schedule. "
    "The file must be in DXF format, and every bar must have its own annotation."
)

with st.expander("Drawing requirements for reliable bar length extraction", expanded=True):
    st.markdown(
        "For a bar length to be extracted with confidence, every bar needs:\n\n"
        "1. **A solid dot, centred on the bar itself.** This is the anchor point "
        "every bar mark is matched from.\n"
        "2. **The dot connected to its text label**, using one of these three "
        "drafting styles:\n"
        "   - An aligned dimension where the dimension and the label text are "
        "one entity. This is the most reliable style.\n"
        "   - A leader line running from the dot to a separate text label, with "
        "no gaps in the line work.\n"
        "   - An aligned dimension with no text of its own, continued by a "
        "leader line from the dimension out to a separate text label.\n"
        "3. **Exactly one label per bar.** If two dots, leader paths, or "
        "dimensions could equally plausibly belong to the same label, the tool "
        "will not guess. It flags the bar for you to check by hand instead.\n\n"
        "A bar that doesn't meet any of these will show up in the report as "
        "unmatched, with a reason, rather than being silently given a wrong or "
        "made-up length."
    )

uploaded_file = st.file_uploader(
    "Upload DXF Drawing",
    type=["dxf"]
)

if uploaded_file:
    st.success(
        f"Loaded: {uploaded_file.name}"
    )
    member_name = st.text_input(
        "Member Name",
        value="MEMBER"
    )
    if st.button("Extract BBS"):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf_path = Path(temp_dir) / uploaded_file.name
            with open(dxf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            result = process_dxf(
                dxf_path=str(dxf_path),
                outdir=temp_dir,
                member_name=member_name,
            )
            st.success("Extraction Complete")
            st.subheader("Summary")
            st.dataframe(
                result["summary_df"],
                use_container_width=True
            )
            with open(
                result["bbs_path"],
                "rb"
            ) as f:
                st.download_button(
                    label="Download BBS Excel",
                    data=f.read(),
                    file_name="BBS_from_DXF.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
