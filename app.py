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
    "Upload a DXF drawing and automatically generate a Bar Bending Schedule.\n"
    "The file is to be in dxf. Every member is to have an annotation.\n\n"
    "Requirements if length of bar is extracted:\n"
    "1. A solid red dot on the bar"
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
