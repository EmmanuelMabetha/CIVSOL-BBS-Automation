import streamlit as st
from pathlib import Path
import tempfile
import pandas as pd

# Import the pipeline
from dxf_bbs_pipeline import process_dxf

st.set_page_config(
    page_title="CIVSOL BBS Extraction",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        color: #155724;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeeba;
        border-radius: 5px;
        padding: 1rem;
        color: #856404;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        color: #0c5460;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">CIVSOL BBS Automation</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Extract Bar Bending Schedules directly from DXF drawings</div>', unsafe_allow_html=True)

# Sidebar configuration
with st.sidebar:
    st.header("⚙️ Configuration")

    member_name = st.text_input(
        "Member Name",
        value="MEMBER",
        help="Name shown in column A of the BBS (e.g., FOOTING, BEAM, COLUMN)"
    )

    layer_filter = st.text_input(
        "Layer Filter (optional)",
        value="",
        help="Comma-separated layer names to scan. Leave empty to scan all layers."
    )

    st.divider()
    st.markdown("**Requirements for length extraction:**")
    st.markdown("- Solid red dot (HATCH/CIRCLE/DONUT) on the bar")
    st.markdown("- Bar drawn as LINE on REINFORCEMENT layer")
    st.markdown("- Label text with format: `<count><type><dia>-<mark>`")
    st.markdown("- Example: `6Y12-15` or `20Y8-01-300 LINKS`")

    st.divider()
    st.markdown("**Color Key:**")
    st.markdown("🟩 Green = Extracted from DXF (trusted)")
    st.markdown("🟨 Yellow = Needs manual verification")
    st.markdown("🟦 Blue = Link/stirrup (flagged, not auto-numbered)")
    st.markdown("🟩 Green (Shape col) = shape code auto-detected from traced geometry")

# Main content
uploaded_file = st.file_uploader(
    "📁 Upload DXF Drawing",
    type=["dxf"],
    help="Upload a DXF file exported from AutoCAD/BricsCAD"
)

if uploaded_file is not None:
    # Save uploaded file
    with tempfile.TemporaryDirectory() as temp_dir:
        dxf_path = Path(temp_dir) / uploaded_file.name
        with open(dxf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"Loaded: **{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} KB)")

        # Process button
        col1, col2, col3 = st.columns([1, 1, 3])
        with col1:
            process_clicked = st.button("🚀 Extract BBS", type="primary", use_container_width=True)

        if process_clicked:
            with st.spinner("Processing DXF... This may take a moment"):
                # Parse layer filter
                layers = None
                if layer_filter.strip():
                    layers = [l.strip() for l in layer_filter.split(",") if l.strip()]

                # Run pipeline
                result = process_dxf(
                    dxf_path=str(dxf_path),
                    outdir=temp_dir,
                    member_name=member_name,
                    layers=layers,
                )

            # Display logs
            with st.expander("📋 Processing Log", expanded=False):
                for log in result["logs"]:
                    if log.startswith("[!]"):
                        st.warning(log)
                    else:
                        st.text(log)

            # Stats dashboard
            st.subheader("📊 Extraction Summary")
            stats = result["stats"]

            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.metric("Labels Found", stats["parsed"], f"{stats['unparsed']} unparsed")
            with c2:
                st.metric("Dots Total", stats["dots_total"], f"{stats['dots_matched']} matched")
            with c3:
                st.metric("Bar Marks", stats["marks_found"])
            with c4:
                st.metric("Conflicts", stats["conflicts"])
            with c5:
                if stats["missing_marks"]:
                    st.metric("Missing Marks", len(stats["missing_marks"]))
                else:
                    st.metric("Missing Marks", 0)

            # Missing marks warning
            if stats["missing_marks"]:
                st.warning(
                    f"⚠️ Gap in bar mark numbering: marks **{', '.join(str(m) for m in stats['missing_marks'])}** "
                    f"do not appear on the drawing but sit between marks that do. Verify these were not omitted."
                )

            # Summary table
            st.subheader("📋 Bar Mark Summary")

            summary_df = result["summary_df"]
            if not summary_df.empty:
                # Format the dataframe for display
                display_df = summary_df.copy()

                # Highlight links
                def highlight_links(row):
                    if row.get("Is Link"):
                        return ['background-color: #CFE2F3'] * len(row)
                    return [''] * len(row)

                st.dataframe(
                    display_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Bar Mark": st.column_config.TextColumn("Mark", width="small"),
                        "Type": st.column_config.TextColumn("Type", width="small"),
                        "Diameter (mm)": st.column_config.NumberColumn("Dia", width="small"),
                        "Total No. Off": st.column_config.NumberColumn("Total Off", width="small"),
                        "Distinct DXF Instances Traced": st.column_config.NumberColumn("Traced #", width="small", help="Distinct geometrically-matched instances found for this mark -- compare against Total Off"),
                        "Longest DXF Length (mm)": st.column_config.NumberColumn("Length (mm)", width="medium"),
                        "Suggested Shape Code (verify)": st.column_config.TextColumn("Shape", width="small"),
                        "Flag": st.column_config.TextColumn("Flags", width="large"),
                    }
                )
            else:
                st.info("No bar marks found in the drawing.")

            # Dot matches table
            st.subheader("🔍 Dot-to-Bar Matching Details")

            dot_df = result["dot_matches_df"]
            if not dot_df.empty:
                # Create tabs for different match methods
                tab1, tab2, tab3 = st.tabs(["All Matches", "Accepted Only", "Unmatched/Errors"])

                with tab1:
                    st.dataframe(dot_df, use_container_width=True, hide_index=True)

                with tab2:
                    accepted = dot_df[dot_df["Accepted"] == True]
                    st.dataframe(accepted, use_container_width=True, hide_index=True)
                    st.success(f"✅ {len(accepted)} dots accepted with high confidence")

                with tab3:
                    unmatched = dot_df[dot_df["Accepted"] == False]
                    if not unmatched.empty:
                        st.dataframe(unmatched, use_container_width=True, hide_index=True)
                        st.info(f"ℹ️ {len(unmatched)} dots need manual review")
                    else:
                        st.success("All dots matched successfully!")

            # Download section
            st.subheader("⬇️ Downloads")

            dl_col1, dl_col2, dl_col3, dl_col4 = st.columns(4)

            with dl_col1:
                with open(result["bbs_path"], "rb") as f:
                    st.download_button(
                        label="📊 BBS Excel",
                        data=f.read(),
                        file_name=f"BBS_{member_name}_{uploaded_file.name.replace('.dxf', '')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            with dl_col2:
                with open(result["summary_path"], "rb") as f:
                    st.download_button(
                        label="📋 Summary CSV",
                        data=f.read(),
                        file_name=f"summary_{uploaded_file.name.replace('.dxf', '')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

            with dl_col3:
                with open(result["report_path"], "rb") as f:
                    st.download_button(
                        label="📄 Full Report CSV",
                        data=f.read(),
                        file_name=f"report_{uploaded_file.name.replace('.dxf', '')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

            with dl_col4:
                with open(result["dot_match_path"], "rb") as f:
                    st.download_button(
                        label="🔍 Dot Matches CSV",
                        data=f.read(),
                        file_name=f"dot_matches_{uploaded_file.name.replace('.dxf', '')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

else:
    # Empty state
    st.info("👆 Upload a DXF file to get started")

    # Example of expected label format
    with st.expander("📝 Expected Label Format"):
        st.markdown("""
        The tool looks for text labels in this format:

        ```
        6Y12-15 (3T,3B)      → 6 bars, Y-type, 12mm dia, mark 15
        20Y8-01-300 LINKS    → 20 links, Y-type, 8mm dia, mark 01, 300mm spacing
        3H16-07              → 3 bars, H-type, 16mm dia, mark 07
        ```

        **Pattern:** `<count><type><diameter>-<mark>[-<spacing>] [suffix]`

        - **count**: Number of bars (integer)
        - **type**: Steel type — Y (high yield), R (mild), H (hard drawn)
        - **diameter**: Bar diameter in mm
        - **mark**: Bar mark number
        - **spacing**: Optional — spacing for links/stirrups
        - **suffix**: Optional — "LINKS", "STIRRUP", "TIE" identifies links
        """)

    with st.expander("🎨 Drawing Requirements"):
        st.markdown("""
        For accurate extraction, your DXF should have:

        1. **Bar lines** on layer `REINFORCEMENT` or `REINF`
        2. **Dimension dots** (solid circles) on layer `DIMENSION` or `DIMENSIONS`
        3. **Leader lines** connecting dots to labels on `DIMENSION` layer
        4. **Labels** as DIMENSION, TEXT, or MTEXT entities

        The tool traces:
        ```
        DOT (on bar) → leader line → leader line → ... → LABEL (text)
        ```

        If the label is directly traceable from the dot by a clear straight
        leader path, it will match even when there is no explicit multi-segment
        chain.
        """)
