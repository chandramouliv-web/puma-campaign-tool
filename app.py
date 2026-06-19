import streamlit as st

st.set_page_config(
    page_title="PUMA Campaign Tool",
    page_icon="🏷️",
    layout="wide"
)

st.title("🏷️ PUMA Campaign Tool")

col1,col2 = st.columns(2)

with col1:
    region = st.selectbox(
        "Region",
        ["PH","MY","SG"]
    )

with col2:
    marketplace = st.selectbox(
        "Marketplace",
        [
            "Lazada",
            "Shopee",
            "Zalora",
            "TikTok"
        ]
    )

st.success(
    f"{region} | {marketplace}"
)

st.markdown("---")

st.subheader("📂 Upload Files")

zecom_file = st.file_uploader(
    "ZeCom Tracker",
    type=["xlsx"]
)

content_file = st.file_uploader(
    "Content File",
    type=["xlsx"]
)

inventory_file = st.file_uploader(
    "Inventory File",
    type=["xlsx","csv"]
)

marketplace_file = st.file_uploader(
    "Marketplace Export",
    type=["xlsx","zip"]
)
