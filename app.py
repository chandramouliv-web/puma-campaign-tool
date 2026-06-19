import streamlit as st

st.set_page_config(
    page_title="PUMA Campaign Tool",
    page_icon="🏷️",
    layout="wide"
)

st.title("🏷️ PUMA Campaign & Voucher Management Tool")

region = st.selectbox(
    "Region",
    ["PH", "MY", "SG"]
)

marketplace = st.selectbox(
    "Marketplace",
    ["Lazada", "Shopee", "Zalora", "TikTok"]
)

st.success(
    f"Selected: {region} / {marketplace}"
)
