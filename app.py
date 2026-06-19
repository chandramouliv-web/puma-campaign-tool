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

st.markdown("---")

voucher_type = st.radio(
    "Voucher Type",
    [
        "Regular VC",
        "Bundle Discount"
    ]
)

voucher_pct = st.number_input(
    "Voucher %",
    min_value=1,
    max_value=90,
    value=20
)

st.markdown("---")

st.subheader("Pricing Simulator")

md_rrp = st.number_input(
    "MD RRP",
    value=2990.0
)

md_srp = st.number_input(
    "MD SRP",
    value=1990.0
)

campaign_price = (
    md_srp if md_srp > 0
    else md_rrp
)

campaign_discount = (
    (
        md_rrp-campaign_price
    )/md_rrp
)*100

final_price = (
    campaign_price
    *
    (1-voucher_pct/100)
)

total_discount = (
    (
        md_rrp-final_price
    )/md_rrp
)*100

c1,c2,c3,c4 = st.columns(4)

c1.metric(
    "Campaign Price",
    round(campaign_price,2)
)

c2.metric(
    "Campaign Discount %",
    round(campaign_discount,2)
)

c3.metric(
    "Final Price",
    round(final_price,2)
)

c4.metric(
    "Total Discount %",
    round(total_discount,2)
)

st.markdown("---")

st.subheader("Validation")

status="PASS"

if campaign_price >= md_rrp:
    status="FAIL"

if total_discount > 80:
    status="FAIL"

if final_price <= 0:
    status="FAIL"

if status=="PASS":
    st.success("PASS")
else:
    st.error("FAIL")
