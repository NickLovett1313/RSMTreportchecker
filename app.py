import streamlit as st
import pandas as pd
from datetime import datetime

# --- Page title ---
st.set_page_config(page_title="Awaiting Shipping Checker")
st.title("📦 Awaiting Shipping Checker")

# --- Format date for subject ---
def format_date_suffix(date_obj):
    day = int(date_obj.strftime("%d"))
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return date_obj.strftime(f"%B {day}{suffix}, %Y")

# --- Upload ---
uploaded_file = st.file_uploader("Upload the latest Excel sheet", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    spartans = sorted(df["CONTACT_NM"].dropna().unique().tolist())

    st.markdown("### ✅ Step 1: Choose Spartans")
    with st.form("spartan_form"):
        selected = st.multiselect("Select Spartans to check:", options=spartans)
        run = st.form_submit_button("🚀 Run Analysis")

    if run:
        data = []
        for s in selected:
            sub = df[df["CONTACT_NM"] == s]
            pos = sub["PO"].dropna().unique().tolist()
            a, t = [], []
            for po in pos:
                block = sub[sub["PO"] == po]
                if (block["LINE_STATUS"] == "AWAITING_SHIPPING").any():
                    clean = str(int(float(po))) if str(po).replace(".0", "").isdigit() else str(po)
                    a.append(clean)
                    if (block["SHIP_TO_CUSTOMER"] == "TO BE DETERMINED").any():
                        t.append(clean)
            data.append({
                "Spartan": s,
                "Awaiting Shipping POs": ", ".join(a) or "None",
                "TBD Ship To POs":      ", ".join(t) or "None"
            })

        summary_df = pd.DataFrame(data)
        st.subheader("📊 Summary Table")
        st.dataframe(summary_df, use_container_width=True)
        st.session_state["summary_df"] = summary_df

    if "summary_df" in st.session_state and st.button("📋 Ready to send to team?"):
        summary_df = st.session_state["summary_df"]
        date_str = format_date_suffix(datetime.today())

        # Build copy-paste text
        lines = []
        lines.append("Hi Team,")
        lines.append("")
        lines.append("The Daily Open Orders Report for your Rosemount purchase orders has been reviewed for")
        lines.append("those CC’d.")
        lines.append("")
        lines.append("Note: for those PO#s with items awaiting shipment: If you haven’t yet received a packing")
        lines.append("slip for release, I recommend reaching out to your factory contact.")
        lines.append("")
        lines.append("Note: for those PO#s with a TBD ship-to address: This information must be provided to")
        lines.append("the factory before they can issue a packing slip.")
        lines.append("")
        lines.append("See information below:")
        lines.append("")
        lines.append("----------------------------")
        lines.append("")

        for idx, row in enumerate(summary_df.itertuples(index=False), start=1):
            s, a, tbd = row
            lines.append(f"{idx}. {s}")
            lines.append(f"o PO#s Awaiting Shipping: {a}")
            lines.append(f"o PO#s with TBD Ship-To Address: {tbd}")
            lines.append("")

        lines.append("----------------------------")
        lines.append("")
        lines.append("Thanks!")

        subject = f"Rosemount Orders – Daily Open Orders Report Review: {date_str}"

        st.markdown("### ✉️ Email Subject")
        st.code(subject, language="")

        st.markdown("### 📩 Email Body")
        st.code("\n".join(lines), language="")
else:
    st.info("👆 Upload an Excel file to get started.")
