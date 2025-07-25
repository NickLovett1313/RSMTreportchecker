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
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    unique_spartans = df['CONTACT_NM'].dropna().unique().tolist()
    unique_spartans.sort()

    st.markdown("### ✅ Step 1: Choose Spartans")
    # Use a form so nothing runs until you click "Run Analysis"
    with st.form("spartan_form"):
        selected_spartans = st.multiselect("Select Spartans to check:", options=unique_spartans)
        run = st.form_submit_button("🚀 Run Analysis")

    if run:
        summary_data = []
        for spartan in selected_spartans:
            spartan_df = df[df['CONTACT_NM'] == spartan]
            unique_pos = spartan_df['PO'].dropna().unique().tolist()

            awaiting_shipping_pos = []
            tbd_ship_to_pos = []

            for po in unique_pos:
                po_df = spartan_df[spartan_df['PO'] == po]
                if (po_df['LINE_STATUS'] == 'AWAITING_SHIPPING').any():
                    try:
                        clean_po = str(int(float(po)))
                    except:
                        clean_po = str(po)
                    awaiting_shipping_pos.append(clean_po)

                    if (po_df['SHIP_TO_CUSTOMER'] == 'TO BE DETERMINED').any():
                        tbd_ship_to_pos.append(clean_po)

            summary_data.append({
                'Spartan': spartan,
                'Awaiting Shipping POs': ', '.join(awaiting_shipping_pos) if awaiting_shipping_pos else 'None',
                'TBD Ship To POs': ', '.join(tbd_ship_to_pos) if tbd_ship_to_pos else 'None'
            })

        summary_df = pd.DataFrame(summary_data)
        summary_df = summary_df[['Spartan', 'Awaiting Shipping POs', 'TBD Ship To POs']]
        st.subheader("📊 Summary Table")
        st.dataframe(summary_df, use_container_width=True)
        st.session_state['summary_df'] = summary_df

    # --- Generate Email ---
    if 'summary_df' in st.session_state and st.button("📋 Ready to send to team?"):
        summary_df = st.session_state['summary_df']
        formatted_date = format_date_suffix(datetime.today())

        subject = f"Rosemount Orders – Daily Open Orders Report Review: {formatted_date}"
        body_intro = (
            "Hi Team,\n\n"
            "The Rosemount Daily Open Orders Report has been reviewed for all of you CC'd on this email.\n"
            "See your section below for details on your orders.\n\n"
            "Thanks!\n"
        )

        # Build the bullet list into the body
        lines = [body_intro]
        for row in summary_df.itertuples(index=False):
            spartan, awaiting, tbd = row
            lines.append(f"- **{spartan}**")
            lines.append(f"    - Awaiting Shipping POs: {awaiting}")
            lines.append(f"    - TBD Ship To POs: {tbd}")
        body_full = "\n".join(lines)

        st.markdown("### ✉️ Email Subject", unsafe_allow_html=True)
        st.markdown(f"<div style='font-family: Arial'>{subject}</div>", unsafe_allow_html=True)

        st.markdown("### 📩 Email Body", unsafe_allow_html=True)
        st.markdown(f"<div style='font-family: Arial; white-space: pre-wrap'>{body_full}</div>", unsafe_allow_html=True)

else:
    st.info("👆 Upload an Excel file to get started.")
