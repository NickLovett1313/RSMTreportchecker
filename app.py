import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Awaiting Shipping Checker")
st.title("📦 Awaiting Shipping Checker")

uploaded_file = st.file_uploader("Upload the latest Excel sheet", type=["xlsx"])

# Format today's date nicely
def format_date_suffix(date_obj):
    day = int(date_obj.strftime("%d"))
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return date_obj.strftime(f"%B {day}{suffix}, %Y")

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    unique_reps = df['CONTACT_NM'].dropna().unique().tolist()
    unique_reps.sort()

    with st.form("rep_form"):
        selected_reps = st.multiselect("Select reps to check:", options=unique_reps)
        submitted = st.form_submit_button("🚀 Run Analysis")

    if submitted and selected_reps:
        summary_data = []

        for rep in selected_reps:
            rep_df = df[df['CONTACT_NM'] == rep]
            unique_pos = rep_df['PO'].dropna().unique().tolist()

            awaiting_shipping_pos = []
            tbd_ship_to_pos = []

            for po in unique_pos:
                po_df = rep_df[rep_df['PO'] == po]

                if (po_df['LINE_STATUS'] == 'AWAITING_SHIPPING').any():
                    try:
                        clean_po = str(int(float(po)))
                    except:
                        clean_po = str(po)
                    awaiting_shipping_pos.append(clean_po)

                    if (po_df['SHIP_TO_CUSTOMER'] == 'TO BE DETERMINED').any():
                        tbd_ship_to_pos.append(clean_po)

            summary_data.append({
                'Rep Name': rep,
                'Awaiting Shipping POs': '\n'.join(awaiting_shipping_pos) if awaiting_shipping_pos else 'None',
                'TBD Ship To POs': '\n'.join(tbd_ship_to_pos) if tbd_ship_to_pos else 'None'
            })

        st.subheader("📊 Summary Table")
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True)

        # --- Generate email view ---
        if st.button("📋 Ready to send to team?"):
            formatted_date = format_date_suffix(datetime.today())

            subject = f"Rosemount Orders – Daily Open Orders Report Review: {formatted_date}"
            body = """
Hi Team,

The Rosemount Daily Open Orders Report has been reviewed for all of you CC'd on this email.
See the table below and find your name for information on your orders.

Thanks!
"""

            st.markdown(f"**✉️ Email Subject:** `{subject}`")
            st.markdown("**📩 Email Body:**")
            st.code(body.strip())

            # Convert DataFrame to HTML table
            html_table = summary_df.to_html(index=False, escape=False).replace('\n', '<br>')

            st.markdown("**📎 Copyable Table for Outlook/Gmail:**", unsafe_allow_html=True)
            st.markdown(html_table, unsafe_allow_html=True)

else:
    st.info("👆 Upload an Excel file to get started.")
