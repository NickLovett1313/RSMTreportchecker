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
    unique_reps = df['CONTACT_NM'].dropna().unique().tolist()
    unique_reps.sort()

    st.markdown("### ✅ Step 1: Choose reps")
    # Use a form so nothing runs until you click "Run Analysis"
    with st.form("rep_form"):
        selected_reps = st.multiselect("Select reps to check:", options=unique_reps)
        run = st.form_submit_button("🚀 Run Analysis")

    if run:
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
                'Awaiting Shipping POs': ', '.join(awaiting_shipping_pos) if awaiting_shipping_pos else 'None',
                'TBD Ship To POs': ', '.join(tbd_ship_to_pos) if tbd_ship_to_pos else 'None'
            })

        summary_df = pd.DataFrame(summary_data)
        st.subheader("📊 Summary Table")
        st.dataframe(summary_df, use_container_width=True)
        st.session_state['summary_df'] = summary_df

        csv = summary_df.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Download CSV", data=csv, file_name='awaiting_shipping_summary.csv', mime='text/csv')

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
            rep, awaiting, tbd = row
            lines.append(f"- **{rep}**")
            lines.append(f"    - Awaiting Shipping POs: {awaiting}")
            lines.append(f"    - TBD Ship To POs: {tbd}")
        body_full = "\n".join(lines)

        st.markdown("### ✉️ Email Subject")
        st.code(subject, language='')

        st.markdown("### 📩 Email Body")
        st.code(body_full, language='')

else:
    st.info("👆 Upload an Excel file to get started.")
