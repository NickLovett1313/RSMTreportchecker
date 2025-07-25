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
    unique_spartans = sorted(df['CONTACT_NM'].dropna().unique().tolist())

    st.markdown("### ✅ Step 1: Choose Spartans")
    with st.form("spartan_form"):
        selected_spartans = st.multiselect("Select Spartans to check:", options=unique_spartans)
        run = st.form_submit_button("🚀 Run Analysis")

    if run:
        summary_data = []
        for spartan in selected_spartans:
            sdf = df[df['CONTACT_NM'] == spartan]
            pos = sdf['PO'].dropna().unique().tolist()

            awaiting, tbd = [], []
            for po in pos:
                pod = sdf[sdf['PO'] == po]
                if (pod['LINE_STATUS'] == 'AWAITING_SHIPPING').any():
                    clean = str(int(float(po))) if str(po).replace('.0','').isdigit() else str(po)
                    awaiting.append(clean)
                    if (pod['SHIP_TO_CUSTOMER'] == 'TO BE DETERMINED').any():
                        tbd.append(clean)

            summary_data.append({
                'Spartan': spartan,
                'Awaiting Shipping POs': ", ".join(awaiting) or "None",
                'TBD Ship To POs': ", ".join(tbd) or "None"
            })

        summary_df = pd.DataFrame(summary_data)
        st.subheader("📊 Summary Table")
        st.dataframe(summary_df, use_container_width=True)
        st.session_state['summary_df'] = summary_df

    if 'summary_df' in st.session_state and st.button("📋 Ready to send to team?"):
        summary_df = st.session_state['summary_df']
        formatted_date = format_date_suffix(datetime.today())
        subject = f"Rosemount Orders – Daily Open Orders Report Review: {formatted_date}"

        # Subject in Arial
        st.markdown("### ✉️ Email Subject", unsafe_allow_html=True)
        st.markdown(f"<div style='font-family:Arial'>{subject}</div>", unsafe_allow_html=True)

        # Body HTML to match your PDF exactly
        body_html = """
<div style='font-family:Arial;'>
  <p>Hi Team,&nbsp;&nbsp;</p>
  <p>The Daily Open Orders Report for your Rosemount purchase orders has been reviewed for<br>
  those CC’d.&nbsp;</p>
  <p><strong>Note:</strong> for those PO#s with items awaiting shipment: If you haven’t yet received a packing<br>
  slip for release, I recommend reaching out to your factory contact.&nbsp;</p>
  <p><strong>Note:</strong> for those PO#s with a TBD ship-to address: This information must be provided to<br>
  the factory before they can issue a packing slip.&nbsp;</p>
  <p>See information below:&nbsp;</p>
  <hr style='border-top:1px dashed #333;'/>
  <ol style='margin-left:0; padding-left:20px;'>
"""
        # Add each Spartan
        for row in summary_df.itertuples(index=False):
            spartan, awaiting, tbd = row
            body_html += f"""
    <li>{spartan}
      <ul style='list-style-type:circle; margin:4px 0 4px 20px;'>
        <li>PO#s Awaiting Shipping: {awaiting}</li>
        <li>PO#s with TBD Ship-To Address: {tbd}</li>
      </ul>
    </li>
"""
        body_html += """
  </ol>
  <hr style='border-top:1px dashed #333;'/>
  <p>Thanks!</p>
</div>
"""

        st.markdown("### 📩 Email Body", unsafe_allow_html=True)
        st.markdown(body_html, unsafe_allow_html=True)

else:
    st.info("👆 Upload an Excel file to get started.")
