import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image
import os

# Page config
st.set_page_config(page_title="RSMT – Daily Open Orders Report Checker")

# Try to display Spartan logo + title side-by-side
logo_path = "/mnt/data/spartan logo.jpg"
col1, col2 = st.columns([1, 6])
with col1:
    if os.path.exists(logo_path):
        logo = Image.open(logo_path)
        st.image(logo, width=100)
    else:
        st.warning("⚠️ Spartan logo not found.")
with col2:
    st.title("📦 RSMT – Daily Open Orders Report Checker")

# App description and steps
st.markdown("""
This app checks the latest RSMT Open Orders report for selected Spartans and returns whether any of their assigned PO#s:
- Have items **awaiting shipment**, or
- Have a **TBD ship-to** address.

#### 👉 Four Easy Steps:
1. **Upload** the latest copy of the Excel sheet, located here:  
   [Open Rosemount Daily Report](https://spartancontrols.sharepoint.com/sites/collab/vms/Rosemount%20daily%20open%20order%20report/Forms/AllItems.aspx?ovuser=0123b3b0%2Ddfd2%2D4b73%2Dbce1%2D5761a6245688%2CLovett%2ENick%40spartancontrols%2Ecom&OR=Teams%2DHL&CT=1749048035545&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNTA1MDQwMTYyNCIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D)

2. **Select** the Spartans you want to check.

3. Click **Run Analysis** to generate the summary table.

4. Click **Generate email to send to team**, then copy and send it to the selected Spartans.

---

⚠️ *Disclaimer: No data is stored or uploaded to any external server. Everything is processed locally through Streamlit during your session.*
""")

def format_date_suffix(date_obj):
    day = int(date_obj.strftime("%d"))
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return date_obj.strftime(f"%B {day}{suffix}, %Y")

uploaded_file = st.file_uploader("Upload the latest Excel sheet", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    spartans = sorted(df["CONTACT_NM"].dropna().unique().tolist())

    st.markdown("### ✅ Step 2: Choose Spartans")
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

    if "summary_df" in st.session_state and st.button("📧 Generate email to send to team"):
        summary_df = st.session_state["summary_df"]
        date_str = format_date_suffix(datetime.today())
        subject = f"Rosemount Orders – Daily Open Orders Report Review: {date_str}"

        st.markdown("### ✉️ Email Subject")
        st.markdown(f"<div style='font-family:Arial; font-size:14px; color:#333'>{subject}</div>", unsafe_allow_html=True)

        # Build formatted HTML email body
        email_body = f"""
<div style="font-family:Arial,sans-serif; font-size:11pt; line-height:1.5; color:#000; background:#f9f9f9; padding:16px; border-radius:8px;">
<p>Hi Team,</p>

<p>The Daily Open Orders Report for your Rosemount purchase orders has been reviewed for those CC’d.</p>

<p style="margin-left:1.5em;">
  <b><i><span style="color:black">Note:</span></i></b><i><span style="color:black"> for those PO#s with items awaiting shipment</span><span style="color:#ED7D31">: If you haven’t yet received a packing slip for release, I recommend reaching out to your factory contact.</span></i>
</p>

<p style="margin-left:1.5em;">
  <b><i><span style="color:black">Note:</span></i></b><i><span style="color:black"> for those PO#s with a TBD ship-to address: </span><span style="color:#0070C0">This information must be provided to the factory before they can issue a packing slip.</span></i>
</p>

<p><b>See information below:</b></p>
<p style="margin-bottom:6px;">----------------------------</p>
"""

        for idx, row in enumerate(summary_df.itertuples(index=False), start=1):
            name, aw, tbd = row
            email_body += f"""
<p style="margin-bottom:0;"><b>{idx}. {name}</b></p>
<ul style="list-style:none; margin-top:0; margin-left:1.5em; padding-left:0;">
  <li><span style='color:#0070C0'>-</span> <span style='color:#ED7D31'>PO#s Awaiting Shipping: {aw}</span></li>
  <li><span style='color:#0070C0'>-</span> <span style='color:#0070C0'>PO#s with TBD Ship-To Address: {tbd}</span></li>
</ul>
"""

        email_body += """
<p style="margin-top:0;">----------------------------</p>
<p>Thanks!</p>
</div>
"""

        st.markdown("### 📩 Email Body")
        st.markdown(email_body, unsafe_allow_html=True)

else:
    st.info("👆 Upload an Excel file to get started.")
