import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Awaiting Shipping Checker")
st.title("📦 Awaiting Shipping Checker")

def format_date_suffix(date_obj):
    day = int(date_obj.strftime("%d"))
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return date_obj.strftime(f"%B {day}{suffix}, %Y")

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
                    clean = str(int(float(po))) if str(po).replace(".0","").isdigit() else str(po)
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

        # Subject
        subject = f"Rosemount Orders – Daily Open Orders Report Review: {date_str}"
        st.markdown("### ✉️ Email Subject")
        st.markdown(f"<div style='font-family:Arial; font-size:14px; color:#333'>{subject}</div>", unsafe_allow_html=True)

        # Styled HTML Email Body
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

<p>----------------------------</p>
"""

        for idx, row in enumerate(summary_df.itertuples(index=False), start=1):
            name, aw, tbd = row
            email_body += f"""
<p style="margin-left:0.25in;"><b>{idx}. {name}</b></p>
<ul style="list-style-type:none; margin-top:0; margin-bottom:1em; margin-left:1.5in; padding-left:0;">
  <li><span style="font-family:'Courier New'; color:#ED7D31;">o</span> <span style="color:#ED7D31;">PO#s Awaiting Shipping: {aw}</span></li>
  <li><span style="font-family:'Courier New'; color:#0070C0;">o</span> <span style="color:#0070C0;">PO#s with TBD Ship-To Address: {tbd}</span></li>
</ul>
"""

        email_body += """
<p>----------------------------</p>
<p>Thanks!</p>
</div>
"""
        st.markdown("### 📩 Email Body")
        st.markdown(email_body, unsafe_allow_html=True)

else:
    st.info("👆 Upload an Excel file to get started.")
