import streamlit as st
import pandas as pd

st.set_page_config(page_title="Awaiting Shipping Checker")
st.title("📦 Awaiting Shipping Checker")

uploaded_file = st.file_uploader("Upload the latest Excel sheet", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    unique_reps = df['CONTACT_NM'].dropna().unique().tolist()
    unique_reps.sort()

    # Let user choose reps from available names
    selected_reps = st.multiselect(
        "Select reps to check:",
        options=unique_reps
    )

    # Only run if the button is clicked
    if st.button("🚀 Run Analysis"):
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

        # Summary for the rep
            summary_data.append({
                'Rep Name': rep,
                'Awaiting Shipping POs': '\n'.join(awaiting_shipping_pos) if awaiting_shipping_pos else 'None',
                'TBD Ship To POs': '\n'.join(tbd_ship_to_pos) if tbd_ship_to_pos else 'None'
            })

        # Display the summary
        st.subheader("📊 Summary Table")
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True)

        # Download button
        csv = summary_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="⬇️ Download CSV",
            data=csv,
            file_name='awaiting_shipping_summary.csv',
            mime='text/csv'
        )
else:
    st.info("👆 Upload an Excel file to get started.")
