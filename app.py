import streamlit as st
import pandas as pd
import os

# ---------------------------
# File path
# ---------------------------
FILE_PATH = "company_data.xlsx"

# ---------------------------
# Load Excel data
# ---------------------------
if os.path.exists(FILE_PATH):
    df = pd.read_excel(FILE_PATH)
    df.columns = df.columns.str.strip().str.replace(".", "", regex=False).str.replace(" ", "", regex=False).str.lower()
else:
    df = pd.DataFrame(columns=[
        "srno", "companyname", "csmname", "externalinternal",
        "contactperson", "mailid", "contactnumber",
        "supportmailid", "supportcontactnumber"
    ])

# ---------------------------
# Streamlit App Layout
# ---------------------------
st.set_page_config(page_title="Company Dashboard", layout="wide")

st.markdown(
    """
    <h1 style='text-align:center; color:#0078D4;'>🏢 Company Management Dashboard</h1>
    <p style='text-align:center; color:gray;'>Search, edit, and manage your company database easily</p>
    """,
    unsafe_allow_html=True
)
st.divider()

# Tabs
tab1, tab2, tab3 = st.tabs(["🔍 Search Company", "✏️ Edit Records", "➕ Add New Company"])

# ---------------------------
# SEARCH TAB
# ---------------------------
with tab1:
    st.subheader("🔎 Search Company Details")

    if "companyname" not in df.columns:
        st.error("❌ 'Company Name' column not found in Excel.")
    else:
        col1, col2, col3 = st.columns(3)

        # Dropdown filters
        with col1:
            company_filter = st.selectbox("🏢 Company Name", [""] + sorted(df["companyname"].dropna().unique().tolist()))
        with col2:
            csm_filter = st.selectbox("🎯 CSM Name", [""] + sorted(df["csmname"].dropna().unique().tolist()))
        with col3:
            contact_filter = st.selectbox("👤 Contact Person", [""] + sorted(df["contactperson"].dropna().unique().tolist()))

        # Filter logic
        filtered_df = df.copy()
        if company_filter:
            filtered_df = filtered_df[filtered_df["companyname"].str.contains(company_filter, case=False, na=False)]
        if csm_filter:
            filtered_df = filtered_df[filtered_df["csmname"].str.contains(csm_filter, case=False, na=False)]
        if contact_filter:
            filtered_df = filtered_df[filtered_df["contactperson"].str.contains(contact_filter, case=False, na=False)]

        st.markdown("---")

        if not filtered_df.empty:
            st.success(f"✅ Found {len(filtered_df)} result(s)")
            for _, row in filtered_df.iterrows():
                st.markdown(
                    f"""
                    <div style='border:1px solid #ddd; border-radius:10px; padding:15px; margin:10px 0;
                                background:#f9f9ff; box-shadow:2px 2px 5px rgba(0,0,0,0.05);'>
                        <h3 style='color:#0078D4;'>{row.get('companyname','')}</h3>
                        <p><b>🎯 CSM Name:</b> {row.get('csmname','')}</p>
                        <p><b>👤 Contact Person:</b> {row.get('contactperson','')}</p>
                        <p><b>📧 Mail ID:</b> {row.get('mailid','')}</p>
                        <p><b>📞 Contact Number:</b> {row.get('contactnumber','')}</p>
                        <p><b>📧 Support Mail:</b> {row.get('supportmailid','')}</p>
                        <p><b>☎️ Support Contact:</b> {row.get('supportcontactnumber','')}</p>
                        <p><b>🔖 External/Internal:</b> {row.get('externalinternal','')}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        else:
            st.warning("No matching records found.")

# ---------------------------
# EDIT TAB
# ---------------------------
with tab2:
    st.subheader("📝 Edit Existing Records")

    st.dataframe(df, use_container_width=True)
    if "companyname" in df.columns:
        company_to_edit = st.selectbox("Choose a company to edit", [""] + df["companyname"].dropna().tolist())

        if company_to_edit:
            row = df[df["companyname"] == company_to_edit].iloc[0]

            with st.form("edit_form", clear_on_submit=True):
                cols = st.columns(2)
                with cols[0]:
                    s_no = st.text_input("Sr.No", row.get("srno", ""))
                    company_name = st.text_input("Company Name", row.get("companyname", ""))
                    csm_name = st.text_input("CSM Name", row.get("csmname", ""))
                    ext_int = st.text_input("External/Internal", row.get("externalinternal", ""))
                with cols[1]:
                    contact_person = st.text_input("Contact Person", row.get("contactperson", ""))
                    mail_id = st.text_input("Mail ID", row.get("mailid", ""))
                    contact_number = st.text_input("Contact Number", row.get("contactnumber", ""))
                    support_mail = st.text_input("Support Mail ID", row.get("supportmailid", ""))
                    support_contact = st.text_input("Support Contact Number", row.get("supportcontactnumber", ""))

                submitted = st.form_submit_button("💾 Save Changes")

                if submitted:
                    df.loc[df["companyname"] == company_to_edit, :] = [
                        s_no, company_name, csm_name, ext_int, contact_person,
                        mail_id, contact_number, support_mail, support_contact
                    ]
                    df.columns = df.columns.str.strip().str.replace(".", "", regex=False).str.replace(" ", "", regex=False).str.lower()
                    df.to_excel(FILE_PATH, index=False)
                    st.success("✅ Record updated successfully! Please refresh to see changes.")
    else:
        st.error("❌ 'Company Name' column not found in Excel.")

# ---------------------------
# ADD NEW TAB
# ---------------------------
with tab3:
    st.subheader("➕ Add a New Company")

    with st.form("add_form", clear_on_submit=True):
        cols = st.columns(2)
        with cols[0]:
            s_no = st.text_input("Sr.No")
            company_name = st.text_input("Company Name")
            csm_name = st.text_input("CSM Name")
            ext_int = st.text_input("External/Internal")
        with cols[1]:
            contact_person = st.text_input("Contact Person")
            mail_id = st.text_input("Mail ID")
            contact_number = st.text_input("Contact Number")
            support_mail = st.text_input("Support Mail ID")
            support_contact = st.text_input("Support Contact Number")

        submitted = st.form_submit_button("✅ Add Record")

        if submitted:
            if not company_name.strip():
                st.error("❌ Company Name is required!")
            else:
                new_record = pd.DataFrame([{
                    "srno": s_no,
                    "companyname": company_name,
                    "csmname": csm_name,
                    "externalinternal": ext_int,
                    "contactperson": contact_person,
                    "mailid": mail_id,
                    "contactnumber": contact_number,
                    "supportmailid": support_mail,
                    "supportcontactnumber": support_contact
                }])
                df = pd.concat([df, new_record], ignore_index=True)
                df.columns = df.columns.str.strip().str.replace(".", "", regex=False).str.replace(" ", "", regex=False).str.lower()
                df.to_excel(FILE_PATH, index=False)
                st.success(f"✅ '{company_name}' added successfully!")
