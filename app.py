import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Insurance Data Standardization Engine",
    page_icon="📊",
    layout="wide"
)

# =====================================================
# TITLE
# =====================================================

st.title("📊 Insurance Data Standardization Engine")

st.write(
    "Upload multiple insurance Excel files with different formats and generate one fixed insurer-ready output file."
)

# =====================================================
# SETTINGS
# =====================================================

RATE_PER_LAKH = 320.3
GST_RATE = 0.18

# =====================================================
# FIXED OUTPUT FORMAT
# =====================================================

MASTER_COLUMNS = [

    "Sr. no.",
    "Zone",
    "Branch Name",
    "Branch Code",
    "Loan Type",
    "Loan Account No.",
    "Name of Primary Loan borrower",
    "Name of Coborrower(if applicable)",
    "Loan Amount Coborrower",
    "Gender",
    "Date of Birth (DDMMMYYYY)",
    "Type of    Age Proof",
    "Address            (First Life)",
    "Address 1            (First Life)",
    "Address 2            (First Life)",
    "Pincode",
    "Mobile No",
    "Email Id",
    "Nominee Name",
    "Relationship of the Nominee with Insurance covered Person",
    "Nominee DOB(DDMMMYYYY)*please mention name of the month Eg 10Feb1991",
    "Nominee Age",
    "Appointee Name",
    "Relationship with Borrower",
    "Appointee DOB(DDMMMYYYY)*please mention name of the month Eg 10Feb1991",
    "Loan Outstanding Amount",
    "Sum Assured",
    "Loan Disbursement Date (DDMMYYYY)",
    "Loan End date (DDMMYYYY)",
    "Loan Term (in months)",
    "Loan Term (Year)",
    "MAIN MEMBER AGE",
    "Rate",
    "Premium (Excl. GST)",
    "GST amount",
    "Total Premium (incl GST)",
    "Premium Transfer Amount",
    "Premium Transfer Date to Aviva",
    "Premium Transaction ID/JOURNAL NUMBER/UTR",
    "DGH / Membeship form Collected (Yes/No)",
    "Any adverse answer to DGH Questioniare (Yes/No)",
    "Date when Applicant Signed",
    "Height of Person covered",
    "Weight of Person covered",
    "Aviva Calculation SA",
    "Premium Excl. Gst",
    "GST",
    "Total Premium",
    "Aviva Remarks"

]

# =====================================================
# SMART COLUMN ALIASES
# =====================================================

ALIASES = {

    "Loan Account No.": [

        "loan account",
        "account no",
        "a/c number",
        "membership no",
        "loan no",
        "lan"

    ],

    "Name of Primary Loan borrower": [

        "member name",
        "customer name",
        "borrower name",
        "insured name",
        "primary borrower",
        "name"

    ],

    "Gender": [

        "gender",
        "sex"

    ],

    "Date of Birth (DDMMMYYYY)": [

        "dob",
        "date of birth",
        "birth date"

    ],

    "Mobile No": [

        "mobile",
        "mobile number",
        "phone",
        "contact"

    ],

    "Address            (First Life)": [

        "address",
        "residence"

    ],

    "Pincode": [

        "pincode",
        "pin code",
        "zip"

    ],

    "Branch Name": [

        "branch",
        "branch name"

    ],

    "Loan Outstanding Amount": [

        "loan amount",
        "outstanding amount",
        "loan outstanding"

    ],

    "Sum Assured": [

        "sum assured",
        "sum insured",
        "coverage",
        "insured amount",
        "calculation sa",
        "aviva calculation sa",
        "sa"

    ],

    "Nominee Name": [

        "nominee"

    ],

    "Nominee Age": [

        "nominee age"

    ],

    "Loan Disbursement Date (DDMMYYYY)": [

        "loan start",
        "disbursement"

    ],

    "Loan End date (DDMMYYYY)": [

        "loan end",
        "loan maturity"

    ],

    "MAIN MEMBER AGE": [

        "age",
        "member age"

    ]

}

# =====================================================
# COLUMN DETECTION FUNCTION
# =====================================================

def detect_column(columns, aliases):

    for col in columns:

        cleaned_col = str(col).lower().strip()

        cleaned_col = cleaned_col.replace("_", " ")

        for alias in aliases:

            if alias in cleaned_col:

                return col

    return None

# =====================================================
# FILE UPLOAD
# =====================================================

uploaded_files = st.file_uploader(

    "📂 Upload Insurance Excel Files",

    type=["xlsx"],

    accept_multiple_files=True

)

# =====================================================
# PROCESS FILES
# =====================================================

if uploaded_files:

    final_master_df = pd.DataFrame()

    error_log = []

    for file in uploaded_files:

        try:

            # =====================================================
            # READ FILE
            # =====================================================

            df = pd.read_excel(file)

            df = df.copy()

            # REMOVE EMPTY ROWS

            df.dropna(
                how='all',
                inplace=True
            )

            # CLEAN COLUMN NAMES

            df.columns = df.columns.astype(str)

            df.columns = df.columns.str.strip()

            # =====================================================
            # CREATE STANDARDIZED DATAFRAME
            # =====================================================

            standardized_df = pd.DataFrame()

            for col in MASTER_COLUMNS:

                standardized_df[col] = "NA"

            # =====================================================
            # AUTO COLUMN MAPPING
            # =====================================================

            detected_mapping = {}

            for standard_col, alias_list in ALIASES.items():

                detected_col = detect_column(

                    df.columns,

                    alias_list

                )

                if detected_col:

                    standardized_df[
                        standard_col
                    ] = df[detected_col]

                    detected_mapping[
                        standard_col
                    ] = detected_col

            # =====================================================
            # SHOW DETECTED MAPPING
            # =====================================================

            with st.expander(f"🧠 Mapping - {file.name}"):

                st.write(detected_mapping)

            # =====================================================
            # CLEAN SUM ASSURED
            # =====================================================

            standardized_df["Sum Assured"] = (

                standardized_df["Sum Assured"]

                .astype(str)

                .str.replace(",", "")

                .str.replace("₹", "")

                .str.strip()

            )

            standardized_df["Sum Assured"] = pd.to_numeric(

                standardized_df["Sum Assured"],

                errors="coerce"

            ).fillna(0)

            # =====================================================
            # REMOVE INVALID ROWS
            # =====================================================

            standardized_df = standardized_df[

                standardized_df["Sum Assured"] > 0

            ]

            # =====================================================
            # SERIAL NUMBER
            # =====================================================

            standardized_df["Sr. no."] = range(

                1,

                len(standardized_df) + 1

            )

            # =====================================================
            # RATE
            # =====================================================

            standardized_df["Rate"] = RATE_PER_LAKH

            # =====================================================
            # PREMIUM CALCULATION
            # =====================================================

            standardized_df["Premium (Excl. GST)"] = (

                standardized_df["Sum Assured"]

                / 100000

            ) * RATE_PER_LAKH

            # =====================================================
            # GST
            # =====================================================

            standardized_df["GST amount"] = (

                standardized_df["Premium (Excl. GST)"]

                * GST_RATE

            )

            # =====================================================
            # TOTAL PREMIUM
            # =====================================================

            standardized_df["Total Premium (incl GST)"] = (

                standardized_df["Premium (Excl. GST)"]

                + standardized_df["GST amount"]

            )

            # =====================================================
            # DUPLICATE AVIVA FIELDS
            # =====================================================

            standardized_df["Aviva Calculation SA"] = (

                standardized_df["Sum Assured"]

            )

            standardized_df["Premium Excl. Gst"] = (

                standardized_df["Premium (Excl. GST)"]

            )

            standardized_df["GST"] = (

                standardized_df["GST amount"]

            )

            standardized_df["Total Premium"] = (

                standardized_df["Total Premium (incl GST)"]

            )

            # =====================================================
            # REMARKS
            # =====================================================

            standardized_df["Aviva Remarks"] = (

                "Processed Successfully"

            )

            # =====================================================
            # APPEND FINAL DATA
            # =====================================================

            final_master_df = pd.concat(

                [final_master_df, standardized_df],

                ignore_index=True

            )

        except Exception as e:

            error_log.append({

                "File": file.name,

                "Error": str(e)

            })

    # =====================================================
    # DASHBOARD
    # =====================================================

    st.subheader("📊 Portfolio Summary")

    total_members = len(final_master_df)

    total_sa = final_master_df["Sum Assured"].sum()

    total_premium = final_master_df["Premium (Excl. GST)"].sum()

    total_gst = final_master_df["GST amount"].sum()

    total_total = final_master_df["Total Premium (incl GST)"].sum()

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:

        st.metric(
            "Total Members",
            total_members
        )

    with col2:

        st.metric(
            "Total SA",
            f"₹ {total_sa:,.0f}"
        )

    with col3:

        st.metric(
            "Premium Excl GST",
            f"₹ {total_premium:,.2f}"
        )

    with col4:

        st.metric(
            "Total GST",
            f"₹ {total_gst:,.2f}"
        )

    with col5:

        st.metric(
            "Total Premium",
            f"₹ {total_total:,.2f}"
        )

    # =====================================================
    # OUTPUT TABLE
    # =====================================================

    st.subheader("📋 Final Standardized Output")

    st.dataframe(
        final_master_df,
        use_container_width=True
    )

    # =====================================================
    # ERROR REPORT
    # =====================================================

    if len(error_log) > 0:

        st.subheader("⚠ Error Report")

        error_df = pd.DataFrame(error_log)

        st.dataframe(error_df)

    # =====================================================
    # DOWNLOAD EXCEL
    # =====================================================

    output = BytesIO()

    with pd.ExcelWriter(

        output,

        engine='openpyxl'

    ) as writer:

        final_master_df.to_excel(

            writer,

            index=False,

            sheet_name='Final Output'

        )

        if len(error_log) > 0:

            error_df.to_excel(

                writer,

                index=False,

                sheet_name='Errors'

            )

    processed_data = output.getvalue()

    st.download_button(

        label="⬇ Download Final Standardized Excel",

        data=processed_data,

        file_name="Final_Aviva_Output.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )

# =====================================================
# FOOTER
# =====================================================

st.markdown("---")

st.caption(
    "Built for Insurance Underwriting & Data Standardization"
)
