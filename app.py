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
    "Upload any insurance Excel file and automatically generate a fixed insurer-ready output format."
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
        "insured amount",
        "coverage",
        "coverage amount",
        "calculation sa",
        "aviva calculation sa",
        "sa",
        "gtl"
    ],

    "Nominee Name": [
        "nominee"
    ],

    "Nominee Age": [
        "nominee age"
    ],

    "Loan Disbursement Date (DDMMYYYY)": [
        "loan start",
        "disbursement",
        "start date"
    ],

    "Loan End date (DDMMYYYY)": [
        "loan end",
        "loan maturity",
        "end date"
    ],

    "MAIN MEMBER AGE": [
        "age",
        "member age"
    ]

}

# =====================================================
# SMART COLUMN DETECTION
# =====================================================

def detect_column(columns, aliases):

    for alias in aliases:

        for col in columns:

            cleaned_col = str(col).lower().strip()

            if cleaned_col == alias:

                return col

            elif alias in cleaned_col:

                return col

    return None

# =====================================================
# CLEAN DATAFRAME
# =====================================================

def clean_dataframe(df):

    df.dropna(
        how='all',
        inplace=True
    )

    df.dropna(
        axis=1,
        how='all',
        inplace=True
    )

    total_keywords = [

        "total",
        "grand total",
        "subtotal",
        "summary"

    ]

    mask = pd.Series(
        [False] * len(df)
    )

    for col in df.columns:

        mask = mask | df[col].astype(str).str.lower().str.contains(

            "|".join(total_keywords),

            na=False

        )

    df = df[~mask]

    return df

# =====================================================
# DATE CLEANING
# =====================================================

def clean_date_column(series):

    try:

        cleaned = pd.to_datetime(
            series,
            errors='coerce'
        )

        return cleaned.dt.strftime('%d%b%Y')

    except:

        return series

# =====================================================
# MOBILE CLEANING
# =====================================================

def clean_mobile(series):

    cleaned = (

        series.astype(str)

        .str.replace(r'\D', '', regex=True)

        .str[-10:]

    )

    return cleaned

# =====================================================
# AADHAR CLEANING
# =====================================================

def clean_aadhar(series):

    cleaned = (

        series.astype(str)

        .str.replace(r'\D', '', regex=True)

        .str[-12:]

    )

    return cleaned

# =====================================================
# AGE CLEANING
# =====================================================

def clean_age(series):

    cleaned = pd.to_numeric(

        series,

        errors='coerce'

    ).fillna(0).astype(int)

    cleaned = cleaned.clip(
        lower=0,
        upper=99
    )

    return cleaned

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

            df = pd.read_excel(

                file,

                header=0

            )

            # =====================================================
            # CLEAN DATAFRAME
            # =====================================================

            df = clean_dataframe(df)

            # =====================================================
            # CLEAN COLUMNS
            # =====================================================

            df.columns = [

                str(col)
                .replace("\n", " ")
                .replace("_", " ")
                .strip()
                .lower()

                for col in df.columns

            ]

            # =====================================================
            # CREATE OUTPUT DATAFRAME
            # =====================================================

            standardized_df = pd.DataFrame()

            for col in MASTER_COLUMNS:

                standardized_df[col] = "NA"

            # =====================================================
            # AUTO MAPPING
            # =====================================================

            for standard_col, alias_list in ALIASES.items():

                detected_col = detect_column(

                    df.columns,

                    alias_list

                )

                if detected_col:

                    standardized_df[
                        standard_col
                    ] = df[detected_col]

            # =====================================================
            # CLEAN SA
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

            )

            # =====================================================
            # CLEAN LOAN AMOUNT
            # =====================================================

            standardized_df["Loan Outstanding Amount"] = (

                standardized_df["Loan Outstanding Amount"]

                .astype(str)

                .str.replace(",", "")

                .str.replace("₹", "")

                .str.strip()

            )

            standardized_df["Loan Outstanding Amount"] = pd.to_numeric(

                standardized_df["Loan Outstanding Amount"],

                errors="coerce"

            )

            # =====================================================
            # SMART SA LOGIC
            # =====================================================

            standardized_df["Final SA"] = np.where(

                standardized_df["Sum Assured"].notna()

                &

                (standardized_df["Sum Assured"] > 0),

                standardized_df["Sum Assured"],

                standardized_df["Loan Outstanding Amount"]

            )

            standardized_df["Final SA"] = (

                standardized_df["Final SA"]

                .fillna(0)

            )

            standardized_df["Sum Assured"] = (

                standardized_df["Final SA"]

            )

            # =====================================================
            # REMOVE INVALID ROWS
            # =====================================================

            standardized_df = standardized_df[

                standardized_df["Final SA"] > 0

            ]

            standardized_df = standardized_df[

                standardized_df[
                    "Name of Primary Loan borrower"
                ].astype(str).str.strip().ne("")
            ]

            # =====================================================
            # CLEAN DATES
            # =====================================================

            standardized_df["Date of Birth (DDMMMYYYY)"] = clean_date_column(

                standardized_df["Date of Birth (DDMMMYYYY)"]

            )

            standardized_df["Loan Disbursement Date (DDMMYYYY)"] = clean_date_column(

                standardized_df["Loan Disbursement Date (DDMMYYYY)"]

            )

            standardized_df["Loan End date (DDMMYYYY)"] = clean_date_column(

                standardized_df["Loan End date (DDMMYYYY)"]

            )

            # =====================================================
            # CLEAN MOBILE
            # =====================================================

            standardized_df["Mobile No"] = clean_mobile(

                standardized_df["Mobile No"]

            )

            # =====================================================
            # CLEAN AGE
            # =====================================================

            standardized_df["MAIN MEMBER AGE"] = clean_age(

                standardized_df["MAIN MEMBER AGE"]

            )

            # =====================================================
            # CALCULATE LOAN TERM
            # =====================================================

            start_date = pd.to_datetime(

                standardized_df[
                    "Loan Disbursement Date (DDMMYYYY)"
                ],

                errors='coerce'

            )

            end_date = pd.to_datetime(

                standardized_df[
                    "Loan End date (DDMMYYYY)"
                ],

                errors='coerce'

            )

            months = (

                (end_date.dt.year - start_date.dt.year) * 12

                +

                (end_date.dt.month - start_date.dt.month)

            )

            standardized_df["Loan Term (in months)"] = (

                months.fillna(0).astype(int)

            )

            standardized_df["Loan Term (Year)"] = (

                standardized_df[
                    "Loan Term (in months)"
                ] / 12

            ).round(1)

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
            # PREMIUM
            # =====================================================

            standardized_df["Premium (Excl. GST)"] = (

                standardized_df["Final SA"]

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

                standardized_df["Final SA"]

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
            # FINAL COLUMN ORDER
            # =====================================================

            standardized_df = standardized_df[MASTER_COLUMNS]

            # =====================================================
            # APPEND DATA
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

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:

        st.metric(
            "Total Members",
            len(final_master_df)
        )

    with col2:

        st.metric(
            "Total SA",
            f"₹ {final_master_df['Sum Assured'].sum():,.0f}"
        )

    with col3:

        st.metric(
            "Premium Excl GST",
            f"₹ {final_master_df['Premium (Excl. GST)'].sum():,.2f}"
        )

    with col4:

        st.metric(
            "Total GST",
            f"₹ {final_master_df['GST amount'].sum():,.2f}"
        )

    with col5:

        st.metric(
            "Total Premium",
            f"₹ {final_master_df['Total Premium (incl GST)'].sum():,.2f}"
        )

    # =====================================================
    # OUTPUT
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

        st.dataframe(
            pd.DataFrame(error_log)
        )

    # =====================================================
    # DOWNLOAD
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

    processed_data = output.getvalue()

    st.download_button(

        label="⬇ Download Final Excel",

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
