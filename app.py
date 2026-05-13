# =====================================================
# INSURANCE AI STANDARDIZATION ENGINE
# STRICT RULE-BASED VERSION
# =====================================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Insurance Mapper",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Insurance Data Standardization Engine")

st.write(
    "Upload insurance excel files and generate insurer-ready output."
)

# =====================================================
# SETTINGS
# =====================================================

RATE_PER_LAKH = 320.3
GST_RATE = 0.18

# =====================================================
# OUTPUT FORMAT
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
# STRICT COLUMN RULES
# =====================================================

COLUMN_RULES = {

    "Loan Account No.": [
        "a/c number",
        "account number",
        "account no",
        "loan account no",
        "membership no"
    ],

    "Name of Primary Loan borrower": [
        "member name"
    ],

    "Gender": [
        "gender"
    ],

    "Date of Birth (DDMMMYYYY)": [
        "dob",
        "date of birth"
    ],

    "MAIN MEMBER AGE": [
        "age"
    ],

    "Mobile No": [
        "mobile number",
        "mobile no",
        "mobile"
    ],

    "Pincode": [
        "pincode",
        "pin code"
    ],

    "Nominee Name": [
        "nominee name"
    ],

    "Relationship of the Nominee with Insurance covered Person": [
        "nominee relationship"
    ],

    "Nominee Age": [
        "nominee age"
    ],

    "Loan Outstanding Amount": [
        "loan amount"
    ],

    "Sum Assured": [
        "sum assured",
        "sum insured"
    ],

    "Loan Disbursement Date (DDMMYYYY)": [
        "loan start date"
    ],

    "Loan End date (DDMMYYYY)": [
        "loan end date"
    ],

    "Address            (First Life)": [
        "address",
        "society name"
    ],

    "Address 1            (First Life)": [
        "district",
        "place"
    ]

}

# =====================================================
# FIND COLUMN
# =====================================================

def find_column(columns, keywords):

    for col in columns:

        clean_col = str(col).lower().strip()

        for keyword in keywords:

            if keyword in clean_col:

                return col

    return None

# =====================================================
# CLEAN FUNCTIONS
# =====================================================

def clean_money(series):

    series = (

        series.astype(str)

        .str.replace(",", "", regex=False)

        .str.replace("₹", "", regex=False)

        .str.strip()

    )

    return pd.to_numeric(
        series,
        errors="coerce"
    )

def clean_mobile(series):

    cleaned = (

        series.astype(str)

        .str.replace(r"\D", "", regex=True)

    )

    cleaned = cleaned.where(
        cleaned.str.len() == 10
    )

    return cleaned

def clean_pincode(series):

    cleaned = (

        series.astype(str)

        .str.replace(r"\D", "", regex=True)

    )

    cleaned = cleaned.where(
        cleaned.str.len() == 6
    )

    return cleaned

def clean_age(series):

    cleaned = pd.to_numeric(
        series,
        errors="coerce"
    )

    cleaned = cleaned.where(
        cleaned.between(1, 99)
    )

    return cleaned

def clean_date(series):

    cleaned = pd.to_datetime(
        series,
        errors="coerce"
    )

    return cleaned.dt.strftime("%d%b%Y")

# =====================================================
# REMOVE TOTAL ROWS
# =====================================================

def remove_total_rows(df):

    keywords = [
        "total",
        "grand total",
        "subtotal",
        "summary"
    ]

    mask = pd.Series(
        [False] * len(df),
        index=df.index
    )

    for col in df.columns:

        mask = mask | (

            df[col]

            .astype(str)

            .str.lower()

            .str.contains(
                "|".join(keywords),
                na=False
            )

        )

    return df[~mask]

# =====================================================
# FILE UPLOAD
# =====================================================

uploaded_files = st.file_uploader(
    "📂 Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# =====================================================
# PROCESS FILES
# =====================================================

if uploaded_files:

    final_master_df = pd.DataFrame()

    loan_type = st.selectbox(
        "Select Loan Type",
        ["GTL", "PA", "GCL"]
    )

    for file in uploaded_files:

        st.markdown("---")
        st.subheader(f"📄 {file.name}")

        try:

            # =====================================================
            # READ RAW
            # =====================================================

            raw_df = pd.read_excel(
                file,
                header=None
            )

            header_row = 0

            for i in range(min(10, len(raw_df))):

                row_text = " ".join(

                    raw_df.iloc[i]

                    .astype(str)

                    .fillna("")

                    .tolist()

                ).lower()

                if (
                    "member" in row_text
                    or "account" in row_text
                    or "loan" in row_text
                ):

                    header_row = i
                    break

            # =====================================================
            # READ ACTUAL FILE
            # =====================================================

            df = pd.read_excel(
                file,
                header=header_row
            )

            df.dropna(
                how="all",
                inplace=True
            )

            df.dropna(
                axis=1,
                how="all",
                inplace=True
            )

            df = remove_total_rows(df)

            # =====================================================
            # CLEAN COLUMN NAMES
            # =====================================================

            df.columns = [

                str(col)

                .replace("\n", " ")

                .replace("_", " ")

                .strip()

                for col in df.columns

            ]

            # =====================================================
            # OUTPUT DF
            # =====================================================

            standardized_df = pd.DataFrame()

            for col in MASTER_COLUMNS:
                standardized_df[col] = np.nan

            # =====================================================
            # STRICT MAPPING
            # =====================================================

            for output_col, keywords in COLUMN_RULES.items():

                matched_col = find_column(
                    df.columns,
                    keywords
                )

                if matched_col is not None:

                    standardized_df[output_col] = df[matched_col]

            # =====================================================
            # LOAN TYPE
            # =====================================================

            standardized_df["Loan Type"] = loan_type

            # =====================================================
            # CLEAN IMPORTANT FIELDS
            # =====================================================

            standardized_df["Mobile No"] = clean_mobile(
                standardized_df["Mobile No"]
            )

            standardized_df["Pincode"] = clean_pincode(
                standardized_df["Pincode"]
            )

            standardized_df["MAIN MEMBER AGE"] = clean_age(
                standardized_df["MAIN MEMBER AGE"]
            )

            standardized_df["Nominee Age"] = clean_age(
                standardized_df["Nominee Age"]
            )

            standardized_df["Loan Outstanding Amount"] = clean_money(
                standardized_df["Loan Outstanding Amount"]
            )

            standardized_df["Sum Assured"] = clean_money(
                standardized_df["Sum Assured"]
            )

            # =====================================================
            # ADDRESS CLEANING
            # =====================================================

            standardized_df["Address            (First Life)"] = np.where(

                standardized_df[
                    "Address            (First Life)"
                ]

                .astype(str)

                .str.contains(r"[A-Za-z]", regex=True),

                standardized_df[
                    "Address            (First Life)"
                ],

                np.nan

            )

            # =====================================================
            # DOB
            # =====================================================

            standardized_df["Date of Birth (DDMMMYYYY)"] = clean_date(
                standardized_df["Date of Birth (DDMMMYYYY)"]
            )

            standardized_df["Loan Disbursement Date (DDMMYYYY)"] = clean_date(
                standardized_df["Loan Disbursement Date (DDMMYYYY)"]
            )

            standardized_df["Loan End date (DDMMYYYY)"] = clean_date(
                standardized_df["Loan End date (DDMMYYYY)"]
            )

            # =====================================================
            # REMOVE DOB FROM LOAN DATES
            # =====================================================

            standardized_df["Loan Disbursement Date (DDMMYYYY)"] = np.where(

                standardized_df["Loan Disbursement Date (DDMMYYYY)"]

                ==

                standardized_df["Date of Birth (DDMMMYYYY)"],

                np.nan,

                standardized_df["Loan Disbursement Date (DDMMYYYY)"]

            )

            standardized_df["Loan End date (DDMMYYYY)"] = np.where(

                standardized_df["Loan End date (DDMMYYYY)"]

                ==

                standardized_df["Date of Birth (DDMMMYYYY)"],

                np.nan,

                standardized_df["Loan End date (DDMMYYYY)"]

            )

            # =====================================================
            # SUM ASSURED LOGIC
            # =====================================================

            standardized_df["Final SA"] = np.where(

                standardized_df["Sum Assured"].notna(),

                standardized_df["Sum Assured"],

                standardized_df["Loan Outstanding Amount"]

            )

            standardized_df["Sum Assured"] = standardized_df["Final SA"]

            standardized_df["Aviva Calculation SA"] = standardized_df["Final SA"]

            # =====================================================
            # LOAN TERM
            # =====================================================

            start_date = pd.to_datetime(
                standardized_df["Loan Disbursement Date (DDMMYYYY)"],
                errors="coerce"
            )

            end_date = pd.to_datetime(
                standardized_df["Loan End date (DDMMYYYY)"],
                errors="coerce"
            )

            months = np.where(

                start_date.notna() & end_date.notna(),

                (
                    (end_date.dt.year - start_date.dt.year) * 12
                    +
                    (end_date.dt.month - start_date.dt.month)
                ),

                np.nan

            )

            standardized_df["Loan Term (in months)"] = months

            standardized_df["Loan Term (Year)"] = (
                standardized_df["Loan Term (in months)"] / 12
            ).round(1)

            # =====================================================
            # REMOVE SAME NOMINEE
            # =====================================================

            standardized_df["Nominee Name"] = np.where(

                standardized_df["Nominee Name"]

                .astype(str)

                .str.lower()

                ==

                standardized_df[
                    "Name of Primary Loan borrower"
                ]

                .astype(str)

                .str.lower(),

                np.nan,

                standardized_df["Nominee Name"]

            )

            # =====================================================
            # AGE PROOF DETECTION
            # =====================================================

            for col in df.columns:

                cleaned = (

                    df[col]

                    .astype(str)

                    .str.replace(r"\D", "", regex=True)

                )

                valid_mask = cleaned.str.len() == 10

                if valid_mask.any():

                    if col != find_column(
                        df.columns,
                        COLUMN_RULES["Mobile No"]
                    ):

                        standardized_df[
                            "Type of    Age Proof"
                        ] = cleaned

                        break

            # =====================================================
            # PREMIUM CALCULATION
            # =====================================================

            standardized_df["Rate"] = RATE_PER_LAKH

            standardized_df["Premium (Excl. GST)"] = (
                standardized_df["Final SA"] / 100000
            ) * RATE_PER_LAKH

            standardized_df["GST amount"] = (
                standardized_df["Premium (Excl. GST)"] * GST_RATE
            )

            standardized_df["Total Premium (incl GST)"] = (
                standardized_df["Premium (Excl. GST)"]
                +
                standardized_df["GST amount"]
            )

            standardized_df["Premium Excl. Gst"] = standardized_df[
                "Premium (Excl. GST)"
            ]

            standardized_df["GST"] = standardized_df[
                "GST amount"
            ]

            standardized_df["Total Premium"] = standardized_df[
                "Total Premium (incl GST)"
            ]

            # =====================================================
            # SERIAL NUMBER
            # =====================================================

            standardized_df["Sr. no."] = range(
                1,
                len(standardized_df) + 1
            )

            # =====================================================
            # FINAL FORMAT
            # =====================================================

            standardized_df = standardized_df[MASTER_COLUMNS]

            final_master_df = pd.concat(
                [final_master_df, standardized_df],
                ignore_index=True
            )

        except Exception as e:

            st.error(
                f"Error processing {file.name}: {str(e)}"
            )

    # =====================================================
    # FINAL CLEANING
    # =====================================================

    final_master_df = final_master_df.replace(
        [np.inf, -np.inf],
        np.nan
    )

    final_master_df = final_master_df.fillna("")

    # =====================================================
    # DISPLAY
    # =====================================================

    st.markdown("---")

    st.subheader("📋 Final Output")

    st.dataframe(
        final_master_df.astype(str),
        width="stretch"
    )

    # =====================================================
    # DOWNLOAD
    # =====================================================

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        final_master_df.to_excel(
            writer,
            index=False,
            sheet_name="Final Output"
        )

    st.download_button(
        label="⬇ Download Final Excel",
        data=output.getvalue(),
        file_name="Final_Aviva_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
