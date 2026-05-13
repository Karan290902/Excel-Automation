# =====================================================
# IMPORTS
# =====================================================

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from difflib import SequenceMatcher
import re

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Insurance AI Mapper",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Insurance AI Standardization Engine")

st.write(
    "Upload insurance excel files and generate insurer-ready output automatically."
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
# ALIASES
# =====================================================

ALIASES = {

    "Loan Account No.": [

        "loan account no",
        "loan account",
        "loan ac",
        "loan a/c",
        "a/c number",
        "account number",
        "account no",
        "membership no",
        "membership number",
        "member id",
        "certificate no",
        "certificate number"

    ],

    "Name of Primary Loan borrower": [

        "member name",
        "customer name",
        "borrower name",
        "insured name",
        "client name",
        "full name",
        "name"

    ],

    "Gender": [
        "gender",
        "sex"
    ],

    "Date of Birth (DDMMMYYYY)": [

        "dob",
        "date of birth",
        "birth date",
        "d.o.b"

    ],

    "MAIN MEMBER AGE": [
        "age",
        "member age",
        "insured age"
    ],

    "Pincode": [
        "pincode",
        "pin code",
        "postal code"
    ],

    "Mobile No": [

        "mobile number",
        "mobile no",
        "mobile",
        "mob no",
        "mob",
        "phone",
        "phone no",
        "contact no"

    ],

    "Nominee Name": [
        "nominee name",
        "nominee"
    ],

    "Relationship of the Nominee with Insurance covered Person": [

        "nominee relationship",
        "relationship",
        "relation"

    ],

    "Nominee Age": [
        "nominee age"
    ],

    "Loan Outstanding Amount": [

        "loan amount",
        "loan amt",
        "outstanding amount",
        "loan outstanding",
        "disbursed amount"

    ],

    "Sum Assured": [

        "sum assured",
        "sum insured",
        "coverage",
        "sa"

    ],

    "Loan Disbursement Date (DDMMYYYY)": [

        "loan start date",
        "loan disbursement date",
        "disbursement date",
        "loan start"

    ],

    "Loan End date (DDMMYYYY)": [

        "loan end date",
        "loan expiry date",
        "loan end"

    ],

    "Address            (First Life)": [

        "address",
        "residential address",
        "communication address",
        "society name"

    ],

    "Address 1            (First Life)": [

        "district",
        "place",
        "city",
        "location"

    ]

}

# =====================================================
# SMART COLUMN DETECTION
# =====================================================

def detect_column(columns, aliases):

    best_match = None
    best_score = 0

    for col in columns:

        clean_col = (

            str(col)

            .lower()

            .replace("_", " ")

            .replace("-", " ")

            .replace("/", " ")

            .replace(".", " ")

            .replace("\n", " ")

            .strip()

        )

        for alias in aliases:

            alias_clean = alias.lower().strip()

            if alias_clean in clean_col:

                score = 100

            elif any(
                word in clean_col
                for word in alias_clean.split()
            ):

                score = 85

            else:

                score = SequenceMatcher(
                    None,
                    clean_col,
                    alias_clean
                ).ratio() * 100

            if score > best_score:

                best_score = score
                best_match = col

    if best_score >= 40:
        return best_match

    return None

# =====================================================
# CLEAN FUNCTIONS
# =====================================================

def clean_money(series):

    cleaned = (

        series.astype(str)

        .str.replace(",", "", regex=False)

        .str.replace("₹", "", regex=False)

        .str.replace("/-", "", regex=False)

        .str.strip()

    )

    cleaned = cleaned.replace(
        ["", "nan", "None"],
        np.nan
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce"
    )

def clean_mobile(series):

    cleaned = (

        series.astype(str)

        .str.replace(r"\D", "", regex=True)

        .str[-10:]

    )

    cleaned = cleaned.where(
        cleaned.str.len() == 10
    )

    return cleaned

def clean_pincode(series):

    cleaned = (

        series.astype(str)

        .str.replace(r"\D", "", regex=True)

        .str[-6:]

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
        cleaned.between(0, 99)
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
        "summary",
        "subtotal"
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
            # READ RAW FILE
            # =====================================================

            raw_df = pd.read_excel(
                file,
                header=None
            )

            header_row = 0

            for i in range(min(15, len(raw_df))):

                row_text = " ".join(

                    raw_df.iloc[i]

                    .astype(str)

                    .fillna("")

                    .tolist()

                ).lower()

                if (
                    "member" in row_text
                    or "name" in row_text
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
            # CLEAN COLUMNS
            # =====================================================

            df.columns = [

                str(col)

                .replace("\n", " ")

                .replace("_", " ")

                .strip()

                for col in df.columns

            ]

            # =====================================================
            # CREATE OUTPUT DF
            # =====================================================

            standardized_df = pd.DataFrame()

            for col in MASTER_COLUMNS:
                standardized_df[col] = np.nan

            # =====================================================
            # AUTO MAPPING
            # =====================================================

            for output_col, aliases in ALIASES.items():

                detected_col = detect_column(
                    df.columns,
                    aliases
                )

                if (
                    detected_col is not None
                    and detected_col in df.columns
                ):

                    standardized_df[output_col] = df[detected_col]

            # =====================================================
            # LOAN TYPE
            # =====================================================

            standardized_df["Loan Type"] = loan_type

            # =====================================================
            # CLEAN MONEY
            # =====================================================

            standardized_df["Loan Outstanding Amount"] = clean_money(
                standardized_df["Loan Outstanding Amount"]
            )

            standardized_df["Sum Assured"] = clean_money(
                standardized_df["Sum Assured"]
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
            # CLEAN DATES
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
            # CLEAN MOBILE
            # =====================================================

            standardized_df["Mobile No"] = clean_mobile(
                standardized_df["Mobile No"]
            )

            # =====================================================
            # CLEAN ADDRESS
            # =====================================================

            standardized_df["Address            (First Life)"] = np.where(

                standardized_df[
                    "Address            (First Life)"
                ]

                .astype(str)

                .str.replace(r"\D", "", regex=True)

                .str.len()

                >= 10,

                np.nan,

                standardized_df[
                    "Address            (First Life)"
                ]

            )

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
            # CLEAN PINCODE
            # =====================================================

            standardized_df["Pincode"] = clean_pincode(
                standardized_df["Pincode"]
            )

            # =====================================================
            # CLEAN AGE
            # =====================================================

            standardized_df["MAIN MEMBER AGE"] = clean_age(
                standardized_df["MAIN MEMBER AGE"]
            )

            standardized_df["Nominee Age"] = clean_age(
                standardized_df["Nominee Age"]
            )

            # =====================================================
            # REMOVE SAME NOMINEE NAME
            # =====================================================

            standardized_df["Nominee Name"] = np.where(

                standardized_df["Nominee Name"]

                .astype(str)

                .str.lower()

                ==

                standardized_df["Name of Primary Loan borrower"]

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

                    mobile_col = detect_column(
                        df.columns,
                        ALIASES["Mobile No"]
                    )

                    if col != mobile_col:

                        standardized_df[
                            "Type of    Age Proof"
                        ] = cleaned

                        break

            # =====================================================
            # PREMIUM CALCULATIONS
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

    display_df = final_master_df.astype(str)

    st.dataframe(
        display_df,
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

st.markdown("---")

st.caption(
    "Built for Insurance Underwriting Automation"
)
