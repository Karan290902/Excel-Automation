import streamlit as st
import pandas as pd
import numpy as np
import json
import os
from io import BytesIO
from rapidfuzz import fuzz

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="AI Insurance Mapper",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AI Insurance Data Standardization Engine")

st.write(
    "Upload any insurance excel file and auto-convert into insurer-ready format."
)

# =====================================================
# SETTINGS
# =====================================================

RATE_PER_LAKH = 320.3
GST_RATE = 0.18
MAPPING_FILE = "mapping_memory.json"

# =====================================================
# LOAD MAPPING MEMORY
# =====================================================

if os.path.exists(MAPPING_FILE):

    with open(MAPPING_FILE, "r") as f:

        SAVED_MAPPINGS = json.load(f)

else:

    SAVED_MAPPINGS = {}

# =====================================================
# SAVE MAPPING MEMORY
# =====================================================

def save_mapping_memory():

    with open(MAPPING_FILE, "w") as f:

        json.dump(
            SAVED_MAPPINGS,
            f,
            indent=4
        )

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
        "a/c number",
        "account number",
        "account no",
        "loan account",
        "membership no",
        "loan no",
        "lan"
    ],

    "Name of Primary Loan borrower": [
        "member name",
        "customer name",
        "borrower name",
        "insured name",
        "name"
    ],

    "Gender": [
        "gender",
        "sex"
    ],

    "Date of Birth (DDMMMYYYY)": [
        "dob",
        "date of birth"
    ],

    "MAIN MEMBER AGE": [
        "age"
    ],

    "Mobile No": [
        "mobile",
        "mobile number",
        "phone"
    ],

    "Pincode": [
        "pin code",
        "pincode"
    ],

    "Loan Outstanding Amount": [
        "loan amount",
        "loan outstanding"
    ],

    "Sum Assured": [
        "sum assured",
        "sum insured",
        "coverage",
        "gtl"
    ],

    "Nominee Name": [
        "nominee name",
        "nominee"
    ],

    "Relationship of the Nominee with Insurance covered Person": [
        "nominee relationship",
        "relation"
    ],

    "Nominee Age": [
        "nominee age"
    ],

    "Loan Disbursement Date (DDMMYYYY)": [
        "loan start date",
        "disbursement"
    ],

    "Loan End date (DDMMYYYY)": [
        "loan end date",
        "loan end"
    ]

}

# =====================================================
# SMART MATCHING
# =====================================================

def detect_column(columns, aliases):

    best_match = None
    best_score = 0

    for col in columns:

        clean_col = str(col).lower().strip()

        # CHECK SAVED MEMORY

        if clean_col in SAVED_MAPPINGS:

            return SAVED_MAPPINGS[clean_col]

        for alias in aliases:

            score = fuzz.token_sort_ratio(
                clean_col,
                alias
            )

            if score > best_score:

                best_score = score
                best_match = col

    if best_score >= 75:

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

        .str.strip()

    )

    cleaned = cleaned.replace(
        ["", "nan"],
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

    cleaned = cleaned.replace(
        ["", "nan"],
        np.nan
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
        "subtotal",
        "summary"
    ]

    mask = pd.Series(
        [False] * len(df)
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

    type=["xlsx"],

    accept_multiple_files=True

)

# =====================================================
# PROCESS FILES
# =====================================================

if uploaded_files:

    final_master_df = pd.DataFrame()

    for file in uploaded_files:

        st.markdown("---")

        st.subheader(f"📄 {file.name}")

        # =====================================================
        # READ RAW FILE
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

                .str.lower()

                .tolist()

            )

            if (

                "name" in row_text

                or

                "member" in row_text

                or

                "account" in row_text

            ):

                header_row = i
                break

        # =====================================================
        # READ AGAIN
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
        # OUTPUT DF
        # =====================================================

        standardized_df = pd.DataFrame()

        for col in MASTER_COLUMNS:

            standardized_df[col] = np.nan

        # =====================================================
        # AUTO MAPPING
        # =====================================================

        st.write("### 🤖 AI Auto Mapping")

        for output_col, aliases in ALIASES.items():

            detected_col = detect_column(
                df.columns,
                aliases
            )

            # LOW CONFIDENCE USER CONFIRMATION

            if detected_col is None:

                selected = st.selectbox(

                    f"Select column for ➜ {output_col}",

                    [""] + list(df.columns),

                    key=f"{file.name}_{output_col}"

                )

                if selected != "":

                    detected_col = selected

            if detected_col is not None:

                standardized_df[output_col] = df[detected_col]

                # SAVE MEMORY

                SAVED_MAPPINGS[
                    str(detected_col).lower()
                ] = output_col

        # SAVE MEMORY FILE

        save_mapping_memory()

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
        # SMART SA
        # =====================================================

        standardized_df["Final SA"] = np.where(

            standardized_df["Sum Assured"].notna(),

            standardized_df["Sum Assured"],

            standardized_df["Loan Outstanding Amount"]

        )

        standardized_df["Sum Assured"] = standardized_df["Final SA"]

        # =====================================================
        # CLEAN DOB
        # =====================================================

        standardized_df["Date of Birth (DDMMMYYYY)"] = clean_date(

            standardized_df["Date of Birth (DDMMMYYYY)"]

        )

        # =====================================================
        # VALIDATE LOAN DATES
        # =====================================================

        raw_start = standardized_df[
            "Loan Disbursement Date (DDMMYYYY)"
        ]

        raw_end = standardized_df[
            "Loan End date (DDMMYYYY)"
        ]

        raw_dob = standardized_df[
            "Date of Birth (DDMMMYYYY)"
        ]

        clean_start = clean_date(raw_start)
        clean_end = clean_date(raw_end)
        clean_dob = clean_date(raw_dob)

        standardized_df["Loan Disbursement Date (DDMMYYYY)"] = np.where(

            (
                raw_start.notna()
            )

            &

            (
                clean_start != clean_dob
            ),

            clean_start,

            np.nan

        )

        standardized_df["Loan End date (DDMMYYYY)"] = np.where(

            (
                raw_end.notna()
            )

            &

            (
                clean_end != clean_dob
            ),

            clean_end,

            np.nan

        )

        # =====================================================
        # CLEAN OTHER FIELDS
        # =====================================================

        standardized_df["Mobile No"] = clean_mobile(

            standardized_df["Mobile No"]

        )

        standardized_df["MAIN MEMBER AGE"] = clean_age(

            standardized_df["MAIN MEMBER AGE"]

        )

        standardized_df["Nominee Age"] = clean_age(

            standardized_df["Nominee Age"]

        )

        # =====================================================
        # LOAN TERM
        # =====================================================

        start_date = pd.to_datetime(

            standardized_df[
                "Loan Disbursement Date (DDMMYYYY)"
            ],

            errors="coerce"

        )

        end_date = pd.to_datetime(

            standardized_df[
                "Loan End date (DDMMYYYY)"
            ],

            errors="coerce"

        )

        valid_dates = (
            start_date.notna()
            &
            end_date.notna()
        )

        months = np.where(

            valid_dates,

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
        # SERIAL NUMBER
        # =====================================================

        standardized_df["Sr. no."] = range(
            1,
            len(standardized_df) + 1
        )

        # =====================================================
        # PREMIUM
        # =====================================================

        standardized_df["Rate"] = np.where(

            standardized_df["Final SA"].notna(),

            RATE_PER_LAKH,

            np.nan

        )

        standardized_df["Premium (Excl. GST)"] = np.where(

            standardized_df["Final SA"].notna(),

            (
                standardized_df["Final SA"] / 100000
            ) * RATE_PER_LAKH,

            np.nan

        )

        standardized_df["GST amount"] = np.where(

            standardized_df["Premium (Excl. GST)"].notna(),

            standardized_df["Premium (Excl. GST)"] * GST_RATE,

            np.nan

        )

        standardized_df["Total Premium (incl GST)"] = np.where(

            standardized_df["Premium (Excl. GST)"].notna(),

            standardized_df["Premium (Excl. GST)"]
            +
            standardized_df["GST amount"],

            np.nan

        )

        # =====================================================
        # DUPLICATE FIELDS
        # =====================================================

        standardized_df["Aviva Calculation SA"] = standardized_df["Final SA"]

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
        # FINAL FORMAT
        # =====================================================

        standardized_df = standardized_df[MASTER_COLUMNS]

        # =====================================================
        # APPEND
        # =====================================================

        final_master_df = pd.concat(

            [final_master_df, standardized_df],

            ignore_index=True

        )

    # =====================================================
    # OUTPUT
    # =====================================================

    st.markdown("---")

    st.subheader("📋 Final Output")

    st.dataframe(
        final_master_df,
        use_container_width=True
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
