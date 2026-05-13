import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from difflib import SequenceMatcher

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Insurance Standardization Engine",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Insurance Data Standardization Engine")

st.write(
    "Upload any insurance excel and convert it into fixed insurer output format."
)

# =====================================================
# SETTINGS
# =====================================================

RATE_PER_LAKH = 320.3
GST_RATE = 0.18

# =====================================================
# FINAL OUTPUT FORMAT
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
# COLUMN ALIASES
# =====================================================

ALIASES = {

    "Loan Account No.": [

        "a/c number",
        "account number",
        "account no",
        "loan account",
        "membership no",
        "membership account no",
        "loan no",
        "lan"

    ],

    "Name of Primary Loan borrower": [

        "member name",
        "borrower name",
        "customer name",
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

    "Branch Name": [

        "branch name",
        "branch"

    ],

    "Loan Outstanding Amount": [

        "loan amount",
        "loan outstanding"

    ],

    "Sum Assured": [

        "sum assured",
        "sum insured",
        "coverage",
        "calculation sa",
        "gtl"

    ],

    "Nominee Name": [

        "nominee name",
        "nominee"

    ],

    "Relationship of the Nominee with Insurance covered Person": [

        "nominee relationship",
        "nominee relatonship",
        "relationship",
        "relation"

    ],

    "Nominee Age": [

        "nominee age"

    ],

    "Loan Disbursement Date (DDMMYYYY)": [

        "loan start date",
        "loan start",
        "disbursement"

    ],

    "Loan End date (DDMMYYYY)": [

        "loan end date",
        "loan end"

    ]

}

# =====================================================
# AI MATCHING
# =====================================================

def similarity(a, b):

    return SequenceMatcher(
        None,
        a,
        b
    ).ratio()

def detect_column(columns, aliases):

    best_match = None
    best_score = 0

    for col in columns:

        cleaned_col = (

            str(col)

            .lower()

            .replace("_", " ")

            .replace("-", " ")

            .strip()

        )

        for alias in aliases:

            score = similarity(
                cleaned_col,
                alias
            )

            if alias in cleaned_col:

                score += 0.3

            if score > best_score:

                best_score = score
                best_match = col

    if best_score >= 0.45:

        return best_match

    return None

# =====================================================
# CLEAN FUNCTIONS
# =====================================================

def clean_money(series):

    return pd.to_numeric(

        series.astype(str)

        .str.replace(",", "", regex=False)

        .str.replace("₹", "", regex=False)

        .str.replace("/-", "", regex=False)

        .str.replace(" ", "", regex=False),

        errors="coerce"

    )

def clean_mobile(series):

    return (

        series.astype(str)

        .str.replace(r"\D", "", regex=True)

        .str[-10:]

    )

def clean_aadhar(series):

    return (

        series.astype(str)

        .str.replace(r"\D", "", regex=True)

        .str[-12:]

    )

def clean_age(series):

    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0).astype(int)

def clean_date(series):

    dates = pd.to_datetime(
        series,
        errors="coerce"
    )

    return dates.dt.strftime("%d%b%Y")

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

    error_log = []

    for file in uploaded_files:

        try:

            # =====================================================
            # READ FILE
            # =====================================================

            raw_df = pd.read_excel(
                file,
                header=None
            )

            # =====================================================
            # FIND HEADER ROW
            # =====================================================

            header_row = 0

            for i in range(min(10, len(raw_df))):

                row_text = " ".join(

                    raw_df.iloc[i]

                    .astype(str)

                    .str.lower()

                    .tolist()

                )

                if (

                    "member" in row_text

                    or

                    "name" in row_text

                    or

                    "account" in row_text

                ):

                    header_row = i
                    break

            # =====================================================
            # READ AGAIN USING HEADER
            # =====================================================

            df = pd.read_excel(

                file,

                header=header_row

            )

            # =====================================================
            # CLEAN DATAFRAME
            # =====================================================

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

                .lower()

                for col in df.columns

            ]

            # =====================================================
            # CREATE OUTPUT DF
            # =====================================================

            standardized_df = pd.DataFrame()

            for col in MASTER_COLUMNS:

                standardized_df[col] = ""

            # =====================================================
            # AUTO MAP
            # =====================================================

            mapping = {}

            for standard_col, alias_list in ALIASES.items():

                detected_col = detect_column(

                    df.columns,

                    alias_list

                )

                if detected_col:

                    standardized_df[
                        standard_col
                    ] = df[detected_col]

                    mapping[
                        standard_col
                    ] = detected_col

            # =====================================================
            # CLEAN IMPORTANT FIELDS
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

                standardized_df[
                    "Name of Primary Loan borrower"
                ].astype(str).str.strip() != ""

            ]

            standardized_df = standardized_df[

                standardized_df["Final SA"] > 0

            ]

            # =====================================================
            # CLEAN DATES
            # =====================================================

            date_cols = [

                "Date of Birth (DDMMMYYYY)",

                "Loan Disbursement Date (DDMMYYYY)",

                "Loan End date (DDMMYYYY)"

            ]

            for col in date_cols:

                standardized_df[col] = clean_date(

                    standardized_df[col]

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

            standardized_df["Nominee Age"] = clean_age(

                standardized_df["Nominee Age"]

            )

            # =====================================================
            # CALCULATE LOAN TERM
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
            # PREMIUM
            # =====================================================

            standardized_df["Rate"] = RATE_PER_LAKH

            standardized_df["Premium (Excl. GST)"] = (

                standardized_df["Final SA"]

                / 100000

            ) * RATE_PER_LAKH

            standardized_df["GST amount"] = (

                standardized_df["Premium (Excl. GST)"]

                * GST_RATE

            )

            standardized_df["Total Premium (incl GST)"] = (

                standardized_df["Premium (Excl. GST)"]

                +

                standardized_df["GST amount"]

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
            # KEEP FINAL FORMAT
            # =====================================================

            standardized_df = standardized_df[MASTER_COLUMNS]

            # =====================================================
            # APPEND
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

    col1, col2, col3, col4 = st.columns(4)

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
            "Total GST",
            f"₹ {final_master_df['GST amount'].sum():,.2f}"
        )

    with col4:

        st.metric(
            "Total Premium",
            f"₹ {final_master_df['Total Premium (incl GST)'].sum():,.2f}"
        )

    # =====================================================
    # OUTPUT
    # =====================================================

    st.subheader("📋 Final Output")

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
