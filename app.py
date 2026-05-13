import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="Insurance Data Mapper",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Insurance Data Standardization Engine")

st.write(
    "Upload any insurance excel file and map fields dynamically."
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
# UPLOAD FILES
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

        st.subheader(f"📄 Processing File: {file.name}")

        # =====================================================
        # READ RAW FILE
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
        # PREVIEW INPUT
        # =====================================================

        st.write("### Input Data Preview")

        st.dataframe(
            df.head(),
            use_container_width=True
        )

        # =====================================================
        # CREATE OUTPUT DF
        # =====================================================

        standardized_df = pd.DataFrame()

        for col in MASTER_COLUMNS:

            standardized_df[col] = np.nan

        # =====================================================
        # DYNAMIC FIELD MAPPING
        # =====================================================

        st.write("## 🛠 Dynamic Field Mapping")

        all_columns = list(df.columns)

        mapping_config = {}

        important_fields = [

            "Loan Account No.",
            "Name of Primary Loan borrower",
            "Gender",
            "Date of Birth (DDMMMYYYY)",
            "MAIN MEMBER AGE",
            "Mobile No",
            "Pincode",
            "Branch Name",
            "Zone",
            "Loan Type",
            "Loan Outstanding Amount",
            "Sum Assured",
            "Nominee Name",
            "Relationship of the Nominee with Insurance covered Person",
            "Nominee Age",
            "Loan Disbursement Date (DDMMYYYY)",
            "Loan End date (DDMMYYYY)",
            "Address            (First Life)",
            "Address 1            (First Life)",
            "Address 2            (First Life)",
            "Email Id"

        ]

        for output_field in important_fields:

            selected_columns = st.multiselect(

                f"Select input field(s) for ➜ {output_field}",

                options=all_columns,

                default=[],

                key=f"{file.name}_{output_field}"

            )

            mapping_config[output_field] = selected_columns

        # =====================================================
        # APPLY USER MAPPING
        # =====================================================

        for output_field, selected_columns in mapping_config.items():

            if len(selected_columns) == 1:

                standardized_df[output_field] = df[
                    selected_columns[0]
                ]

            elif len(selected_columns) > 1:

                merged_data = df[
                    selected_columns
                ].astype(str)

                merged_data = merged_data.replace(
                    "nan",
                    ""
                )

                standardized_df[output_field] = merged_data.apply(

                    lambda row: " ".join(

                        [

                            str(x).strip()

                            for x in row

                            if str(x).strip() != ""

                        ]

                    ),

                    axis=1

                )

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

            standardized_df["Sum Assured"].notna(),

            standardized_df["Sum Assured"],

            standardized_df["Loan Outstanding Amount"]

        )

        standardized_df["Sum Assured"] = (

            standardized_df["Final SA"]

        )

        # =====================================================
        # CLEAN DOB
        # =====================================================

        standardized_df["Date of Birth (DDMMMYYYY)"] = clean_date(

            standardized_df["Date of Birth (DDMMMYYYY)"]

        )

        # =====================================================
        # CLEAN LOAN DATES
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

        standardized_df["Rate"] = np.where(

            standardized_df["Final SA"].notna(),

            RATE_PER_LAKH,

            np.nan

        )

        standardized_df["Premium (Excl. GST)"] = np.where(

            standardized_df["Final SA"].notna(),

            (

                standardized_df["Final SA"]

                / 100000

            ) * RATE_PER_LAKH,

            np.nan

        )

        standardized_df["GST amount"] = np.where(

            standardized_df["Premium (Excl. GST)"].notna(),

            standardized_df["Premium (Excl. GST)"]

            * GST_RATE,

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
