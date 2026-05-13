import streamlit as st
import pandas as pd
from io import BytesIO

# ============================================
# PAGE CONFIG
# ============================================

st.set_page_config(
    page_title="Insurance Standardization Engine",
    page_icon="📊",
    layout="wide"
)

# ============================================
# HEADER
# ============================================

st.title("📊 Insurance Data Standardization Engine")

st.write(
    "Upload multiple insurance Excel files and convert them into one fixed insurer format."
)

# ============================================
# SETTINGS
# ============================================

RATE_PER_LAKH = 320.3
GST_RATE = 0.18

# ============================================
# MASTER OUTPUT FORMAT
# ============================================

MASTER_COLUMNS = [

    'Sr. no.',
    'Zone',
    'Branch Name',
    'Branch Code',
    'Loan Type',
    'Loan Account No.',
    'Name of Primary Loan borrower',
    'Gender',
    'Date of Birth (DDMMMYYYY)',
    'Mobile No',
    'Loan Outstanding Amount',
    'Sum Assured',
    'Rate',
    'Premium (Excl. GST)',
    'GST amount',
    'Total Premium (incl GST)',
    'Aviva Remarks'

]

# ============================================
# COLUMN ALIASES
# ============================================

ALIASES = {

    'Loan Account No.': [

        'a/c number',
        'account no',
        'loan account no',
        'membership no',
        'lan'

    ],

    'Name of Primary Loan borrower': [

        'member name',
        'customer name',
        'borrower name',
        'name'

    ],

    'Gender': [

        'gender',
        'sex'

    ],

    'Date of Birth (DDMMMYYYY)': [

        'dob',
        'date of birth'

    ],

    'Mobile No': [

        'mobile',
        'mobile number',
        'phone'

    ],

    'Branch Name': [

        'branch',
        'branch name'

    ],

    'Loan Outstanding Amount': [

        'loan amount',
        'outstanding amount'

    ],

    'Sum Assured': [

        'sum assured',
        'sum insured',
        'coverage',
        'calculation sa',
        'sa'

    ]

}

# ============================================
# DETECT COLUMN FUNCTION
# ============================================

def detect_column(df_columns, aliases):

    for col in df_columns:

        cleaned_col = str(col).lower().strip()

        for alias in aliases:

            if alias in cleaned_col:

                return col

    return None

# ============================================
# FILE UPLOAD
# ============================================

uploaded_files = st.file_uploader(
    "📂 Upload Excel Files",
    type=['xlsx'],
    accept_multiple_files=True
)

# ============================================
# PROCESS FILES
# ============================================

if uploaded_files:

    final_master_df = pd.DataFrame()

    for file in uploaded_files:

        try:

            df = pd.read_excel(file)

            df.dropna(
                how='all',
                inplace=True
            )

            df.columns = df.columns.astype(str)

            df.columns = df.columns.str.strip()

            # ============================================
            # CREATE STANDARD OUTPUT
            # ============================================

            standardized_df = pd.DataFrame()

            for col in MASTER_COLUMNS:

                standardized_df[col] = 'NA'

            # ============================================
            # MAP COLUMNS
            # ============================================

            for standard_col, aliases in ALIASES.items():

                detected_col = detect_column(
                    df.columns,
                    aliases
                )

                if detected_col:

                    standardized_df[
                        standard_col
                    ] = df[detected_col]

            # ============================================
            # CLEAN SA
            # ============================================

            standardized_df['Sum Assured'] = (

                standardized_df['Sum Assured']
                .astype(str)
                .str.replace(',', '')
                .str.replace('₹', '')
                .str.strip()

            )

            standardized_df['Sum Assured'] = pd.to_numeric(

                standardized_df['Sum Assured'],

                errors='coerce'

            ).fillna(0)

            # ============================================
            # REMOVE INVALID ROWS
            # ============================================

            standardized_df = standardized_df[

                standardized_df['Sum Assured'] > 0

            ]

            # ============================================
            # SERIAL NUMBER
            # ============================================

            standardized_df['Sr. no.'] = range(

                1,

                len(standardized_df) + 1

            )

            # ============================================
            # RATE
            # ============================================

            standardized_df['Rate'] = RATE_PER_LAKH

            # ============================================
            # PREMIUM
            # ============================================

            standardized_df['Premium (Excl. GST)'] = (

                standardized_df['Sum Assured']

                / 100000

            ) * RATE_PER_LAKH

            # ============================================
            # GST
            # ============================================

            standardized_df['GST amount'] = (

                standardized_df['Premium (Excl. GST)']

                * GST_RATE

            )

            # ============================================
            # TOTAL PREMIUM
            # ============================================

            standardized_df['Total Premium (incl GST)'] = (

                standardized_df['Premium (Excl. GST)']

                + standardized_df['GST amount']

            )

            # ============================================
            # REMARKS
            # ============================================

            standardized_df['Aviva Remarks'] = (

                'Processed Successfully'

            )

            # ============================================
            # APPEND FILE
            # ============================================

            final_master_df = pd.concat(

                [final_master_df, standardized_df],

                ignore_index=True

            )

        except Exception as e:

            st.error(
                f"Error in {file.name}: {e}"
            )

    # ============================================
    # DASHBOARD
    # ============================================

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

    # ============================================
    # SHOW OUTPUT
    # ============================================

    st.subheader("📋 Final Output")

    st.dataframe(
        final_master_df,
        use_container_width=True
    )

    # ============================================
    # DOWNLOAD FILE
    # ============================================

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