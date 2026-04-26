import streamlit as st
import pandas as pd
import warnings
import io
import re
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font

# Hide harmless Excel formatting warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configuration and Constants
BANK_NAME_MAP = {
    "People's Bank": "Peoples Bank",
    "Commercial Bank Of Ceylon PLC": "Commercial Bank PLC",
}
NSB_ACCOUNT_L = 857
OTHER_ACCOUNT_L = 100011378759

COL_WIDTHS = {
    'A': 4.62, 'B': 4.62, 'C': 3.62, 'D': 12.62, 'E': 20.62,
    'F': 2.62, 'G': 9.62, 'H': 12.62, 'I': 3.62, 'J': 4.62,
    'K': 3.62, 'L': 12.62, 'M': 20.62, 'N': 15.62, 'O': 15.62,
    'P': 6.62, 'Q': 6.62
}

COL_FORMATS = {
    'A': '0000', 'B': '0000', 'C': '000', 'D': '000000000000', 'E': 'General',
    'F': '00', 'G': '000000000', 'H': '000000000000', 'I': 'General',
    'J': '0000', 'K': '000', 'L': '000000000000', 'M': 'General',
    'N': 'General', 'O': 'General', 'P': '000000', 'Q': '000000'
}

# Helper Functions
def normalize_bank_name(name):
    if pd.isna(name): return name
    return BANK_NAME_MAP.get(str(name).strip(), str(name).strip())

def lookup_bank_code(branch_df, bank_name):
    bank_name = normalize_bank_name(bank_name)
    if pd.isna(bank_name) or bank_name == "": return ""
    m = branch_df[branch_df["Bank"].str.lower() == bank_name.lower()]
    if m.empty:
        m = branch_df[branch_df["Bank"].str.contains(bank_name, case=False, na=False)]
    return str(m.iloc[0]["Bank Code"]) if not m.empty else "NOT FOUND"

def lookup_branch_code(branch_df, bank_name, branch_name):
    bank_name = normalize_bank_name(bank_name)
    branch_name = str(branch_name).strip() if not pd.isna(branch_name) else ""
    if pd.isna(bank_name) or bank_name == "" or branch_name == "": return ""
    bank_rows = branch_df[branch_df["Bank"].str.lower() == bank_name.lower()]
    if bank_rows.empty:
        bank_rows = branch_df[branch_df["Bank"].str.contains(bank_name, case=False, na=False)]
    if bank_rows.empty: return "BANK NOT FOUND"
    bm = bank_rows[bank_rows["Branch Name"].str.lower() == branch_name.lower()]
    if bm.empty:
        bm = bank_rows[bank_rows["Branch Name"].str.contains(branch_name, case=False, na=False)]
    return str(bm.iloc[0]["Branch Code"]) if not bm.empty else "BRANCH NOT FOUND"

def clean_account_number(acc):
    acc = str(acc).strip().replace("-", "").replace(" ", "")
    if len(acc) == 15:
        acc = acc[3:]
    return acc

def to_int_safe(val):
    try: return int(str(val).strip())
    except: return 0

def format_amount_int(amount):
    try: return int(round(float(amount) * 100))
    except: return 0

def format_maturity_int(mat_date):
    if pd.isna(mat_date): return 0
    if isinstance(mat_date, (datetime.datetime, datetime.date)):
        return int(mat_date.strftime("%y%m%d"))
    try: return int(pd.to_datetime(mat_date).strftime("%y%m%d"))
    except: return 0

def format_name(name):
    name = str(name).strip().replace(".", " ")
    return re.sub(r' +', ' ', name).strip()

def build_xls_row(row, account_l, is_nsb):
    acc_raw = clean_account_number(row.get("Acc. No", ""))
    try: acc_num = int(acc_raw)
    except: acc_num = 0

    return [
        0,
        to_int_safe(row.get("Bank Code", 0)),
        to_int_safe(row.get("Branch Code", 0)),
        acc_num,
        format_name(row.get("Customer Name", "")),
        23,
        0,
        format_amount_int(row.get("Amount", 0)),
        "slr",
        7010,
        660,
        account_l,
        "NSBFMC",
        "NSBFMC",
        "NSBFMC",
        format_maturity_int(row.get("Maturity Date", None)),
        0,
    ]

def generate_excel_bytes(df, account_l, is_nsb):
    wb = Workbook()
    ws = wb.active
    for letter, width in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = width

    for _, row in df.iterrows():
        ws.append(build_xls_row(row, account_l, is_nsb))

    tnr = Font(name="Times New Roman", size=10)
    for row_cells in ws.iter_rows():
        for cell in row_cells:
            col_letter = cell.column_letter
            cell.number_format = COL_FORMATS.get(col_letter, 'General')
            cell.font = tnr

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generate_prn_bytes(df):
    lines = ["BankCode|BranchCode|AccountNo|Amount|CustomerName|RefNo"]
    for _, row in df.iterrows():
        try: amount_str = f"{float(row.get('Amount', '')):.2f}"
        except: amount_str = str(row.get("Amount", ""))
        lines.append("|".join([
            str(row.get("Bank Code", "")).strip(),
            str(row.get("Branch Code", "")).strip(),
            str(row.get("Acc. No", "")).strip(),
            amount_str,
            str(row.get("Customer Name", "")).strip(),
            str(row.get("RefNo", "")).strip(),
        ]))
    return "\n".join(lines).encode('utf-8')

# User Interface
st.title("Payment List Processor")
st.write("Upload your required files below.")

col1, col2 = st.columns(2)
with col1:
    payment_file = st.file_uploader("Upload Payment List", type=["xlsx"])
with col2:
    directory_file = st.file_uploader("Upload Bank Directory", type=["xlsx"])

if st.button("Run Process"):
    if payment_file is not None and directory_file is not None:
        st.write("Processing your files. Please wait.")
        try:
            # 1. Load the bank directory
            branch_df = pd.read_excel(directory_file, sheet_name="Branch NEW", header=2)
            branch_df.columns = [
                "Bank", "Bank Code", "Branch Code", "Branch Name",
                "Branch Address", "Tel No1", "Tel No2", "Tel No3", "Fax No", "District"
            ]
            branch_df = branch_df.dropna(subset=["Bank", "Bank Code", "Branch Code"])
            branch_df["Bank"] = branch_df["Bank"].astype(str).str.strip()
            branch_df["Branch Name"] = branch_df["Branch Name"].astype(str).str.strip()
            branch_df["Bank Code"] = branch_df["Bank Code"].astype(str).str.strip()
            branch_df["Branch Code"] = branch_df["Branch Code"].astype(str).str.strip()

            # 2. Load the payment list
            pay_df = pd.read_excel(payment_file, header=1)
            pay_df.columns = pay_df.columns.str.strip()
            pay_df = pay_df.dropna(how="all").reset_index(drop=True)

            # 3. Clean the payment list
            pay_df = pay_df[~pay_df.iloc[:, 0].astype(str).str.strip().str.startswith("Print Date")]
            work_df = pay_df[pay_df["Pay By"].astype(str).str.strip() != "Other"].copy()

            # 4. Lookup Bank Codes and Branch Codes
            work_df["Bank Code"] = work_df.apply(lambda r: lookup_bank_code(branch_df, r["Bank Name"]), axis=1)
            work_df["Branch Code"] = work_df.apply(lambda r: lookup_branch_code(branch_df, r["Bank Name"], r["Branch Name"]), axis=1)

            # 5. Split into three groups
            rtgs_mask = work_df["Pay By"].astype(str).str.strip().str.upper() == "RTGS"
            rtgs_df = work_df[rtgs_mask].copy()
            remaining_df = work_df[~rtgs_mask].copy()

            nsb_mask = (remaining_df["Bank Name"].astype(str).str.strip().str.lower() == "national savings bank")
            nsb_df = remaining_df[nsb_mask].copy()
            other_banks_df = remaining_df[~nsb_mask].copy()

            st.success("Processing complete. Download your files below.")

            # 6. Create download buttons
            st.subheader("RTGS Files")
            st.download_button("Download RTGS CSV", rtgs_df.to_csv(index=False).encode('utf-8'), "output_RTGS.csv", "text/csv")

            st.subheader("NSB Files")
            col3, col4 = st.columns(2)
            with col3:
                st.download_button("Download NSB CSV", nsb_df.to_csv(index=False).encode('utf-8'), "output_NSB.csv", "text/csv")
            with col4:
                nsb_xls = generate_excel_bytes(nsb_df, NSB_ACCOUNT_L, is_nsb=True)
                st.download_button("Download NSB Excel", nsb_xls, "output_NSB.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.subheader("Other Banks Files")
            col5, col6, col7 = st.columns(3)
            with col5:
                st.download_button("Download Other Banks CSV", other_banks_df.to_csv(index=False).encode('utf-8'), "output_OtherBanks.csv", "text/csv")
            with col6:
                other_xls = generate_excel_bytes(other_banks_df, OTHER_ACCOUNT_L, is_nsb=False)
                st.download_button("Download Other Banks Excel", other_xls, "output_OtherBanks.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col7:
                prn_bytes = generate_prn_bytes(other_banks_df)
                st.download_button("Download Other Banks PRN", prn_bytes, "output_OtherBanks.prn", "text/plain")

        except Exception as e:
            st.error(f"An error occurred: {e}")

    else:
        st.warning("Please upload both files before clicking Run Process.")