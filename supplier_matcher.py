import pandas as pd
import re
from rapidfuzz import process, fuzz

# FILE NAMES

PRINTIQ_FILE = "printiq_suppliers.xlsx"
XERO_FILE = "xero_suppliers.xlsx"
OUTPUT_FILE = "acume_supplier_matching.xlsx"

# LOAD FILES

print("Loading Excel files...")

printiq_df = pd.read_excel(PRINTIQ_FILE, engine="openpyxl")
xero_df = pd.read_excel(XERO_FILE, engine="openpyxl")

print("PRINTIQ COLUMNS:")
print(printiq_df.columns.tolist())

print("XERO COLUMNS:")
print(xero_df.columns.tolist())

# Clean column headers (fix hidden spaces / characters)
printiq_df.columns = printiq_df.columns.str.strip().str.replace("\ufeff", "")
xero_df.columns = xero_df.columns.str.strip().str.replace("\ufeff", "")

# COLUMN MAPPING

PRINTIQ_CODE_COL = "Code"
PRINTIQ_NAME_COL = "Name"

XERO_ACCOUNT_COL = "AccountNumber"
XERO_NAME_COL = "ContactName"

# CLEANING FUNCTIONS

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def normalize_company_name(name):
    if pd.isna(name):
        return ""

    name = str(name).upper().strip()

    # remove punctuation
    name = re.sub(r"[^\w\s]", " ", name)

    remove_words = [
        "LIMITED",
        "LTD",
        "PTY",
        "INC",
        "LLC",
        "COMPANY",
        "CO"
    ]

    # remove whole words only
    for word in remove_words:
        name = re.sub(rf"\b{word}\b", "", name)

    return " ".join(name.split())

# CLEAN DATA

print("Cleaning supplier data...")

printiq_df["MATCH_CODE"] = printiq_df[PRINTIQ_CODE_COL].apply(clean_text)
xero_df["MATCH_CODE"] = xero_df[XERO_ACCOUNT_COL].apply(clean_text)


# EXACT MATCH (BY CODE)

print("Performing exact matches...")

matched_df = pd.merge(
    printiq_df,
    xero_df,
    on="MATCH_CODE",
    how="left",
    suffixes=("_PRINTIQ", "_XERO")
)

exact_matches = matched_df[
    matched_df[XERO_ACCOUNT_COL].notna()
].copy()

unmatched = matched_df[
    matched_df[XERO_ACCOUNT_COL].isna()
].copy()

# FUZZY MATCH (BY NAME)

print("Finding fuzzy matches...")

xero_name_list = xero_df[XERO_NAME_COL].apply(normalize_company_name).tolist()

fuzzy_results = []

for _, row in unmatched.iterrows():

    supplier_name = normalize_company_name(row[PRINTIQ_NAME_COL])

    if supplier_name == "":
        continue

    match = process.extractOne(
        supplier_name,
        xero_name_list,
        scorer=fuzz.token_sort_ratio
    )

    if match:

        best_match_name, score = match[0], match[1]

        if score >= 80:

            match_rows = xero_df[
                            xero_df[XERO_NAME_COL].apply(normalize_company_name) == best_match_name
                        ]

            if not match_rows.empty:

                xero_match_row = match_rows.iloc[0]

                fuzzy_results.append({
                    "PrintIQ Supplier Code": row[PRINTIQ_CODE_COL],
                    "PrintIQ Supplier Name": row[PRINTIQ_NAME_COL],
                    "Suggested Xero AccountNumber": xero_match_row[XERO_ACCOUNT_COL],
                    "Suggested Xero ContactName": xero_match_row[XERO_NAME_COL],
                    "Match Score": score
                })

fuzzy_df = pd.DataFrame(fuzzy_results)

# SORT RESULTS

print("Sorting results...")

exact_matches = exact_matches.sort_values(by=[PRINTIQ_NAME_COL])
unmatched = unmatched.sort_values(by=[PRINTIQ_NAME_COL])

if not fuzzy_df.empty:
    fuzzy_df = fuzzy_df.sort_values(by=["Match Score"], ascending=False)

# EXPORT TO EXCEL

print("Exporting results...")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    exact_matches.to_excel(writer, sheet_name="Exact Matches", index=False)
    unmatched.to_excel(writer, sheet_name="Unmatched", index=False)
    fuzzy_df.to_excel(writer, sheet_name="Fuzzy Suggestions", index=False)

# DONE

print("")
print("======================================")
print("ACUME SUPPLIER MATCH COMPLETE")
print(f"Output file: {OUTPUT_FILE}")
print("======================================")

