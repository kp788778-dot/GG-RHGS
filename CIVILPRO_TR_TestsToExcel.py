import streamlit as st
import pdfplumber
import pandas as pd
import re
from collections import Counter
from io import BytesIO

st.title("Getting TR Results in a Pivot Table")
st.write('''WHAT IS THIS?
If you need to check the total number of TRs raised against the claim from Western Geotech, this will
allow you to reformat the total list from CivilPro into an excel table.


HOW TO USE THIS CODE:
1. On CivilPro, select the test requests that you want to look at. I usually press the box in the top right to select all.
2. Click on the printer on the right hand side, and then press 'Test Requests (pdf)'. Save this on your computer.
3. Upload that PDF document (usually called TRs 1_2_etc.pdf) into the drag and drop box.
4. This will take a minute to load. Once loaded, a preview box of the results will be visable. 
5. At the very bottom, there is a button labelled "download excel file". Press this to download.
6. This excel file will have two sheets. "All tests" contains a table with a list of all of the tests. "Tests summary" contains a summary table with the number of each type of test.
7. If you have any questions about using this, or it breaks, please ask Kieran.''')

uploaded_file = st.file_uploader("Upload Test Request PDF", type="pdf")

# Full name replacements -  when the scrape doesn't get all of the name, this will replace the segment with full name 
full_names = {
    "WA 115.1: Particle Size Distribution": "WA 115.1: Particle Size Distribution: Sieving and Decantation Method",
    "WA 141.1: Determination of the California": "WA 141.1: Determination of the California Bearing Ratio of a Soil: Standard Laboratory Method for a Remoulded Specimen",
    "WA 115.2: Particle Size Distribution: Abbreviated": "WA 115.2: Particle Size Distribution: Abbreviated Method for Coarse and Medium Grained Soils",
    "WA 133.1: Dry Density/Moisture Content": "WA 133.1: Dry Density/Moisture Content Relationship: Modified Compaction Fine and Medium Grained Soils",
    "WA 324.2: Determination of Field Density": "WA 324.2: Determination of Field Density: Nuclear Method",
    "Construction Moisture Content (WA 110.1)": "Construction Moisture Content (WA 110.1) - Convection Oven Method",
    "Construction Moisture Content (WA 110.2)": "Construction Moisture Content (WA 110.2) - Microwave Oven Method"
}

def normalize_method_name(name):
    for key, full in full_names.items():
        if key in name:
            return full
    return name

def replace_field_density(methods):
    count_133 = sum(c for m, c in methods if "WA 133.1" in m)
    count_134 = sum(c for m, c in methods if "WA 134.1" in m)
    count_324 = sum(c for m, c in methods if "WA 324.2" in m)

    count_134_combined = count_134 + count_324

    mapping = {
        (2, 6): "Field Density Package - 6 NDM Sites x 2 MDD",
        (2, 3): "Field Density Package - 3 NDM Sites x 2 MDD",
        (3, 3): "Field Density Package - 3 NDM Sites x 3 MDD",
        (3, 6): "Field Density Package - 6 NDM Sites x 3 MDD",
        (3, 9): "Field Density Package - 9 NDM Sites x 3 MDD",
        (6, 6): "Field Density Package - 6 NDM Sites x 6 MDD"
    }

    key = (count_133, count_134_combined)
    if key in mapping:
        methods = [(m, c) for m, c in methods if "WA 133.1" not in m and "WA 134.1" not in m and "WA 324.2" not in m]
        methods.append((mapping[key], 1))

    return methods

def process_pdf(uploaded_file):
    rows = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            tr_match = re.search(r"TR:\s*(\d+)", text)
            if not tr_match:
                continue
            tr = tr_match.group(1)

            lot_match = re.search(r"Lot No:\s*([A-Za-z0-9\-]+)", text)
            lot_no = lot_match.group(1) if lot_match else ""
            lot_type = lot_no[:2] if lot_no else ""

            date_match = re.search(r"When Req'd\s+\w+,\s+(\d{2}\s+\w+\s+\d{4})", text)
            when_reqd = (
                pd.to_datetime(date_match.group(1), format="%d %b %Y").date()
                if date_match else None
            )

            location_method = None
            loc_match = re.search(r"Location Method:\s*([A-Za-z ]+)", text)
            if loc_match:
                location_method = loc_match.group(1).strip().lower()

            methods = []

            if location_method == "tester locates":
                for count, std, code, desc in re.findall(
                    r"(\d+)\s+(WA|AS)\s+([\d\.]+):\s+([^\n]+)", text
                ):
                    method_name = f"{std} {code}: {desc.strip()}"
                    method_name = normalize_method_name(method_name)
                    methods.append((method_name, int(count)))

            elif location_method == "location specified":
                numbered = re.findall(
                    r"\d+-\d+\s+(WA|AS)\s+([\d\.]+):\s+([^\n]+)", text
                )
                if numbered:
                    counter = Counter()
                    for std, code, desc in numbered:
                        clean_desc = re.sub(r"\s+0\.0+.*$", "", desc).strip()
                        method_name = f"{std} {code}: {clean_desc}"
                        method_name = normalize_method_name(method_name)
                        counter[method_name] += 1
                    for method, cnt in counter.items():
                        methods.append((method, cnt))

            methods = replace_field_density(methods)

            if not methods:
                rows.append([tr, when_reqd, "", "not used", lot_no, lot_type])
            else:
                for method, count in methods:
                    rows.append([tr, when_reqd, method, count, lot_no, lot_type])

    df = pd.DataFrame(
        rows,
        columns=["TR", "When Req'd (date)", "Test method", "No. tests","Lot no.","Lot Type"]
    )

    return df


#Runs once the file is uploaded
from io import BytesIO

if uploaded_file is not None:
    df = process_pdf(uploaded_file)

    st.success("Processing is complete.")
    st.dataframe(df)

    #  Create Pivot Summary 
    df["No. tests"] = pd.to_numeric(df["No. tests"], errors="coerce")

    pivot_df = (
        df.groupby("Test method", dropna=False)["No. tests"]
        .sum()
        .reset_index()
        .sort_values("No. tests", ascending=False)
    )

    #  Create Excel in memory 
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Write All tests sheet
        df.to_excel(writer, sheet_name="All tests", index=False)

        # Write summary sheet
        pivot_df.to_excel(writer, sheet_name="Tests summary", index=False)

        # Auto-fit column widths
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column_cells in worksheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = max_length + 2
                worksheet.column_dimensions[column].width = adjusted_width

    output.seek(0)

    st.download_button(
        label="Download Excel File",
        data=output,
        file_name="processed_test_requests.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
