
import os
import sys
import pandas as pd

# ============================================================
# PATIENT AR PROGRESS DASHBOARD
# ============================================================

INPUT_FILE = "Patient AR -Progress 10272026.xlsx"
OUTPUT_FILE = "Patient_AR_Progress_Dashboard.xlsx"

# CHANGE THIS TO "Status" IF YOUR CATEGORY VALUES ARE IN STATUS
CATEGORY_COLUMN = "Remarks"

# Leave as None to auto-detect the outstanding amount column.
# If auto-detection fails, enter the exact Excel column name.
OUTSTANDING_COLUMN = None

DATE_COLUMN = "Post/Void Dt"
GUARANTOR_ID_COLUMN = "Guar id#"

# Paid in Full intentionally excluded
CATEGORIES = [
    "Collection",
    "Deceased",
    "Rebilled",
    "Turn Off",
    "Turn on",
    "Not Yet Worked"
]

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_amount(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)


def find_outstanding_column(df):
    if OUTSTANDING_COLUMN:
        if OUTSTANDING_COLUMN not in df.columns:
            raise ValueError(
                f"Column '{OUTSTANDING_COLUMN}' was not found. "
                f"Available columns: {list(df.columns)}"
            )
        return OUTSTANDING_COLUMN

    candidates = [
        "Outst Amt",
        "Outstanding AR",
        "Outstanding Amount",
        "Outst Amount",
        "Outstanding Balance",
        "Balance",
        "Patient Credit(OC)"
    ]

    for candidate in candidates:
        if candidate in df.columns:
            return candidate

    raise ValueError(
        "Could not identify the outstanding amount column.\n"
        "Please set OUTSTANDING_COLUMN to the exact column name.\n"
        f"Available columns: {list(df.columns)}"
    )


# ============================================================
# LOAD SOURCE DATA
# ============================================================

def load_data():

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}\n"
            "Keep the Excel input file in the same folder as this script."
        )

    print("Reading source Excel workbook...")

    summary = pd.read_excel(
        INPUT_FILE,
        sheet_name="Summary",
        header=0
    )

    raw_data = pd.read_excel(
        INPUT_FILE,
        sheet_name="Sheet1"
    )

    if len(summary.columns) < 8:
        raise ValueError(
            "The Summary sheet does not contain the expected 8 columns."
        )

    # Rename Summary columns according to their positions
    summary = summary.iloc[:, :8].copy()

    summary.columns = [
        "Category",
        "2025 Claims",
        "2025 Amount",
        "2026 Claims",
        "2026 Amount",
        "Total Guarantors",
        "Outstanding AR",
        "Unique Guarantors"
    ]

    # Clean category names
    summary["Category"] = summary["Category"].apply(clean_text)

    # Remove Grand Total and Paid in Full
    summary = summary[
        summary["Category"].str.lower().isin(
            [x.lower() for x in CATEGORIES if x != "Not Yet Worked"]
        )
    ].copy()

    numeric_columns = [
        "2025 Claims",
        "2025 Amount",
        "2026 Claims",
        "2026 Amount",
        "Total Guarantors",
        "Outstanding AR",
        "Unique Guarantors"
    ]

    for column in numeric_columns:
        summary[column] = pd.to_numeric(
            summary[column],
            errors="coerce"
        ).fillna(0)

    if CATEGORY_COLUMN not in raw_data.columns:
        raise ValueError(
            f"Category column '{CATEGORY_COLUMN}' not found.\n"
            f"Available columns: {list(raw_data.columns)}"
        )

    if DATE_COLUMN not in raw_data.columns:
        raise ValueError(
            f"Date column '{DATE_COLUMN}' not found."
        )

    amount_column = find_outstanding_column(raw_data)

    print(f"Category column: {CATEGORY_COLUMN}")
    print(f"Outstanding amount column: {amount_column}")

    return summary, raw_data, amount_column


# ============================================================
# CREATE NOT YET WORKED CATEGORY FROM BLANK VALUES
# ============================================================

def calculate_not_yet_worked(raw_data, amount_column):

    data = raw_data.copy()

    # Convert blank, spaces and missing category values
    # into Not Yet Worked
    data["_Clean_Category"] = data[CATEGORY_COLUMN].apply(clean_text)

    blank_mask = data["_Clean_Category"] == ""

    not_worked = data[blank_mask].copy()

    print(
        f"Not Yet Worked records identified: {len(not_worked):,}"
    )

    # Convert dates and amounts
    not_worked["_Date"] = pd.to_datetime(
        not_worked[DATE_COLUMN],
        errors="coerce"
    )

    not_worked["_Amount"] = clean_amount(
        not_worked[amount_column]
    )

    # Calculate claim counts by year
    year_2025 = not_worked[
        not_worked["_Date"].dt.year == 2025
    ]

    year_2026 = not_worked[
        not_worked["_Date"].dt.year == 2026
    ]

    claims_2025 = len(year_2025)
    claims_2026 = len(year_2026)

    amount_2025 = year_2025["_Amount"].sum()
    amount_2026 = year_2026["_Amount"].sum()

    total_guarantors = len(not_worked)

    if GUARANTOR_ID_COLUMN in not_worked.columns:
        unique_guarantors = (
            not_worked[GUARANTOR_ID_COLUMN]
            .dropna()
            .nunique()
        )
    elif "Guarantor" in not_worked.columns:
        unique_guarantors = (
            not_worked["Guarantor"]
            .dropna()
            .nunique()
        )
    else:
        unique_guarantors = total_guarantors

    result = {
        "Category": "Not Yet Worked",
        "2025 Claims": claims_2025,
        "2025 Amount": amount_2025,
        "2026 Claims": claims_2026,
        "2026 Amount": amount_2026,
        "Total Guarantors": total_guarantors,
        "Outstanding AR": not_worked["_Amount"].sum(),
        "Unique Guarantors": unique_guarantors
    }

    return pd.DataFrame([result]), not_worked


# ============================================================
# PREPARE FINAL SUMMARY
# ============================================================

def prepare_summary(summary, raw_data, amount_column):

    not_worked_summary, not_worked_data = (
        calculate_not_yet_worked(raw_data, amount_column)
    )

    # Add Not Yet Worked to the existing Summary
    final_summary = pd.concat(
        [summary, not_worked_summary],
        ignore_index=True
    )

    # Keep categories in the requested order
    final_summary["Category"] = pd.Categorical(
        final_summary["Category"],
        categories=CATEGORIES,
        ordered=True
    )

    final_summary = final_summary.sort_values(
        "Category"
    ).reset_index(drop=True)

    final_summary["Category"] = (
        final_summary["Category"].astype(str)
    )

    return final_summary, not_worked_data


# ============================================================
# CALCULATE DASHBOARD KPIs
# ============================================================

def calculate_kpis(summary):

    return {
        "total_2025_claims": summary["2025 Claims"].sum(),
        "total_2026_claims": summary["2026 Claims"].sum(),
        "total_2025_amount": summary["2025 Amount"].sum(),
        "total_2026_amount": summary["2026 Amount"].sum(),
        "total_guarantors": summary["Total Guarantors"].sum(),
        "total_outstanding": summary["Outstanding AR"].sum(),
        "total_claims": (
            summary["2025 Claims"].sum()
            + summary["2026 Claims"].sum()
        ),
        "unique_guarantors": summary["Unique Guarantors"].sum()
    }


# ============================================================
# CREATE EXCEL DASHBOARD
# ============================================================

def create_dashboard(summary, raw_data, not_worked_data, kpis):

    print("Creating Excel dashboard...")

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="xlsxwriter"
    ) as writer:

        workbook = writer.book

        # ----------------------------------------------------
        # FORMATS
        # ----------------------------------------------------

        title_fmt = workbook.add_format({
            "bold": True,
            "font_size": 21,
            "font_color": "white",
            "bg_color": "#17365D",
            "align": "center",
            "valign": "vcenter"
        })

        section_fmt = workbook.add_format({
            "bold": True,
            "font_size": 12,
            "font_color": "white",
            "bg_color": "#17365D",
            "align": "left",
            "valign": "vcenter"
        })

        header_fmt = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#4472C4",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True
        })

        money_fmt = workbook.add_format({
            "num_format": "$#,##0.00"
        })

        count_fmt = workbook.add_format({
            "num_format": "#,##0"
        })

        percent_fmt = workbook.add_format({
            "num_format": "0.0%"
        })

        total_fmt = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#17365D",
            "num_format": "$#,##0.00"
        })

        # ----------------------------------------------------
        # DASHBOARD SHEET
        # ----------------------------------------------------

        ws = workbook.add_worksheet("Dashboard")

        ws.hide_gridlines(2)
        ws.set_zoom(85)
        ws.freeze_panes(9, 0)

        widths = {
            "A:A": 20,
            "B:B": 14,
            "C:C": 17,
            "D:D": 14,
            "E:E": 17,
            "F:F": 15,
            "G:G": 19,
            "H:H": 17
        }

        for col, width in widths.items():
            ws.set_column(col, width)

        ws.merge_range(
            "A1:H1",
            "PATIENT AR PROGRESS DASHBOARD",
            title_fmt
        )

        ws.set_row(0, 38)

        # ----------------------------------------------------
        # KPI CARDS
        # ----------------------------------------------------

        kpis_list = [
            (
                "TOTAL AR OUTSTANDING",
                kpis["total_outstanding"],
                "$#,##0.00"
            ),
            (
                "TOTAL GUARANTORS",
                kpis["total_guarantors"],
                "#,##0"
            ),
            (
                "2025 AR AMOUNT",
                kpis["total_2025_amount"],
                "$#,##0.00"
            ),
            (
                "2026 AR AMOUNT",
                kpis["total_2026_amount"],
                "$#,##0.00"
            )
        ]

        positions = [
            ("A4:B4", "A5:B6"),
            ("C4:D4", "C5:D6"),
            ("E4:F4", "E5:F6"),
            ("G4:H4", "G5:H6")
        ]

        colors = [
            "#17365D",
            "#1F4E78",
            "#548235",
            "#BF9000"
        ]

        for i, (label, value, number_format) in enumerate(kpis_list):

            label_range, value_range = positions[i]

            label_fmt = workbook.add_format({
                "bold": True,
                "font_color": "white",
                "bg_color": colors[i],
                "align": "center",
                "valign": "vcenter"
            })

            value_fmt = workbook.add_format({
                "bold": True,
                "font_size": 16,
                "font_color": colors[i],
                "bg_color": "#EAF0F6",
                "align": "center",
                "valign": "vcenter",
                "num_format": number_format
            })

            ws.merge_range(label_range, label, label_fmt)
            ws.merge_range(value_range, value, value_fmt)

        # ----------------------------------------------------
        # CATEGORY TABLE
        # ----------------------------------------------------

        ws.merge_range(
            "A8:H8",
            "CATEGORY-WISE AR PROGRESS",
            section_fmt
        )

        headers = [
            "Category",
            "2025 Claims",
            "2025 Amount",
            "2026 Claims",
            "2026 Amount",
            "Total Claims",
            "Outstanding AR",
            "AR Contribution"
        ]

        for col, header in enumerate(headers):
            ws.write(8, col, header, header_fmt)

        for row_num, (_, row) in enumerate(
            summary.iterrows(),
            start=9
        ):

            total_claims = (
                row["2025 Claims"] + row["2026 Claims"]
            )

            contribution = (
                row["Outstanding AR"] / kpis["total_outstanding"]
                if kpis["total_outstanding"] else 0
            )

            ws.write(row_num, 0, row["Category"])

            ws.write_number(
                row_num, 1, row["2025 Claims"], count_fmt
            )

            ws.write_number(
                row_num, 2, row["2025 Amount"], money_fmt
            )

            ws.write_number(
                row_num, 3, row["2026 Claims"], count_fmt
            )

            ws.write_number(
                row_num, 4, row["2026 Amount"], money_fmt
            )

            ws.write_number(
                row_num, 5, total_claims, count_fmt
            )

            ws.write_number(
                row_num, 6, row["Outstanding AR"], money_fmt
            )

            ws.write_number(
                row_num, 7, contribution, percent_fmt
            )

        # ----------------------------------------------------
        # GRAND TOTAL
        # ----------------------------------------------------

        total_row = 9 + len(summary)

        ws.write(total_row, 0, "GRAND TOTAL", total_fmt)

        ws.write_number(
            total_row, 1,
            kpis["total_2025_claims"], total_fmt
        )

        ws.write_number(
            total_row, 2,
            kpis["total_2025_amount"], total_fmt
        )

        ws.write_number(
            total_row, 3,
            kpis["total_2026_claims"], total_fmt
        )

        ws.write_number(
            total_row, 4,
            kpis["total_2026_amount"], total_fmt
        )

        ws.write_number(
            total_row, 5,
            kpis["total_claims"], total_fmt
        )

        ws.write_number(
            total_row, 6,
            kpis["total_outstanding"], total_fmt
        )

        ws.write_number(
            total_row, 7,
            1 if kpis["total_outstanding"] else 0,
            workbook.add_format({
                "bold": True,
                "font_color": "white",
                "bg_color": "#17365D",
                "num_format": "0.0%"
            })
        )

        # ----------------------------------------------------
        # CHART 1 - OUTSTANDING AR
        # ----------------------------------------------------

        first_data_row = 9
        last_data_row = 8 + len(summary)

        chart1 = workbook.add_chart({"type": "bar"})

        chart1.add_series({
            "name": "Outstanding AR",
            "categories": [
                "Dashboard",
                first_data_row, 0,
                last_data_row, 0
            ],
            "values": [
                "Dashboard",
                first_data_row, 6,
                last_data_row, 6
            ],
            "data_labels": {
                "value": True,
                "num_format": "$#,##0"
            }
        })

        chart1.set_title({
            "name": "Outstanding AR by Category"
        })

        chart1.set_x_axis({
            "name": "Outstanding Amount ($)"
        })

        chart1.set_y_axis({
            "name": "Category"
        })

        chart1.set_legend({"none": True})

        chart1.set_size({
            "width": 620,
            "height": 340
        })

        ws.insert_chart("A19", chart1)

        # ----------------------------------------------------
        # CHART 2 - CLAIM COMPARISON
        # ----------------------------------------------------

        chart2 = workbook.add_chart({"type": "column"})

        chart2.add_series({
            "name": "2025 Claims",
            "categories": [
                "Dashboard",
                first_data_row, 0,
                last_data_row, 0
            ],
            "values": [
                "Dashboard",
                first_data_row, 1,
                last_data_row, 1
            ]
        })

        chart2.add_series({
            "name": "2026 Claims",
            "categories": [
                "Dashboard",
                first_data_row, 0,
                last_data_row, 0
            ],
            "values": [
                "Dashboard",
                first_data_row, 3,
                last_data_row, 3
            ]
        })

        chart2.set_title({
            "name": "2025 vs 2026 Claim Comparison"
        })

        chart2.set_x_axis({
            "name": "Category"
        })

        chart2.set_y_axis({
            "name": "Claim Count"
        })

        chart2.set_size({
            "width": 620,
            "height": 340
        })

        ws.insert_chart("G19", chart2)

        # ----------------------------------------------------
        # CHART 3 - AR CONTRIBUTION
        # ----------------------------------------------------

        chart3 = workbook.add_chart({"type": "doughnut"})

        chart3.add_series({
            "name": "AR Contribution",
            "categories": [
                "Dashboard",
                first_data_row, 0,
                last_data_row, 0
            ],
            "values": [
                "Dashboard",
                first_data_row, 6,
                last_data_row, 6
            ],
            "data_labels": {
                "category": True,
                "percentage": True,
                "leader_lines": True
            }
        })

        chart3.set_title({
            "name": "Category-wise AR Contribution"
        })

        chart3.set_hole_size(55)

        chart3.set_size({
            "width": 620,
            "height": 340
        })

        ws.insert_chart("A38", chart3)

        # ----------------------------------------------------
        # ADDITIONAL EXCEL SHEETS
        # ----------------------------------------------------

        summary.to_excel(
            writer,
            sheet_name="Filtered Summary",
            index=False
        )

        raw_data.to_excel(
            writer,
            sheet_name="Raw Data",
            index=False
        )

        not_worked_data.to_excel(
            writer,
            sheet_name="Not Yet Worked",
            index=False
        )

        # ----------------------------------------------------
        # FORMAT ADDITIONAL SHEETS
        # ----------------------------------------------------

        for sheet_name in [
            "Filtered Summary",
            "Raw Data",
            "Not Yet Worked"
        ]:

            sheet = writer.sheets[sheet_name]

            sheet.freeze_panes(1, 0)
            sheet.set_default_row(18)

            if sheet_name == "Filtered Summary":
                rows, cols = summary.shape

            elif sheet_name == "Not Yet Worked":
                rows, cols = not_worked_data.shape

            else:
                rows, cols = raw_data.shape

            if cols > 0:
                sheet.autofilter(
                    0, 0,
                    max(rows, 1),
                    cols - 1
                )

            sheet.set_column(0, max(cols - 1, 0), 18)

    # ========================================================
    # COMPLETION MESSAGE
    # ========================================================

    print("\n" + "=" * 55)
    print("DASHBOARD CREATED SUCCESSFULLY!")
    print("=" * 55)

    print(f"\nOutput file: {OUTPUT_FILE}")

    print(
        f"Total Outstanding AR: "
        f"${kpis['total_outstanding']:,.2f}"
    )

    print(
        f"Total Guarantors: "
        f"{kpis['total_guarantors']:,.0f}"
    )

    print(
        f"Total Claims: "
        f"{kpis['total_claims']:,.0f}"
    )

    print("\nIncluded categories:")

    for category in CATEGORIES:
        print(f" - {category}")

    print("\nPaid in Full: EXCLUDED")
    print("\nOpen the generated Excel file to view your dashboard.")


# ============================================================
# MAIN PROGRAM
# ============================================================

if __name__ == "__main__":

    try:

        summary, raw_data, amount_column = load_data()

        final_summary, not_worked_data = prepare_summary(
            summary,
            raw_data,
            amount_column
        )

        kpis = calculate_kpis(final_summary)

        create_dashboard(
            final_summary,
            raw_data,
            not_worked_data,
            kpis
        )

    except Exception as error:

        print("\nERROR OCCURRED:")
        print(error)

        sys.exit(1)
