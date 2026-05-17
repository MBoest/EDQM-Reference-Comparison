from pathlib import Path
import datetime as dt
import pandas as pd
import openpyxl as ox
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
import requests
import xml.etree.ElementTree as ET


pathxl = Path(".") / "input_file" / "example_reference_substances.xlsx"
pathxml = Path(".") / "edqm_catalogue_destination" / "web_catalog_XML.xml"

url = "https://crs.edqm.eu/db/4DCGI/web_catalog_XML.xml"

worksheet_name = "Reference Substances"


def create_df(ws) -> pd.DataFrame:
    if not ws:
        raise ValueError("No worksheet passed")
    
    start_row = 3

    data = {
        "Substance Specific Code": [
            ws.cell(row=i, column=1).value for i in range(start_row, ws.max_row + 1)
        ],
        "Substance Received Year": [
            ws.cell(row=i, column=2).value for i in range(start_row, ws.max_row + 1)
        ],
        "Substance Received Index": [
            ws.cell(row=i, column=3).value for i in range(start_row, ws.max_row + 1)
        ],
        "Article No.": [
            ws.cell(row=i, column=5).value for i in range(start_row, ws.max_row + 1)
        ],
        "In Stock": [
            ws.cell(row=i, column=6).value for i in range(start_row, ws.max_row + 1)
        ],
        "Category": [
            ws.cell(row=i, column=7).value for i in range(start_row, ws.max_row + 1)
        ],
        "Supplier Batch No.": [
            ws.cell(row=i, column=8).value for i in range(start_row, ws.max_row + 1)
        ],
    }

    return pd.DataFrame(data)

def filter_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.sort_values(by="Substance Specific Code")

    df_crs_in_stock = df[
        df["In Stock"].str.contains("yes", case=False, na=False)
        & df["Category"].str.contains("CRS", case=False, na=False)
    ].copy()

    df_crs_in_stock["In Stock"] = df_crs_in_stock["In Stock"].str.strip()
    df_crs_in_stock["Category"] = df_crs_in_stock["Category"].str.strip()
    df_crs_in_stock["Article No."] = df_crs_in_stock["Article No."].str.strip()

    df_crs_in_stock["Supplier Batch No."] = (
        df_crs_in_stock["Supplier Batch No."]
        .astype("string")
        .str.replace(r"[A-Za-z]", "", regex=True)
        .str.lstrip("0")
        .str.split(".")
        .str[0]
    )

    return df_crs_in_stock

def download_edqm_xml(url: str) -> None:
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
    except requests.RequestException as e:
        raise ValueError(f"Download failed: {e}") from e

    with open(pathxml, "wb") as file:
        file.write(response.content)

def read_xml_file(pathxml: Path) -> pd.DataFrame:
    tree = ET.parse(pathxml)
    root = tree.getroot()

    edqm_references = []

    for ref in root.findall("Reference"):
        order_code = ref.get("Order_Code")
        batch_edqm = ref.find("Batch_No")
        if batch_edqm is None or batch_edqm.text is None:
          continue
        batch_edqm = batch_edqm.text

        edqm_references.append({
            "Article No.": order_code,
            "Current EDQM Batch": batch_edqm,
            })
    
    return pd.DataFrame(edqm_references)

def compare_batchnumbers(edqm_references: pd.DataFrame, own_references: pd.DataFrame) -> pd.DataFrame:
    shared_codes = pd.merge(own_references, edqm_references, on="Article No.", how="left",
    )

    return shared_codes[
        shared_codes["Supplier Batch No."] != shared_codes["Current EDQM Batch"]
    ]

def write_edqm_check_to_workbook(wb: ox.Workbook, standards_to_check_df: pd.DataFrame, file_path: Path) -> None:
    sheet_name = f"EDQM_check_{dt.date.today():%Y_%m_%d}"

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(title=sheet_name)

    df = standards_to_check_df.copy()

    df.columns = [str(col) for col in df.columns]

    if len(df.columns) != len(set(df.columns)):
        raise ValueError("DataFrame enthält doppelte Spaltennamen.")

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"

    max_row = ws.max_row
    max_col = ws.max_column

    if max_row < 2:
        wb.save(file_path)
        return

    table_ref = f"A1:{get_column_letter(max_col)}{max_row}"
    table_name = f"EDQMCheck_{dt.datetime.now():%Y%m%d_%H%M%S}"

    table = Table(
        displayName=table_name,
        ref=table_ref,
    )

    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    table.tableStyleInfo = style
    ws.add_table(table)

    for column_cells in ws.columns:
        max_len = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        ws.column_dimensions[column_cells[0].column_letter].width = min(max_len + 2, 40)

    wb.save(file_path)


if __name__ == "__main__":
    wb = ox.load_workbook(pathxl)

    ws = wb[worksheet_name]

    df = create_df(ws)

    if df.empty:
        raise ValueError("No DataFrame created")

    df_filtered = filter_df(df)

    download_edqm_xml(url)
    edqm_references_df = read_xml_file(pathxml)

    standards_to_check_df = compare_batchnumbers(edqm_references_df, df_filtered)

    write_edqm_check_to_workbook(
        wb=wb,
        standards_to_check_df=standards_to_check_df,
        file_path=pathxl,
    )