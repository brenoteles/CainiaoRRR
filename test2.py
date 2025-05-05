import sqlite3
import win32com.client
from pathlib import Path

def fetch_parcels_by_awb(db_file: str, awb_number: int) -> list:
    """
    Retrieves parcel data for a specific AWB from a SQLite database.

    Args:
        db_file: Path to the SQLite database file.
        awb_number: The AWB number to filter by.

    Returns:
        A list of parcel data (each item is a tuple representing a row),
        or an empty list if an error occurs or no parcels are found.
    """
    try:
        with sqlite3.connect(db_file) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT hawb, dir, cpf, nome, peso, valor, status, manifesto, data, awb "
                "FROM RRR "
                "WHERE awb = ?;",
                (awb_number,),
            )
            parcels = cursor.fetchall()
        return parcels
    except sqlite3.Error as e:
        print(f"Database error occurred while fetching AWB {awb_number}: {e}")
        return []

def write_parcel_data_to_excel(excel_template_path: str, output_directory: Path, sheet_name: str, parcel_data: list, mawb: int):
    """
    Writes parcel data to a specified Excel sheet and saves the file with the MAWB number in the filename.

    Args:
        excel_template_path: Path to the Excel template file (.xlsm).
        output_directory: Path to the directory where the new Excel file will be saved.
        sheet_name: The name of the sheet to write data to.
        parcel_data: A list of parcel data (list of lists or tuples).
        mawb: The Master Air Waybill number being processed.
    """
    excel_app = None
    workbook = None
    output_file_path = output_directory / f"CAINIAO AWB {str(mawb)[:3]}-{str(mawb)[3:]}.xlsm"

    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(excel_template_path)
        sheet = workbook.Sheets(sheet_name)

        if parcel_data:
            # Write static headers using the first parcel's data
            first_parcel = parcel_data[0]
            sheet.Cells(11, 5).Value = first_parcel[7]  # MANIFESTO
            sheet.Cells(12, 5).Value = f"AWB {first_parcel[9]}"  # NUMERO VOO
            sheet.Cells(11, 10).Value = first_parcel[8]  # DATA
            sheet.Cells(12, 10).Value = first_parcel[8]  # DATA

            # Prepare data for writing (adjust column order as needed)
            data_to_write = [
                [
                    parcel[0],  # hawb (Column B)
                    parcel[1],  # dir (Column C)
                    parcel[2],  # cpf (Column D)
                    parcel[3],  # nome (Column E)
                    parcel[4],  # peso (Column F)
                    parcel[5],  # valor (Column G)
                    '',         # Column H (Empty)
                    '',         # Column I (Empty)
                    '',  # manifesto (Column J)
                    '',  # data (Column K)
                    '',         # Column L (Empty)
                    parcel[6],  # status (Column M - motivo)
                ]
                for parcel in parcel_data
            ]

            # Determine the data range
            start_row = 16
            start_col = 2
            num_rows = len(data_to_write)
            num_cols = 12  # Writing to columns B to M

            # Get the range and write the data
            data_range = sheet.Range(
                sheet.Cells(start_row, start_col),
                sheet.Cells(start_row + num_rows - 1, start_col + num_cols - 1),
            )
            data_range.Value = data_to_write
            print(f"Wrote {num_rows} rows (Columns B to M) for MAWB {mawb} to '{sheet_name}'.")
            column_to_autofit = sheet.Columns(13)
            column_to_autofit.AutoFit()
        else:
            print(f"No parcel data to write for MAWB {mawb} to '{sheet_name}'.")

        # Save the workbook with the MAWB in the filename
        workbook.SaveAs(str(output_file_path))
        print(f'Saved as "{output_file_path}".')

    except Exception as e:
        print(f"An error occurred while working with Excel for MAWB {mawb}: {e}")
    finally:
        # Ensure Excel objects are properly closed
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel_app:
            excel_app.Quit()

if __name__ == "__main__":
    mawbs = ['36993991914',
'36993994736',
'36993996206',
]
    db_file = 'rrr.db'  # Replace with the actual path to your database
    template_file = Path(r'C:\Users\aguia\PycharmProjects\CainiaoRRR\modelo.xlsm')
    output_directory = Path(r'C:\Users\aguia\PycharmProjects\CainiaoRRR')
    excel_sheet_name = 'RRR'

    for mawb in mawbs:
        print(f"\nProcessing AWB: {mawb}")
        parcel_list = fetch_parcels_by_awb(db_file, mawb)

        if parcel_list:
            print("Parcels retrieved:")
            for parcel in parcel_list:
                print(parcel)
            write_parcel_data_to_excel(
                str(template_file), output_directory, excel_sheet_name, parcel_list, mawb
            )
        else:
            print(f"Could not retrieve parcels for AWB: {mawb}")

    print("\nProcessing complete.")