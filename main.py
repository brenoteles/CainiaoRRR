import os
import random
import datetime
import string
import sqlite3
import openpyxl
from datetime import datetime  # Import datetime for date handling
from pathlib import Path
import win32com.client

from PyQt5.QtWidgets import QApplication, QMainWindow, QListWidgetItem, QMessageBox
from PyQt5.uic import loadUi
import sys
from qtpy import QtCore

CONST_DBNAME = "rrr.db"
CONST_TEMPLATE_RRR = Path(r'C:\Users\aguia\PycharmProjects\CainiaoRRR\modelo.xlsm')
CONST_OUTPUT_RRR = Path(r'C:\Users\aguia\PycharmProjects\CainiaoRRR')
CONST_EXCEL_SHEET_NAME = 'RRR'

def create_rrr_table(db_name="rrr.db"):
    """
    Creates the table RRR in the specified database if it does not exist.

    Args:
        db_name (str): The name of the SQLite database file.
    """
    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()

        # Execute a CREATE TABLE IF NOT EXISTS statement
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS RRR (
                id INTEGER PRIMARY KEY,
                hawb TEXT,
                awb TEXT,
                dir TEXT,
                nome TEXT,
                cpf TEXT,
                peso TEXT,
                valor TEXT,
                status TEXT,
                manifesto TEXT,
                data TEXT
            )
        """)
        conn.commit()
        print(f"Table 'RRR' created successfully (or already existed) in database '{db_name}'.")

    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the database connection
        if conn:
            conn.close()

def generate_random_data(num_rows=10):
    """
    Generates a list of tuples, where each tuple represents a row of random data
    for the RRR table.

    Args:
        num_rows (int): The number of rows of random data to generate.  Defaults to 10.

    Returns:
        list: A list of tuples, where each tuple contains the random data for one row.
    """
    data = []
    for _ in range(num_rows):
        hawb = ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))
        awb = ''.join(random.choices(string.ascii_uppercase + string.digits, k=10))
        dir_val = ''.join(random.choices(string.ascii_uppercase + string.digits, k=5)) if random.random() < 0.8 else None  # 80% chance of not being None
        nome = ''.join(random.choices(string.ascii_letters, k=20)) if random.random() < 0.8 else None
        cpf = ''.join(random.choices(string.digits, k=11)) if random.random() < 0.8 else None
        peso = random.uniform(0, 100) if random.random() < 0.8 else None
        valor = random.uniform(10, 1000) if random.random() < 0.8 else None
        status = random.choice(["aguardando", "transito", "entregue", "problema"])
        manifesto = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
        date_val = datetime.date(random.randint(2023, 2024), random.randint(1, 12), random.randint(1, 28)).isoformat()
        data.append((hawb, awb, dir_val, nome, cpf, peso, valor, status, manifesto, date_val)) #changed data to data_val
    return data

def insert_rrr_data(db_name="rrr.db", data=None):
    """
    Inserts data into the RRR table.

    Args:
        db_name (str): The name of the SQLite database file.
        data (list): A list of tuples, where each tuple represents a row of data to insert.
    """
    if data is None:
        data = generate_random_data()  # Generate default data if none is provided.

    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()

        # Use executemany to insert multiple rows at once
        cursor.executemany("""
            INSERT INTO RRR (hawb, awb, dir, nome, cpf, peso, valor, status, manifesto, data)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, data)
        conn.commit()
        print(f"{len(data)} rows of data inserted into table 'RRR' in database '{db_name}'.")

    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the database connection
        if conn:
            conn.close()

def print_rrr_table(db_name="rrr.db", table_name="RRR"):
    """
    Opens an SQLite database, connects to a specified table,
    and prints all the data in that table.

    Args:
        db_name (str): The name of the SQLite database file.
        table_name (str): The name of the table to print data from.
    """
    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(db_name)
        cursor = conn.cursor()

        # Execute a SELECT query to retrieve all data from the table
        cursor.execute(f"SELECT * FROM {table_name}")

        # Fetch all the rows returned by the query
        rows = cursor.fetchall()

        # Print the header (column names) if the table is not empty
        if rows:
            column_names = [description[0] for description in cursor.description]
            print("| " + " | ".join(column_names) + " |")
            print("-" * (3 * len(column_names) + 5 * (len(column_names) -1))) # Print a separator

            # Print each row of data
            for row in rows:
                print("| " + " | ".join(map(str, row)) + " |")
        else:
            print(f"Table '{table_name}' is empty.")

    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the database connection
        if conn:
            conn.close()

def read_and_insert_excel_fast(file_path):
    """
    Reads an Excel file using openpyxl and inserts data into the RRR table in an SQLite database efficiently.

    Args:
        file_path (str): The path to the Excel file.
    """
    try:
        # Load the workbook in read-only mode for efficiency
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet = workbook.active

        # Connect to the SQLite database
        with sqlite3.connect(CONST_DBNAME) as conn:
            cursor = conn.cursor()

            # Prepare the SQL INSERT statement with placeholders
            sql = """
                INSERT INTO RRR (hawb, awb, dir, nome, cpf, peso, valor, status, manifesto, data)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """

            data_to_insert = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[1]:  # Assuming 'hawb' in the second column, break if empty
                    break

                hawb = row[1]
                awb = row[0] if len(row) > 0 else None
                dir_ = row[2] if len(row) > 2 else None
                nome = row[4] if len(row) > 4 else None
                cpf = row[3] if len(row) > 3 else None
                peso = row[5] if len(row) > 5 else None
                valor = row[6] if len(row) > 6 else None
                status = row[7] if len(row) > 7 else None
                manifesto = row[8] if len(row) > 8 else None
                data = row[9] if len(row) > 9 else None

                # Convert date to ISO format if it's a datetime object
                if isinstance(data, datetime):
                    data = data.isoformat()

                data_to_insert.append((hawb, awb, dir_, nome, cpf, peso, valor, status, manifesto, data))

            # Execute many inserts at once for efficiency
            cursor.executemany(sql, data_to_insert)
            conn.commit()
            print(f"{cursor.rowcount} rows inserted successfully.")

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
    except openpyxl.Error as excel_error:
        print(f"Error reading Excel file: {excel_error}")
    except sqlite3.Error as db_error:
        print(f"Database error: {db_error}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def fetch_parcels_by_awb(awb_number: int) -> list:
    try:
        with sqlite3.connect(CONST_DBNAME) as conn:
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
    excel_app = None
    workbook = None
    output_file_path = output_directory / f"CAINIAO AWB {mawb[:3]}-{mawb[3:]}.xlsm"
    print("OKAY")
    try:
        # Check if the output file already exists and delete it
        if os.path.exists(output_file_path):
            os.remove(output_file_path)
            print(f"Previous file '{output_file_path}' deleted.")

        excel_app = win32com.client.Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(excel_template_path)
        sheet = workbook.Sheets(sheet_name)

        first_parcel = parcel_data[0]

        # Write static headers
        sheet.Cells(11, 5).Value = first_parcel[7]  # MANIFESTO
        sheet.Cells(12, 5).Value = ("AWB ", first_parcel[9])  # NUMERO VOO
        sheet.Cells(11, 10).Value = first_parcel[8]  # DATA
        sheet.Cells(12, 10).Value = first_parcel[8]  # DATA

        if parcel_data:
            # Prepare data for writing (adjust column order as needed)
            data_to_write = [
                [
                    parcel[0],  # hawb (Column B)
                    parcel[1],  # dir (Column C)
                    parcel[2],  # cpf (Column D)
                    parcel[3],  # nome (Column E)
                    parcel[4],  # peso (Column F)
                    parcel[5],  # valor (Column G)
                    '1',  # Column H (Empty)
                    '1',  # Column I (Empty)
                    '',  # manifesto (Column J)
                    '',  # data (Column K)
                    '',  # Column L (Empty)
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


class Main(QMainWindow):
    def __init__(self):
        super(Main, self).__init__()
        loadUi("main.ui", self)
        self.update_list()
        self.toggleAll.clicked.connect(self.print_checked_items)  # connect new button

    def update_list(self):
        try:
            with sqlite3.connect(CONST_DBNAME) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT DISTINCT awb FROM RRR;")
                rows = cursor.fetchall()

            mawb_list = [row[0] for row in rows] if rows else []

            self.listWidget.clear()
            for mawb in mawb_list:
                item = QListWidgetItem(mawb)
                item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
                item.setCheckState(QtCore.Qt.Unchecked)
                self.listWidget.addItem(item)

            if not mawb_list:
                print("No AWB data found.")

        except sqlite3.Error as e:
            print(f"An error occurred: {e}")

    def print_checked_items(self):
        checked_items = []
        for index in range(self.listWidget.count()):
            item = self.listWidget.item(index)
            if item.checkState() == QtCore.Qt.Checked:
                checked_items.append(item.text())

        if checked_items:
            for mawb in checked_items:
                mawb_number = int(mawb)
                print(f"\nProcessing AWB: {mawb}")
                parcel_list = fetch_parcels_by_awb(mawb_number)

                if parcel_list:
                    print("Parcels retrieved:")
                    for parcel in parcel_list:
                        print(parcel)
                    write_parcel_data_to_excel(
                        str(CONST_TEMPLATE_RRR), CONST_OUTPUT_RRR, CONST_EXCEL_SHEET_NAME, parcel_list, mawb
                    )
                else:
                    print(f"Could not retrieve parcels for AWB: {mawb}")
        else:
            print("No items are checked.")
            QMessageBox.information(self, "No Items Checked", "No items are checked in the list.")  # optional

if __name__ == "__main__":
    # Example usage:
    # 1. Create the table if it doesn't exist
    create_rrr_table()
    # 2. Generate 10 rows of random data
    #random_data = generate_random_data(10)
    # 3. Insert the random data into the table
    #insert_rrr_data(data=random_data)
    # 4. Print the data from the table

    #file_path = "Consolidado de RRRs(1).xlsx"  # Replace with the actual path to your file
    #read_and_insert_excel_fast(file_path)
    #print_rrr_table()

    app = QApplication(sys.argv)
    window = Main()
    window.show()
    app.exec_()


