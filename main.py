import sys
import os
import logging
import requests
import win32com.client
import win32process
import win32gui
from contextlib import contextmanager
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLineEdit,
    QFileDialog,
    QLabel,
    QStatusBar,
)
from PyQt6.QtCore import QTimer
from config.config import API_BASE_URL, API_KEY

# --- Logging Setup ---
log_file_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "app.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(log_file_path), logging.StreamHandler(sys.stdout)],
)


class ExcelReadError(Exception):
    """Custom exception for recoverable Excel read errors."""

    pass


@contextmanager
def excel_manager():
    """A context manager to ensure the Excel process is properly shut down, even on error."""
    excel = None
    pid = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False  # Suppress prompts like the save dialog
        # Get the process ID to ensure termination
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        logging.info(f"Excel application dispatched with PID: {pid}.")
        yield excel
    finally:
        if excel:
            # Wrap graceful quit in a try/except. If the COM object is corrupted
            # by an upstream error, this will fail. We catch it and proceed to
            # the forceful termination, which is the most important step.
            try:
                excel.Quit()
                del excel
                logging.info("Graceful Excel quit command sent.")
            except Exception as e:
                logging.warning(
                    f"Failed to gracefully quit Excel (this is expected if an error occurred). Will now force terminate. Details: {e}"
                )

        if pid:
            try:
                # After attempting a graceful quit, forcefully terminate the process
                # to prevent lingering instances, especially after errors.
                kill_command = f"taskkill /F /PID {pid}"
                # Redirect output to null to keep the console clean
                result = os.system(f"{kill_command} > nul 2>&1")
                if result == 0:
                    logging.info(
                        f"Successfully terminated lingering Excel process with PID: {pid}."
                    )
                else:
                    # A non-zero return code usually means the process was already gone.
                    logging.info(
                        f"Excel process with PID {pid} was not found (likely already closed)."
                    )
            except Exception as e:
                logging.error(
                    f"An error occurred while trying to kill Excel process {pid}: {e}"
                )


def get_version():
    """Reads the version from version.txt"""
    try:
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        version_file_path = os.path.join(base_path, "version.txt")
        with open(version_file_path, "r") as f:
            return f.read().strip()

    except FileNotFoundError:
        return "0.0.0"


__version__ = get_version()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        logging.info("Application starting up.")
        self.setWindowTitle("Stock Data Scraper")
        self.setGeometry(100, 100, 500, 150)
        self.current_job_id = None
        self.api_headers = {"X-API-KEY": API_KEY}

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # UI Elements
        self.file_path_box = QLineEdit()
        self.file_path_box.setReadOnly(True)
        self.select_file_button = QPushButton("Select Excel File")
        self.select_file_button.clicked.connect(self.select_file)

        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_box)
        file_layout.addWidget(self.select_file_button)
        self.layout.addLayout(file_layout)

        self.run_button = QPushButton("Run Scraper")
        self.run_button.clicked.connect(self.run_full_process)
        self.layout.addWidget(self.run_button)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        version_label = QLabel(f"v{__version__}")
        self.status_bar.addPermanentWidget(version_label)
        self.status_bar.showMessage("Ready. Please select a file.")

        self.timer = QTimer(self)
        self.timer.setInterval(10000)
        self.timer.timeout.connect(self.check_job_status)
        logging.info("UI Initialized.")

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Macro-Enabled Workbook (*.xlsm)"
        )
        if file_path:
            logging.info(f"User selected file: {file_path}")
            self.file_path_box.setText(file_path)

    def run_full_process(self):
        file_path = self.file_path_box.text()
        logging.info(f"Starting full process for file: {file_path}")
        if not file_path or not os.path.exists(file_path):
            self.status_bar.showMessage("Please select a valid file first.")
            logging.warning("Run process attempted with no valid file selected.")
            return

        self.run_button.setEnabled(False)
        self.status_bar.showMessage("Step 1/4: Reading tickers from Excel...")
        QApplication.processEvents()

        try:
            stocks_tickers, mf_tickers = self._read_tickers_from_excel(file_path)
            if not stocks_tickers and not mf_tickers:
                self.status_bar.showMessage(
                    "No tickers found in the file. See app.log for details."
                )
                logging.warning("No tickers returned from _read_tickers_from_excel.")
                self.run_button.setEnabled(True)
                return

            self.status_bar.showMessage(
                f"Step 2/4: Found {len(stocks_tickers)} stocks and {len(mf_tickers)} MFs. Submitting job..."
            )

            payload = {"stocks_tickers": stocks_tickers, "mf_tickers": mf_tickers}

            url = f"{API_BASE_URL}/api/v1/stocks-rpa/submit-job"
            response = requests.post(url, headers=self.api_headers, json=payload)

            if response.status_code == 401:
                raise Exception("API Error: Unauthorized. Check API Key.")
            response.raise_for_status()

            self.current_job_id = response.json()["job_id"]
            self.status_bar.showMessage(
                f"Step 3/4: Job submitted (ID: {self.current_job_id[:8]}...). Waiting for results..."
            )
            logging.info(f"Job submitted successfully. Job ID: {self.current_job_id}")
            self.timer.start()

        except ExcelReadError as e:
            self.status_bar.showMessage(str(e))
            logging.error(f"A recoverable Excel error occurred: {e}", exc_info=False)
            self.run_button.setEnabled(True)

        except Exception as e:
            self.status_bar.showMessage(f"An error occurred. See app.log for details.")
            logging.error(
                f"An exception occurred in run_full_process: {e}", exc_info=True
            )
            self.run_button.setEnabled(True)

    def check_job_status(self):
        if not self.current_job_id:
            self.timer.stop()
            return

        try:
            url = f"{API_BASE_URL}/api/v1/job-status/{self.current_job_id}"
            response = requests.get(url, headers=self.api_headers)

            if response.status_code == 401:
                raise Exception("API Error: Unauthorized. Check API Key.")
            response.raise_for_status()

            data = response.json()
            logging.info(f"Received job status response: {data}")

            if data["status"] == "finished":
                self.timer.stop()
                self.status_bar.showMessage(
                    "Step 4/4: Job finished. Writing data to Excel..."
                )
                QApplication.processEvents()

                file_path = self.file_path_box.text()
                stocks_data = data["result"].get("stocks_data", {})
                mf_data = data["result"].get("mf_data", {})

                if stocks_data:
                    self._write_stocks_data_to_excel(file_path, stocks_data)
                if mf_data:
                    self._write_mf_data_to_excel(file_path, mf_data)

                self.status_bar.showMessage(
                    f"Success! Wrote data for {len(stocks_data)} stocks and {len(mf_data)} MFs."
                )
                self.current_job_id = None
                self.run_button.setEnabled(True)

            elif data["status"] == "failed":
                self.timer.stop()
                self.status_bar.showMessage(
                    "Job failed on remote worker. Check server logs."
                )
                self.current_job_id = None
                self.run_button.setEnabled(True)

        except requests.exceptions.RequestException as e:
            self.timer.stop()
            self.status_bar.showMessage(f"API Error while checking status: {e}")
            self.run_button.setEnabled(True)

    def _read_tickers_from_excel(self, file_path):
        logging.info(f"Attempting to read all data in one block from: {file_path}")
        with excel_manager() as excel:
            try:
                workbook = excel.Workbooks.Open(file_path)
                logging.info("Workbook opened.")

                worksheet = workbook.Sheets("screener")
                logging.info("Accessed 'screener' worksheet.")

                used_range = worksheet.UsedRange
                if used_range is None:
                    raise ExcelReadError(
                        "Failed to read Excel file due to a temporary issue. Please re-run the process."
                    )

                all_data = used_range.Value
                logging.info(
                    f"Bulk read complete using UsedRange. Read {len(all_data)} rows."
                )

                if not all_data:
                    logging.warning("UsedRange returned no data, or sheet is empty.")
                    return [], []

                stocks_tickers = []
                mf_tickers = []
                current_table = None

                for i, row_data in enumerate(all_data):
                    is_header = False
                    if not row_data or all(
                        c is None or str(c).strip() == "" for c in row_data
                    ):
                        continue

                    for cell in row_data:
                        cell_str = str(cell).strip().lower()
                        if cell_str.startswith("p/e"):
                            current_table = "stocks"
                            is_header = True
                            logging.info(
                                f"Found STOCKS header at sheet row {i+1}: {row_data}"
                            )
                            break
                        elif cell_str.startswith("expense ratio"):
                            current_table = "mf"
                            is_header = True
                            logging.info(
                                f"Found MF header at sheet row {i+1}: {row_data}"
                            )
                            break

                    if is_header:
                        continue

                    if len(row_data) < 2:
                        continue

                    ticker = row_data[1]
                    if ticker:
                        if current_table == "stocks":
                            stocks_tickers.append(str(ticker))
                        elif current_table == "mf":
                            mf_tickers.append(str(ticker))

                logging.info(
                    f"Found {len(stocks_tickers)} stock tickers: {stocks_tickers}"
                )
                logging.info(f"Found {len(mf_tickers)} MF tickers: {mf_tickers}")
                return stocks_tickers, mf_tickers

            except Exception as e:
                logging.error(
                    f"An error occurred in _read_tickers_from_excel: {e}", exc_info=True
                )
                raise

    def _write_stocks_data_to_excel(self, file_path, data):
        logging.info(f"Attempting to write P/E data to: {file_path}")
        with excel_manager() as excel:
            try:
                workbook = excel.Workbooks.Open(file_path)
                worksheet = workbook.Sheets("screener")
                logging.info("Accessed 'screener' worksheet for writing.")

                header_row_num = None
                pe_col = None
                ticker_col = 2

                try:
                    pe_header_cell = worksheet.UsedRange.Find("P/E", LookAt=2)
                    if pe_header_cell is None:
                        raise Exception("P/E header not found")
                    header_row_num = pe_header_cell.Row
                    pe_col = pe_header_cell.Column
                except Exception:
                    raise Exception(
                        "Could not find a column starting with 'P/E' in the 'screener' sheet."
                    )

                logging.info(
                    f"Found P/E headers in row {header_row_num}: Ticker in col {ticker_col}, P/E in col {pe_col}"
                )

                first_data_row = header_row_num + 1
                last_stock_row = 0
                for row in range(first_data_row, worksheet.Rows.Count):
                    if not worksheet.Cells(row, ticker_col).Value:
                        last_stock_row = row - 1
                        break
                if last_stock_row == 0:
                    last_stock_row = (
                        worksheet.Cells(worksheet.Rows.Count, ticker_col).End(-4162).Row
                    )

                logging.info(
                    f"Stock table range identified: Rows {first_data_row} to {last_stock_row}"
                )

                pe_data_to_write = []
                for row in range(first_data_row, last_stock_row + 1):
                    ticker_in_cell = str(worksheet.Cells(row, ticker_col).Value)
                    if ticker_in_cell in data:
                        pe_value = data[ticker_in_cell].get("P/E", "N/A")
                        pe_data_to_write.append([pe_value])
                    else:
                        pe_data_to_write.append([worksheet.Cells(row, pe_col).Value])

                target_range = worksheet.Range(
                    worksheet.Cells(first_data_row, pe_col),
                    worksheet.Cells(last_stock_row, pe_col),
                )
                logging.info(
                    f"Preparing to write {len(pe_data_to_write)} values to P/E column range {target_range.Address}."
                )
                target_range.Value = pe_data_to_write
                logging.info("Bulk write of P/E column complete.")

                workbook.Save()
                logging.info("Workbook saved.")

            except Exception as e:
                logging.error(
                    f"An error occurred in _write_stocks_data_to_excel: {e}",
                    exc_info=True,
                )
                raise

    def _write_mf_data_to_excel(self, file_path, data):
        logging.info(f"Attempting to write Fund Size data to: {file_path}")
        with excel_manager() as excel:
            try:
                workbook = excel.Workbooks.Open(file_path)
                worksheet = workbook.Sheets("screener")
                logging.info("Accessed 'screener' worksheet for MF writing.")

                header_row_num = None
                fund_size_col = None
                ticker_col = 2

                try:
                    # Find the MF table by looking for its unique header
                    mf_header_cell = worksheet.UsedRange.Find("Expense Ratio", LookAt=2)
                    if mf_header_cell is None:
                        raise Exception("MF 'Expense Ratio' header not found")
                    header_row_num = mf_header_cell.Row

                    # Now find the 'Fund Size' column within that header row
                    header_range = worksheet.Rows(header_row_num)
                    fund_size_cell = header_range.Find("Fund Size", LookAt=2)
                    if fund_size_cell is None:
                        raise Exception("'Fund Size' header not found in MF table")
                    fund_size_col = fund_size_cell.Column

                except Exception as e:
                    raise Exception(f"Could not find MF headers: {e}")

                logging.info(
                    f"Found MF headers in row {header_row_num}: Ticker in col {ticker_col}, Fund Size in col {fund_size_col}"
                )

                first_data_row = header_row_num + 1
                last_mf_row = (
                    worksheet.Cells(worksheet.Rows.Count, ticker_col).End(-4162).Row
                )

                logging.info(
                    f"MF table range identified: Rows {first_data_row} to {last_mf_row}"
                )

                fund_size_data_to_write = []
                for row in range(first_data_row, last_mf_row + 1):
                    ticker_in_cell = str(worksheet.Cells(row, ticker_col).Value)
                    if ticker_in_cell in data:
                        fund_size_value = data[ticker_in_cell].get("Fund Size", "N/A")
                        fund_size_data_to_write.append([fund_size_value])
                    else:
                        fund_size_data_to_write.append(
                            [worksheet.Cells(row, fund_size_col).Value]
                        )

                target_range = worksheet.Range(
                    worksheet.Cells(first_data_row, fund_size_col),
                    worksheet.Cells(last_mf_row, fund_size_col),
                )
                logging.info(
                    f"Preparing to write {len(fund_size_data_to_write)} values to Fund Size column range {target_range.Address}."
                )
                target_range.Value = fund_size_data_to_write
                logging.info("Bulk write of Fund Size column complete.")

                workbook.Save()
                logging.info("Workbook saved.")

            except Exception as e:
                logging.error(
                    f"An error occurred in _write_mf_data_to_excel: {e}", exc_info=True
                )
                raise


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec())

    except Exception as e:
        logging.critical(f"A critical error occurred on startup: {e}", exc_info=True)
        sys.exit(1)
