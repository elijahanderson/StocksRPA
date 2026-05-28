import sys
import os
import logging
import requests
import win32com.client
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
        self.timer.setInterval(3000)
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
            tickers = self._read_tickers_from_excel(file_path)
            if not tickers:
                self.status_bar.showMessage(
                    "No tickers found in the file. See app.log for details."
                )
                logging.warning("No tickers returned from _read_tickers_from_excel.")
                self.run_button.setEnabled(True)
                return

            self.status_bar.showMessage(
                f"Step 2/4: Found {len(tickers)} tickers. Submitting job to API..."
            )
            url = f"{API_BASE_URL}/api/v1/stocks-rpa/submit-job"
            response = requests.post(
                url, headers=self.api_headers, json={"tickers": tickers}
            )

            if response.status_code == 401:
                raise Exception("API Error: Unauthorized. Check API Key.")
            response.raise_for_status()

            self.current_job_id = response.json()["job_id"]
            self.status_bar.showMessage(
                f"Step 3/4: Job submitted (ID: {self.current_job_id[:8]}...). Waiting for results..."
            )
            logging.info(f"Job submitted successfully. Job ID: {self.current_job_id}")
            self.timer.start()

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
                self._write_data_to_excel(file_path, data["result"])

                self.status_bar.showMessage(
                    f"Success! Wrote data for {len(data['result'])} tickers."
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
        logging.info(f"Attempting to read tickers in bulk from: {file_path}")
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            logging.info("Excel application dispatched.")

            workbook = excel.Workbooks.Open(file_path)
            logging.info("Workbook opened.")

            worksheet = workbook.Sheets("screener")
            logging.info("Accessed 'screener' worksheet.")

            last_row = worksheet.Cells(worksheet.Rows.Count, 2).End(-4162).Row
            logging.info(f"Found last row with data in column B at row: {last_row}")

            # --- Read the entire range in one operation ---
            ticker_range = worksheet.Range(f"B2:B{last_row}")
            data = ticker_range.Value
            logging.info(f"Bulk read complete. Read {len(data)} rows.")

            # Process the data in Python
            tickers = [str(row[0]) for row in data if row[0] is not None]

            workbook.Close(SaveChanges=False)
            logging.info(f"Found {len(tickers)} tickers: {tickers}")
            return tickers
        except Exception as e:
            logging.error(
                f"An error occurred in _read_tickers_from_excel: {e}", exc_info=True
            )
            raise
        finally:
            if excel:
                excel.Quit()
            logging.info("Excel application quit.")

    def _write_data_to_excel(self, file_path, data):
        logging.info(f"Attempting to write data in bulk to: {file_path}")
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            logging.info("Excel application dispatched for writing.")

            workbook = excel.Workbooks.Open(file_path)
            logging.info("Workbook opened for writing.")

            worksheet = workbook.Sheets("screener")
            logging.info("Accessed 'screener' worksheet for writing.")

            start_row = 80

            data_to_write = [["Scraped Ticker", "Scraped Price"]]
            for ticker, details in data.items():
                price = details.get("price", "N/A")
                data_to_write.append([ticker, price])

            num_rows = len(data_to_write)
            num_cols = len(data_to_write[0])

            end_row = start_row + num_rows - 1
            end_col = num_cols
            target_range = worksheet.Range(
                worksheet.Cells(start_row, 1), worksheet.Cells(end_row, end_col)
            )

            logging.info(
                f"Preparing to write {num_rows}x{num_cols} block of data to range {target_range.Address}."
            )

            target_range.Value = data_to_write

            logging.info("Bulk write operation complete.")

            workbook.Save()
            logging.info("Workbook saved.")
            workbook.Close(SaveChanges=True)
        except Exception as e:
            logging.error(
                f"An error occurred in _write_data_to_excel: {e}", exc_info=True
            )
            raise
        finally:
            if excel:
                excel.Quit()
            logging.info("Excel application quit after writing.")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logging.critical(f"A critical error occurred on startup: {e}", exc_info=True)
        sys.exit(1)
