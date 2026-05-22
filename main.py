import sys
import os
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

def get_version():
    """Reads the version from version.txt"""
    try:
        # Get the path to the directory where the script is located
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        version_file_path = os.path.join(base_path, 'version.txt')
        with open(version_file_path, "r") as f:
            return f.read().strip()
    except FileNotFoundError:
        return "0.0.0"

__version__ = get_version()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Macro Processor")
        self.setGeometry(100, 100, 500, 150)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        # File selection
        self.file_layout = QHBoxLayout()
        self.file_path_box = QLineEdit()
        self.file_path_box.setReadOnly(True)
        self.select_file_button = QPushButton("Select File")
        self.select_file_button.clicked.connect(self.select_file)
        self.file_layout.addWidget(self.file_path_box)
        self.file_layout.addWidget(self.select_file_button)
        self.layout.addLayout(self.file_layout)

        # Run button
        self.run_button = QPushButton("Run")
        self.run_button.clicked.connect(self.run_process)
        self.layout.addWidget(self.run_button)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

        # Version label
        version_label = QLabel(f"v{__version__}")
        self.status_bar.addPermanentWidget(version_label)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Macro-Enabled Workbook (*.xlsm)"
        )
        if file_path:
            self.file_path_box.setText(file_path)

    def run_process(self):
        file_path = self.file_path_box.text()
        if not file_path or not os.path.exists(file_path):
            self.status_bar.showMessage("Please select a valid file.")
            return

        self.status_bar.showMessage("Processing...")
        QApplication.processEvents()  # Update the GUI

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(file_path)
            worksheet = workbook.Sheets("screener")
            worksheet.Range("A1").Value = "test"
            workbook.Save()
            workbook.Close(SaveChanges=True)
            excel.Quit()
            self.status_bar.showMessage("Complete.")
        except Exception as e:
            self.status_bar.showMessage(f"Error: {e}")
        finally:
            # Ensure Excel is closed
            if 'excel' in locals() and excel is not None:
                excel.Quit()
                del excel


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
