import os
import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QFileDialog, QMessageBox


def process_excel(src, dst):
    if not os.path.exists(src):
        raise FileNotFoundError(f"Source file '{src}' not found.")

    # Read the first two columns (SKU, Description)
    df = pd.read_excel(src, engine="openpyxl", dtype=str)

    # Ensure at least two columns exist
    if df.shape[1] < 2:
        raise ValueError("Expected at least two columns (SKU and Description) in the Excel file.")

    # Normalize column names to work with positional access
    cols = list(df.columns)
    sku_col = cols[0]
    desc_col = cols[1]

    new_rows = []

    # Iterate by index so we can detect contiguous runs of the same SKU
    for i in range(len(df)):
        row = df.iloc[i]
        sku = row.get(sku_col)
        desc = row.get(desc_col)

        # Keep original row as-is
        new_rows.append(row.tolist())

        if pd.isna(sku) or str(sku).strip() == "":
            continue

        sku_key = str(sku).strip()

        # Determine next SKU (if any) to see if this is end of a contiguous block
        is_end_of_block = False
        if i == len(df) - 1:
            is_end_of_block = True
        else:
            next_sku = df.iloc[i + 1].get(sku_col)
            if pd.isna(next_sku) or str(next_sku).strip() != sku_key:
                is_end_of_block = True

        if not is_end_of_block:
            continue

        base_desc = "" if pd.isna(desc) else str(desc)

        # Append three rows after the end of the contiguous SKU block
        new_rows.append([f"{sku_key} FMOH", f"{base_desc} Fixed Manufacturing Overheads"] + [""] * (df.shape[1] - 2))
        new_rows.append([f"{sku_key} VMOH", f"{base_desc} Variable Manufacturing Overheads"] + [""] * (df.shape[1] - 2))
        new_rows.append([f"{sku_key} NDOH", f"{base_desc} Depreciation"] + [""] * (df.shape[1] - 2))

    # Build new DataFrame preserving other columns (if any)
    new_df = pd.DataFrame(new_rows, columns=cols)

    new_df.to_excel(dst, index=False, engine="openpyxl")
    return f"Wrote modified file to '{dst}'"


class ExcelModifier(QWidget):
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.output_file = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel SKU Modifier')
        self.setGeometry(300, 300, 400, 200)

        layout = QVBoxLayout()

        self.input_label = QLabel('Input File: Not selected')
        layout.addWidget(self.input_label)

        self.select_input_btn = QPushButton('Select Input Excel File')
        self.select_input_btn.clicked.connect(self.select_input)
        layout.addWidget(self.select_input_btn)

        self.output_label = QLabel('Output File: Not selected')
        layout.addWidget(self.output_label)

        self.select_output_btn = QPushButton('Select Output Excel File')
        self.select_output_btn.clicked.connect(self.select_output)
        layout.addWidget(self.select_output_btn)

        self.process_btn = QPushButton('Process and Save')
        self.process_btn.clicked.connect(self.process)
        layout.addWidget(self.process_btn)

        self.status_label = QLabel('')
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def select_input(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select Input Excel File', '', 'Excel Files (*.xlsx)')
        if file_name:
            self.input_file = file_name
            self.input_label.setText(f'Input File: {os.path.basename(file_name)}')

    def select_output(self):
        file_name, _ = QFileDialog.getSaveFileName(self, 'Select Output Excel File', '', 'Excel Files (*.xlsx)')
        if file_name:
            self.output_file = file_name
            self.output_label.setText(f'Output File: {os.path.basename(file_name)}')

    def process(self):
        if not self.input_file:
            QMessageBox.warning(self, 'Error', 'Please select an input file.')
            return
        if not self.output_file:
            QMessageBox.warning(self, 'Error', 'Please select an output file.')
            return

        try:
            message = process_excel(self.input_file, self.output_file)
            self.status_label.setText(message)
            QMessageBox.information(self, 'Success', message)
        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelModifier()
    window.show()
    sys.exit(app.exec_())

