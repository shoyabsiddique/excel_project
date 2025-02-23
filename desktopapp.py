# import sys
# import pandas as pd
# import numpy as np
# from PyQt5.QtWidgets import (
#     QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, 
#     QPushButton, QFileDialog, QComboBox, QWidget, QTableWidget, 
#     QTableWidgetItem, QMessageBox
# )
# from PyQt5.QtCore import Qt
# import openpyxl
# from openpyxl.styles import Font, Alignment

# class ExcelProcessorApp(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Distributor Excel Processor")
#         self.setGeometry(100, 100, 800, 600)

#         # Expected columns with their descriptions
#         self.EXPECTED_COLUMNS = {
#             'MARK': 'Product Mark/Type',
#             'CTN_NO': 'Carton Number',
#             'DESCRIPTION': 'Product Description',
#             'PCS_TOTAL': 'Total Pieces',
#             'UNITS': 'Unit of Measurement', 
#             'PCS_CTN': 'Pieces per Carton',
#             'CTN_TOTAL': 'Total Cartons',
#             'CBM_TOTAL': 'Total Cubic Meters',
#             'WEIGHT_TOTAL': 'Total Weight'
#         }

#         self.initUI()
        
#     def initUI(self):
#         # Main widget and layout
#         main_widget = QWidget()
#         main_layout = QVBoxLayout()
        
#         # File selection section
#         file_layout = QHBoxLayout()
#         self.file_path_label = QLabel("No file selected")
#         select_file_btn = QPushButton("Select Excel File")
#         select_file_btn.clicked.connect(self.select_excel_file)
        
#         file_layout.addWidget(self.file_path_label)
#         file_layout.addWidget(select_file_btn)
        
#         # Column mapping section
#         self.column_mapping_table = QTableWidget()
#         self.column_mapping_table.setColumnCount(3)
#         self.column_mapping_table.setHorizontalHeaderLabels([
#             "Original Column", "Map To", "Description"
#         ])
        
#         # Process button
#         process_btn = QPushButton("Process Excel")
#         process_btn.clicked.connect(self.process_excel)
        
#         # Add widgets to main layout
#         main_layout.addLayout(file_layout)
#         main_layout.addWidget(self.column_mapping_table)
#         main_layout.addWidget(process_btn)
        
#         main_widget.setLayout(main_layout)
#         self.setCentralWidget(main_widget)
        
#     def select_excel_file(self):
#         file_path, _ = QFileDialog.getOpenFileName(
#             self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
#         )
#         if file_path:
#             self.file_path_label.setText(file_path)
#             self.populate_column_mapping(file_path)
        
#     def populate_column_mapping(self, file_path):
#         try:
#             # Read Excel file headers
#             df = pd.read_excel(file_path, nrows=0)
#             original_columns = list(df.columns)
            
#             # Reset table
#             self.column_mapping_table.setRowCount(len(original_columns))
            
#             for row, col in enumerate(original_columns):
#                 # Original column name
#                 original_col_item = QTableWidgetItem(col)
#                 original_col_item.setFlags(original_col_item.flags() & ~Qt.ItemIsEditable)
#                 self.column_mapping_table.setItem(row, 0, original_col_item)
                
#                 # Mapping dropdown
#                 map_combo = QComboBox()
#                 map_combo.addItems(list(self.EXPECTED_COLUMNS.keys()))
                
#                 # Try to auto-match columns
#                 matching_keys = [
#                     key for key, desc in self.EXPECTED_COLUMNS.items() 
#                     if col.lower() in desc.lower() or desc.lower() in col.lower()
#                 ]
#                 if matching_keys:
#                     map_combo.setCurrentText(matching_keys[0])
                
#                 self.column_mapping_table.setCellWidget(row, 1, map_combo)
                
#                 # Description
#                 desc_item = QTableWidgetItem(
#                     self.EXPECTED_COLUMNS.get(map_combo.currentText(), "")
#                 )
#                 desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
#                 self.column_mapping_table.setItem(row, 2, desc_item)
                
#                 # Update description when mapping changes
#                 map_combo.currentTextChanged.connect(
#                     lambda text, row=row: self.update_description(row, text)
#                 )
            
#             self.column_mapping_table.resizeColumnsToContents()
#         except Exception as e:
#             QMessageBox.critical(self, "Error", f"Could not read file: {str(e)}")
    
#     def update_description(self, row, mapped_column):
#         desc_item = QTableWidgetItem(
#             self.EXPECTED_COLUMNS.get(mapped_column, "")
#         )
#         desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
#         self.column_mapping_table.setItem(row, 2, desc_item)
    
#     def process_excel(self):
#         file_path = self.file_path_label.text()
#         if file_path == "No file selected":
#             QMessageBox.warning(self, "Warning", "Please select an Excel file first.")
#             return
        
#         try:
#             # Read the file
#             df = pd.read_excel(file_path)
            
#             # Create column mapping dictionary
#             column_map = {}
#             for row in range(self.column_mapping_table.rowCount()):
#                 original_col = self.column_mapping_table.item(row, 0).text()
#                 mapped_col = self.column_mapping_table.cellWidget(row, 1).currentText()
#                 column_map[original_col] = mapped_col
            
#             # Rename columns
#             df = df.rename(columns=column_map)
            
#             # Perform processing (similar to your original script)
#             df = df.drop(columns=[
#                 col for col in ['BIS NO.', 'BIS MODEL NO.', 'MAH', 'MADE IN', 'LOGO'] 
#                 if col in df.columns
#             ])
#             df.dropna(inplace=True, subset=['MARK'])
            
#             df['Block'] = (df['PCS_CTN'] != df['PCS_CTN'].shift()).cumsum()
            
#             # Group and aggregate
#             grouped = df.groupby(['MARK', 'Block']).agg(
#                 First_CTN=('CTN_NO', 'first'),
#                 Last_CTN=('CTN_NO', 'last'),
#                 T_CTN=('CTN_TOTAL', 'sum'),
#                 WT=('WEIGHT_TOTAL', 'first'),
#                 UNIT=('UNITS', 'first'),
#                 QTY=('PCS_CTN', 'first'),
#                 T_QTY=('PCS_CTN', 'sum'),
#                 DESCRIPTION=('DESCRIPTION', 'first')
#             ).reset_index()
            
#             # Generate CTN NO field
#             grouped['CTN_NO'] = grouped.apply(
#                 lambda row: (
#                     f"{row['First_CTN'] if row['First_CTN'] is not None else 1} - "
#                     f"{row['Last_CTN'] if row['Last_CTN'] is not None else int(row['T_CTN'])}"
#                     if row['T_CTN'] > 1
#                     else f"{row['First_CTN'] if row['First_CTN'] is not None else 1}"
#                 ),
#                 axis=1
#             )
            
#             grouped['T.QTY'] = grouped['T_QTY']
#             grouped['T.CTN'] = grouped['T_CTN']
            
#             # Final columns
#             final_columns = ['MARK', 'CTN_NO', 'DESCRIPTION', 'T.CTN', 'QTY', 'UNIT', 'T.QTY', 'WT']
#             final_dataset = grouped[final_columns]
            
#             # Consolidated dataset
#             consolidated_dataset = final_dataset.groupby(['DESCRIPTION']).agg(
#                 Description=('DESCRIPTION', 'first'),
#                 Quantity=('T.QTY', 'sum'),
#             )
            
#             # Save processed file
#             output_path = file_path.replace('.xlsx', '_processed.xlsx')
#             final_dataset.to_excel(output_path, index=False)
            
#             # Style Excel
#             workbook = openpyxl.load_workbook(output_path)
#             sheet = workbook.active
            
#             cambria_font = Font(name='Cambria', size=11)
#             for row in sheet.iter_rows():
#                 for cell in row:
#                     cell.font = cambria_font
#                     cell.border = None
#                     cell.alignment = Alignment(horizontal="center", vertical="center")
            
#             # Append consolidated data
#             start_row = sheet.max_row + 2
#             start_col = 10  # Column J
            
#             # Write headers
#             for col_idx, header in enumerate(consolidated_dataset.columns, start=start_col):
#                 sheet.cell(row=start_row-1, column=col_idx, value=header)
            
#             # Write data
#             for row_idx, row_data in enumerate(consolidated_dataset.values.tolist(), start=start_row):
#                 for col_idx, value in enumerate(row_data, start=start_col):
#                     sheet.cell(row=row_idx, column=col_idx, value=value)
            
#             # Auto-adjust column widths
#             for column in sheet.columns:
#                 max_length = max(len(str(cell.value)) for cell in column if cell.value) + 2
#                 column[0].column_letter
#                 sheet.column_dimensions[column[0].column_letter].width = max_length
            
#             workbook.save(output_path)
            
#             QMessageBox.information(self, "Success", f"Processed file saved as {output_path}")
        
#         except Exception as e:
#             QMessageBox.critical(self, "Processing Error", str(e))

# def main():
#     app = QApplication(sys.argv)
#     main_window = ExcelProcessorApp()
#     main_window.show()
#     sys.exit(app.exec_())

# if __name__ == '__main__':
#     main()

# import sys
# import pandas as pd
# import numpy as np
# from PyQt5.QtWidgets import (
#     QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, 
#     QPushButton, QFileDialog, QComboBox, QWidget, QTableWidget, 
#     QTableWidgetItem, QMessageBox, QCheckBox
# )
# from PyQt5.QtCore import Qt
# import openpyxl
# from openpyxl.styles import Font, Alignment

# class ExcelProcessorApp(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Distributor Excel Processor")
#         self.setGeometry(100, 100, 1000, 700)

#         # Expected columns with their descriptions
#         self.EXPECTED_COLUMNS = {
#             'MARK': 'Product Mark/Type',
#             'CTN_NO': 'Carton Number',
#             'DESCRIPTION': 'Product Description',
#             'PCS_TOTAL': 'Total Pieces',
#             'UNITS': 'Unit of Measurement', 
#             'PCS_CTN': 'Pieces per Carton',
#             'CTN_TOTAL': 'Total Cartons',
#             'CBM_TOTAL': 'Total Cubic Meters',
#             'WEIGHT_TOTAL': 'Total Weight'
#         }

#         self.initUI()
        
#     def initUI(self):
#         # Main widget and layout
#         main_widget = QWidget()
#         main_layout = QVBoxLayout()
        
#         # File selection section
#         file_layout = QHBoxLayout()
#         self.file_path_label = QLabel("No file selected")
#         select_file_btn = QPushButton("Select Excel File")
#         select_file_btn.clicked.connect(self.select_excel_file)
        
#         file_layout.addWidget(self.file_path_label)
#         file_layout.addWidget(select_file_btn)
        
#         # Column selection section
#         column_selection_label = QLabel("Select Columns to Include:")
#         self.column_selection_layout = QVBoxLayout()
        
#         # Column mapping section
#         self.column_mapping_table = QTableWidget()
#         self.column_mapping_table.setColumnCount(4)  # Added one column for checkbox
#         self.column_mapping_table.setHorizontalHeaderLabels([
#             "Include", "Original Column", "Map To", "Description"
#         ])
        
#         # Process button
#         process_btn = QPushButton("Process Excel")
#         process_btn.clicked.connect(self.process_excel)
        
#         # Select/Deselect All buttons
#         selection_buttons_layout = QHBoxLayout()
#         select_all_btn = QPushButton("Select All")
#         deselect_all_btn = QPushButton("Deselect All")
#         select_all_btn.clicked.connect(self.select_all_columns)
#         deselect_all_btn.clicked.connect(self.deselect_all_columns)
#         selection_buttons_layout.addWidget(select_all_btn)
#         selection_buttons_layout.addWidget(deselect_all_btn)
        
#         # Add widgets to main layout
#         main_layout.addLayout(file_layout)
#         main_layout.addWidget(column_selection_label)
#         main_layout.addLayout(selection_buttons_layout)
#         main_layout.addWidget(self.column_mapping_table)
#         main_layout.addWidget(process_btn)
        
#         main_widget.setLayout(main_layout)
#         self.setCentralWidget(main_widget)
        
#     def select_all_columns(self):
#         for row in range(self.column_mapping_table.rowCount()):
#             checkbox = self.column_mapping_table.cellWidget(row, 0)
#             if checkbox:
#                 checkbox.setChecked(True)
                
#     def deselect_all_columns(self):
#         for row in range(self.column_mapping_table.rowCount()):
#             checkbox = self.column_mapping_table.cellWidget(row, 0)
#             if checkbox:
#                 checkbox.setChecked(False)
        
#     def select_excel_file(self):
#         file_path, _ = QFileDialog.getOpenFileName(
#             self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
#         )
#         if file_path:
#             self.file_path_label.setText(file_path)
#             self.populate_column_mapping(file_path)
        
#     def populate_column_mapping(self, file_path):
#         try:
#             # Read Excel file headers
#             df = pd.read_excel(file_path, nrows=0)
#             original_columns = list(df.columns)
            
#             # Reset table
#             self.column_mapping_table.setRowCount(len(original_columns))
            
#             for row, col in enumerate(original_columns):
#                 # Checkbox for column selection
#                 checkbox = QCheckBox()
#                 checkbox.setChecked(True)  # Default to selected
#                 self.column_mapping_table.setCellWidget(row, 0, checkbox)
                
#                 # Original column name
#                 original_col_item = QTableWidgetItem(col)
#                 original_col_item.setFlags(original_col_item.flags() & ~Qt.ItemIsEditable)
#                 self.column_mapping_table.setItem(row, 1, original_col_item)
                
#                 # Mapping dropdown
#                 map_combo = QComboBox()
#                 map_combo.addItems([''] + list(self.EXPECTED_COLUMNS.keys()))  # Add empty option
                
#                 # Try to auto-match columns
#                 matching_keys = [
#                     key for key, desc in self.EXPECTED_COLUMNS.items() 
#                     if col.lower() in desc.lower() or desc.lower() in col.lower()
#                 ]
#                 if matching_keys:
#                     map_combo.setCurrentText(matching_keys[0])
                
#                 self.column_mapping_table.setCellWidget(row, 2, map_combo)
                
#                 # Description
#                 desc_item = QTableWidgetItem(
#                     self.EXPECTED_COLUMNS.get(map_combo.currentText(), "")
#                 )
#                 desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
#                 self.column_mapping_table.setItem(row, 3, desc_item)
                
#                 # Update description when mapping changes
#                 map_combo.currentTextChanged.connect(
#                     lambda text, row=row: self.update_description(row, text)
#                 )
            
#             self.column_mapping_table.resizeColumnsToContents()
#         except Exception as e:
#             QMessageBox.critical(self, "Error", f"Could not read file: {str(e)}")
    
#     def update_description(self, row, mapped_column):
#         desc_item = QTableWidgetItem(
#             self.EXPECTED_COLUMNS.get(mapped_column, "")
#         )
#         desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
#         self.column_mapping_table.setItem(row, 3, desc_item)
    
#     def process_excel(self):
#         file_path = self.file_path_label.text()
#         if file_path == "No file selected":
#             QMessageBox.warning(self, "Warning", "Please select an Excel file first.")
#             return
        
#         try:
#             # Read the file
#             df = pd.read_excel(file_path)
            
#             # Create column mapping dictionary for selected columns only
#             column_map = {}
#             columns_to_drop = []
            
#             for row in range(self.column_mapping_table.rowCount()):
#                 checkbox = self.column_mapping_table.cellWidget(row, 0)
#                 if not checkbox.isChecked():
#                     original_col = self.column_mapping_table.item(row, 1).text()
#                     columns_to_drop.append(original_col)
#                     continue
                
#                 original_col = self.column_mapping_table.item(row, 1).text()
#                 mapped_col = self.column_mapping_table.cellWidget(row, 2).currentText()
                
#                 if mapped_col:  # Only map if a mapping is selected
#                     column_map[original_col] = mapped_col
            
#             # Drop unselected columns
#             if columns_to_drop:
#                 df = df.drop(columns=columns_to_drop)
            
#             # Rename selected columns
#             df = df.rename(columns=column_map)
            
#             # Perform processing
#             df.dropna(inplace=True, subset=['MARK'])
            
#             df['Block'] = (df['PCS_CTN'] != df['PCS_CTN'].shift()).cumsum()
            
#             # Group and aggregate
#             grouped = df.groupby(['MARK', 'Block']).agg(
#                 First_CTN=('CTN_NO', 'first'),
#                 Last_CTN=('CTN_NO', 'last'),
#                 T_CTN=('CTN_TOTAL', 'sum'),
#                 WT=('WEIGHT_TOTAL', 'first'),
#                 UNIT=('UNITS', 'first'),
#                 QTY=('PCS_CTN', 'first'),
#                 T_QTY=('PCS_CTN', 'sum'),
#                 DESCRIPTION=('DESCRIPTION', 'first')
#             ).reset_index()
            
#             # Generate CTN NO field
#             grouped['CTN_NO'] = grouped.apply(
#                 lambda row: (
#                     f"{row['First_CTN'] if row['First_CTN'] is not None else 1} - "
#                     f"{row['Last_CTN'] if row['Last_CTN'] is not None else int(row['T_CTN'])}"
#                     if row['T_CTN'] > 1
#                     else f"{row['First_CTN'] if row['First_CTN'] is not None else 1}"
#                 ),
#                 axis=1
#             )
            
#             grouped['T.QTY'] = grouped['T_QTY']
#             grouped['T.CTN'] = grouped['T_CTN']
            
#             # Final columns
#             final_columns = ['MARK', 'CTN_NO', 'DESCRIPTION', 'T.CTN', 'QTY', 'UNIT', 'T.QTY', 'WT']
#             final_dataset = grouped[final_columns]
            
#             # Consolidated dataset
#             consolidated_dataset = final_dataset.groupby(['DESCRIPTION']).agg(
#                 Description=('DESCRIPTION', 'first'),
#                 Quantity=('T.QTY', 'sum'),
#             )
            
#             # Save processed file
#             output_path = file_path.replace('.xlsx', '_processed.xlsx')
#             with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
#                 final_dataset.to_excel(writer, index=False, sheet_name='Processed Data')
#                 consolidated_dataset.to_excel(writer, 
#                                            sheet_name='Processed Data', 
#                                            startrow=len(final_dataset) + 3,
#                                            startcol=9)
            
#             # Style Excel
#             workbook = openpyxl.load_workbook(output_path)
#             sheet = workbook.active
            
#             cambria_font = Font(name='Cambria', size=11)
#             for row in sheet.iter_rows():
#                 for cell in row:
#                     cell.font = cambria_font
#                     cell.border = None
#                     cell.alignment = Alignment(horizontal="center", vertical="center")
            
#             # Auto-adjust column widths
#             for column in sheet.columns:
#                 max_length = max(len(str(cell.value)) for cell in column if cell.value) + 2
#                 sheet.column_dimensions[column[0].column_letter].width = max_length
            
#             workbook.save(output_path)
            
#             QMessageBox.information(self, "Success", f"Processed file saved as {output_path}")
        
#         except Exception as e:
#             print(e)
#             QMessageBox.critical(self, "Processing Error", str(e))

# def main():
#     app = QApplication(sys.argv)
#     main_window = ExcelProcessorApp()
#     main_window.show()
#     sys.exit(app.exec_())

# if __name__ == '__main__':
#     main()

import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, 
    QHBoxLayout, QWidget, QFileDialog, QTableWidget, QTableWidgetItem, QComboBox, QMessageBox
)
from PyQt5.QtCore import Qt

class ExcelColumnMapper(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Column Mapper")
        self.setGeometry(200, 200, 800, 600)
        self.init_ui()
        self.file_path = None
        self.df = None

    def init_ui(self):
        layout = QVBoxLayout()

        # File selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        file_button = QPushButton("Select File")
        file_button.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(file_button)

        # Table to display data
        self.table_widget = QTableWidget()

        # Dropdowns for column mapping
        self.mapping_layout = QVBoxLayout()
        self.mapping_widgets = []

        # Process and Save buttons
        button_layout = QHBoxLayout()
        process_button = QPushButton("Process")
        process_button.clicked.connect(self.process_data)
        save_button = QPushButton("Save Processed File")
        save_button.clicked.connect(self.save_file)
        button_layout.addWidget(process_button)
        button_layout.addWidget(save_button)

        # Adding all to main layout
        layout.addLayout(file_layout)
        layout.addWidget(self.table_widget)
        layout.addLayout(self.mapping_layout)
        layout.addLayout(button_layout)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self):
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if self.file_path:
            self.file_label.setText(self.file_path)
            self.load_data()

    def load_data(self):
        try:
            self.df = pd.read_excel(self.file_path)
            self.populate_table()
            self.create_mapping_dropdowns()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file: {e}")

    def populate_table(self):
        self.table_widget.setRowCount(len(self.df))
        self.table_widget.setColumnCount(len(self.df.columns))
        self.table_widget.setHorizontalHeaderLabels(self.df.columns)

        for row in range(len(self.df)):
            for col in range(len(self.df.columns)):
                item = QTableWidgetItem(str(self.df.iloc[row, col]))
                self.table_widget.setItem(row, col, item)

    def create_mapping_dropdowns(self):
        # Clear existing dropdowns
        for widget in self.mapping_widgets:
            widget.deleteLater()
        self.mapping_widgets = []

        self.mapping_layout.addWidget(QLabel("Map Columns:"))
        predefined_keys = ["Key1", "Key2", "Key3"]

        for col in self.df.columns:
            layout = QHBoxLayout()
            label = QLabel(f"{col}:")
            dropdown = QComboBox()
            dropdown.addItems(["Select Key"] + predefined_keys)
            layout.addWidget(label)
            layout.addWidget(dropdown)
            self.mapping_layout.addLayout(layout)
            self.mapping_widgets.append(dropdown)

    def process_data(self):
        if not self.df:
            QMessageBox.warning(self, "Warning", "No data loaded to process.")
            return

        mappings = {}
        for i, col in enumerate(self.df.columns):
            selected_key = self.mapping_widgets[i].currentText()
            if selected_key != "Select Key":
                mappings[selected_key] = col

        if not mappings:
            QMessageBox.warning(self, "Warning", "No columns mapped.")
            return

        try:
            # Example transformation based on mappings
            grouped = self.df.groupby(mappings.values()).sum()  # Example aggregation
            self.processed_df = grouped.reset_index()
            QMessageBox.information(self, "Success", "Data processed successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to process data: {e}")

    def save_file(self):
        if not hasattr(self, 'processed_df'):
            QMessageBox.warning(self, "Warning", "No processed data to save.")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Processed File", "", "Excel Files (*.xlsx *.xls)")
        if save_path:
            try:
                self.processed_df.to_excel(save_path, index=False)
                QMessageBox.information(self, "Success", "File saved successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelColumnMapper()
    window.show()
    sys.exit(app.exec_())
