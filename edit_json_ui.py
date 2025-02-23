import sys
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                            QScrollArea, QMessageBox, QLineEdit)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QPalette, QColor, QIcon

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JSON Mapping Editor")
        self.setMinimumSize(800, 600)
        self.json_file_path = None
        self.json_data = None
        self.mapping_inputs = {}
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # File selection area
        file_area = QHBoxLayout()
        self.file_label = QLabel("No JSON file selected")
        select_file_btn = QPushButton("Select JSON File")
        select_file_btn.clicked.connect(self.select_file)
        file_area.addWidget(self.file_label)
        file_area.addWidget(select_file_btn)
        main_layout.addLayout(file_area)
        
        # Add New Mapping Area
        add_new_area = QHBoxLayout()
        self.new_rough = QLineEdit()
        self.new_rough.setPlaceholderText("Enter rough description")
        self.new_formatted = QLineEdit()
        self.new_formatted.setPlaceholderText("Enter formatted description")
        add_btn = QPushButton("Add New Mapping")
        add_btn.clicked.connect(self.add_new_mapping)
        add_new_area.addWidget(self.new_rough)
        add_new_area.addWidget(self.new_formatted)
        add_new_area.addWidget(add_btn)
        main_layout.addLayout(add_new_area)
        
        # Create scroll area for mappings
        scroll = QScrollArea()
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        scroll.setWidget(self.scroll_widget)
        scroll.setWidgetResizable(True)
        main_layout.addWidget(scroll)
        
        # Save button
        self.save_btn = QPushButton("Save Changes")
        self.save_btn.clicked.connect(self.save_changes)
        self.save_btn.setEnabled(False)
        main_layout.addWidget(self.save_btn)
        
        # Status label
        self.status_label = QLabel("")
        main_layout.addWidget(self.status_label)
        
        # Apply styling
        self.apply_styling()
    
    def apply_styling(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
                color: black;
            }
            QWidget {
                background-color: #f0f0f0;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
            QLabel {
                color: #333333;
                font-size: 14px;
            }
            QLineEdit {
                padding: 6px;
                border: 1px solid #BDBDBD;
                border-radius: 4px;
                background-color: white;
                color: black;
            }
            QLineEdit:focus {
                border: 2px solid #2196F3;
            }
        """)
    
    def select_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select JSON Mapping File",
            "",
            "JSON Files (*.json)"
        )
        if file_name:
            try:
                self.json_file_path = file_name
                with open(file_name, 'r') as f:
                    self.json_data = json.load(f)
                self.file_label.setText(f"Selected: {file_name}")
                self.load_mappings()
                self.save_btn.setEnabled(True)
                self.status_label.setText("JSON file loaded successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error loading JSON file: {str(e)}")
    
    def load_mappings(self):
        # Clear existing mappings
        self.mapping_inputs.clear()
        for i in reversed(range(self.scroll_layout.count())):
            widget = self.scroll_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()
        
        # Add mappings from JSON
        for item in self.json_data:
            self.add_mapping_row(item['rough'], item['formatted'])
    
    def add_mapping_row(self, rough, formatted):
        # Create row widget and layout
        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(0, 0, 0, 0)
        
        # Add rough description label
        rough_label = QLabel(rough)
        rough_label.setMinimumWidth(300)
        row_layout.addWidget(rough_label)
        
        # Add formatted description input
        formatted_input = QLineEdit(formatted)
        formatted_input.setMinimumWidth(300)
        row_layout.addWidget(formatted_input)
        
        # Add delete button
        delete_btn = QPushButton("Delete")
        delete_btn.setStyleSheet("background-color: #f44336;")
        delete_btn.clicked.connect(lambda: self.delete_mapping(row_widget, rough))
        row_layout.addWidget(delete_btn)
        
        # Store reference to input
        self.mapping_inputs[rough] = formatted_input
        
        # Add to scroll layout
        self.scroll_layout.addWidget(row_widget)
    
    def add_new_mapping(self):
        rough = self.new_rough.text().strip()
        formatted = self.new_formatted.text().strip()
        
        if not rough or not formatted:
            QMessageBox.warning(self, "Warning", "Both fields are required!")
            return
        
        if rough in self.mapping_inputs:
            QMessageBox.warning(self, "Warning", "This rough description already exists!")
            return
        
        self.add_mapping_row(rough, formatted)
        self.new_rough.clear()
        self.new_formatted.clear()
        self.status_label.setText("New mapping added!")
    
    def delete_mapping(self, row_widget, rough):
        row_widget.deleteLater()
        del self.mapping_inputs[rough]
        self.status_label.setText(f"Mapping for '{rough}' deleted!")
    
    def save_changes(self):
        if not self.json_file_path:
            return
        
        try:
            # Create new JSON data
            new_data = [
                {
                    "rough": rough,
                    "formatted": input_field.text().strip()
                }
                for rough, input_field in self.mapping_inputs.items()
            ]
            
            # Save to file
            with open(self.json_file_path, 'w') as f:
                json.dump(new_data, f, indent=2)
            
            self.status_label.setText("Changes saved successfully!")
            
            # Reload the file to ensure everything is in sync
            with open(self.json_file_path, 'r') as f:
                self.json_data = json.load(f)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error saving changes: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()