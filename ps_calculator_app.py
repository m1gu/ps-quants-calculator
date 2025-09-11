import sys
import pandas as pd
import datetime # Import datetime module
import sqlite3 # For database
import json    # For storing analyte data
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QScrollArea, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy,
    QGroupBox,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QSplitter, QListWidget, QListWidgetItem,
    QDateEdit # Added for date input
)
from PyQt5.QtGui import QDoubleValidator, QFont, QColor
from PyQt5.QtCore import Qt, QDate
from openpyxl.utils import get_column_letter # Import for autofit

LOQ = 0.1 # Limit of Quantitation

STATE_LIMITS = {
    "Abamectin": 0.5,
    "Acephate": 0.4,
    "Acequinocyl": 2.0,
    "Acetamiprid": 0.2,
    "Aldicarb": 0.4,
    "Azoxystrobin": 0.2,
    "Bifenazate": 0.2,
    "Bifenthrin": 0.2,
    "Boscalid": 0.4,
    "Carbaryl": 0.2,
    "Carbofuran": 0.2,
    "Chlorantraniliprole": 0.2,
    "Chlorfenapyr": 1.0,
    "Chlorpyrifos": 0.2,
    "Clofentezine": 0.2,
    "Cyfluthrin": 1.0,
    "Cypermethrin": 1.0,
    "Daminozide": 1.0,
    "Diazinon": 0.2,
    "Dichlorvos": 1.0,
    "Dimethoate": 0.2,
    "Ethoprophos": 0.2,
    "Etofenprox": 0.4,
    "Etoxazole": 0.2,
    "Fenoxycarb": 0.2,
    "Fenpyroximate": 0.4,
    "Fipronil": 0.4,
    "Flonicamid": 1.0,
    "Fludioxonil": 0.4,
    "Hexythiazox": 1.0,
    "Imazalil": 0.2,
    "Imidacloprid": 0.4,
    "Kresoxim-methyl": 0.4,
    "Malathion A": 0.2,
    "Metalaxyl": 0.2,
    "Methiocarb": 0.2,
    "Methomyl": 0.4,
    "Methyl parathion": 0.2,
    "MGK 264": 0.2,
    "Myclobutanil": 0.2,
    "Naled": 0.5,
    "Oxamyl": 1.0,
    "Paclobutrazol": 0.4,
    "Permethrins*": 0.2, # Assuming key matches ANALYTES list
    "Phosmet": 0.2,
    "Piperonyl butoxide": 2.0,
    "Prallethrin": 0.2,
    "Propiconazole": 0.4,
    "Propoxure": 0.2,
    "Pyrethrins*": 1.0, # Assuming key matches ANALYTES list
    "Pyridaben": 0.2,
    "Spinosad*": 0.2, # Assuming key matches ANALYTES list
    "Spiromesifen": 0.2,
    "Spirotetramat": 0.2,
    "Spiroxamine": 0.4,
    "Tebuconazole": 0.4,
    "Thiacloprid": 0.2,
    "Thiamethoxam": 0.2,
    "Trifloxystrobin": 0.2
}

ANALYTES = list(STATE_LIMITS.keys()) # Ensure ANALYTES list matches keys

DB_NAME = "saved_samples.db"

# Basic Stylesheet
STYLESHEET = """
QWidget {
    font-size: 10pt;
    background-color: #fcfcfc; /* Light background */
    color: #333333; /* Darker text */
}
QGroupBox {
    font-weight: bold;
    border: 1px solid #e0e0e0; /* Softer border */
    border-radius: 6px;
    margin-top: 6px;
    background-color: #ffffff; /* White background for groupbox */
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 3px 0 3px;
}
QLabel {
    padding-top: 2px;
    background-color: transparent; /* Ensure label background is transparent */
}
QLineEdit, QDateEdit {
    padding: 4px;
    border: 1px solid #cccccc; /* Slightly darker border for inputs */
    border-radius: 4px;
    background-color: #ffffff;
}
QDateEdit::drop-down {
    border: 1px solid #cccccc;
    background-color: #f0f0f0;
}
QDateEdit::down-arrow {
    /* You might need to provide an image for a custom arrow */
    /* image: url(path/to/your/arrow.png); */
    width: 12px;
    height: 12px;
}
QPushButton {
    padding: 6px 15px;
    border: 1px solid #2a9d8f; /* Teal border */
    border-radius: 4px;
    background-color: #2a9d8f; /* Teal background */
    color: white;
    font-weight: bold;
    outline: none; /* Remove focus outline */
}
QPushButton:hover {
    background-color: #268c80; /* Darker teal on hover */
    border-color: #268c80;
}
QPushButton:pressed {
    background-color: #217a70; /* Even darker teal when pressed */
}

/* Style the Delete button specifically */
QPushButton#delete_button {
    background-color: #e76f51; /* Coral red */
    border-color: #e76f51;
}
QPushButton#delete_button:hover {
    background-color: #d66041; /* Darker coral red */
    border-color: #d66041;
}
QPushButton#delete_button:pressed {
    background-color: #c55131; /* Even darker */
}

QListWidget { /* Style ListWidget if you were using it */
    border: 1px solid #cccccc;
    border-radius: 3px;
    background-color: #ffffff;
}
QTableWidget {
    border: 1px solid #e0e0e0; /* Match groupbox border */
    border-radius: 3px;
    gridline-color: #e0e0e0; /* Lighter grid lines */
    background-color: #ffffff;
    alternate-background-color: #f8f8f8; /* Subtle alternating row color */
}
QHeaderView::section {
    background-color: #f0f0f0; /* Light gray header */
    padding: 5px;
    border: none; /* Remove header cell borders */
    border-bottom: 1px solid #e0e0e0; /* Border only at the bottom */
    font-weight: bold;
}
QSplitter::handle {
    background-color: #e0e0e0; /* Make splitter handle visible */
    height: 1px; /* Or width depending on orientation */
}
QSplitter::handle:horizontal {
    width: 1px;
    margin: 0 4px; /* Add some margin */
}
QSplitter::handle:vertical {
    height: 1px;
    margin: 4px 0;
}
"""

class PSCalculatorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.analyte_amount_inputs = {}
        self.db_conn = None
        self.setup_database()
        self.initUI()

    def setup_database(self):
        """Connects to the SQLite DB and creates the table if it doesn't exist."""
        try:
            self.db_conn = sqlite3.connect(DB_NAME)
            cursor = self.db_conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS samples (
                    sample_number TEXT PRIMARY KEY,      -- Will store yyyymmdd_originalSampleNumber
                    original_sample_number TEXT, -- Store the actual sample number entered by user
                    client_name TEXT,
                    sample_date TEXT,          -- Added column for sample date (YYYY-MM-DD)
                    dilution_factor REAL,
                    mass_mg REAL,
                    analyte_data TEXT
                )
            """)
            self.db_conn.commit()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Could not initialize database: {e}")
            # Optionally exit or disable DB features
            self.db_conn = None

    def load_samples_table(self):
        """Loads sample info from DB into the table widget, storing DB key separately."""
        if not self.db_conn:
            return

        # Try to preserve selection
        current_selection_key = None
        selected_row = self.saved_samples_table.currentRow()
        if selected_row >= 0:
            item = self.saved_samples_table.item(selected_row, 0)
            if item:
                current_selection_key = item.data(Qt.UserRole)

        self.saved_samples_table.setRowCount(0) # Clear table content but keep headers
        self.saved_samples_table.setSortingEnabled(False) # Disable sorting during population
        try:
            cursor = self.db_conn.cursor()
            # Select the key and relevant display columns
            cursor.execute("""SELECT sample_number, original_sample_number, sample_date, dilution_factor, mass_mg
                           FROM samples ORDER BY sample_number DESC""")
            samples = cursor.fetchall()
            row_to_reselect = -1

            self.saved_samples_table.setRowCount(len(samples))
            for row, sample_data in enumerate(samples):
                db_key, original_num, sample_date, dilution, mass_mg = sample_data

                item_date = QTableWidgetItem(sample_date or '(NoDate)')
                item_date.setData(Qt.UserRole, db_key) # Store DB key in first column item
                item_num = QTableWidgetItem(original_num or '(NoNum)')
                item_dil = QTableWidgetItem(str(dilution))
                item_mass = QTableWidgetItem(str(mass_mg))

                # Make items read-only
                item_date.setFlags(item_date.flags() & ~Qt.ItemIsEditable)
                item_num.setFlags(item_num.flags() & ~Qt.ItemIsEditable)
                item_dil.setFlags(item_dil.flags() & ~Qt.ItemIsEditable)
                item_mass.setFlags(item_mass.flags() & ~Qt.ItemIsEditable)

                self.saved_samples_table.setItem(row, 0, item_date)
                self.saved_samples_table.setItem(row, 1, item_num)
                self.saved_samples_table.setItem(row, 2, item_dil)
                self.saved_samples_table.setItem(row, 3, item_mass)

                if db_key == current_selection_key:
                    row_to_reselect = row

            # Reselect the previously selected row if it still exists
            if row_to_reselect >= 0:
                 self.saved_samples_table.selectRow(row_to_reselect)

        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Could not load samples table: {e}")
        finally:
             self.saved_samples_table.setSortingEnabled(True) # Re-enable sorting

    def initUI(self):
        self.setWindowTitle('PS Quants Calculator')
        self.setGeometry(100, 100, 1000, 750) # Wider window for splitter
        self.setStyleSheet(STYLESHEET)

        # --- Main Splitter Layout ---
        main_splitter = QSplitter(Qt.Horizontal)
        # Use a QVBoxLayout to hold the splitter
        top_layout = QVBoxLayout(self)
        top_layout.addWidget(main_splitter)
        self.setLayout(top_layout)

        # --- Left Panel (Calculator) ---
        left_widget = QWidget()
        main_layout = QVBoxLayout() # This is the layout for the left panel
        left_widget.setLayout(main_layout)
        main_splitter.addWidget(left_widget)

        # --- Title ---
        title_label = QLabel("Pesticide Quantitation Calculator")
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # --- Input GroupBox ---
        input_groupbox = QGroupBox("Sample Information")
        input_layout = QVBoxLayout()
        input_layout.setSpacing(8)
        input_groupbox.setLayout(input_layout)

        # Sample Number
        sample_layout = QHBoxLayout()
        sample_layout.addWidget(QLabel("Sample Number:"))
        self.sample_input = QLineEdit()
        sample_layout.addWidget(self.sample_input)
        input_layout.addLayout(sample_layout)

        # Client Name - New input field
        client_layout = QHBoxLayout()
        client_layout.addWidget(QLabel("Client Name:"))
        self.client_name_input = QLineEdit()
        client_layout.addWidget(self.client_name_input)
        input_layout.addLayout(client_layout)

        # Sample Date - New Input Field
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Sample Date:"))
        self.sample_date_input = QDateEdit()
        self.sample_date_input.setDate(QDate.currentDate()) # Default to today
        self.sample_date_input.setDisplayFormat("yyyy-MM-dd")
        self.sample_date_input.setCalendarPopup(True)
        date_layout.addWidget(self.sample_date_input)
        input_layout.addLayout(date_layout)

        # Dilution Factor
        dilution_layout = QHBoxLayout()
        dilution_layout.addWidget(QLabel("Dilution Factor:"))
        self.dilution_input = QLineEdit()
        self.dilution_input.setValidator(QDoubleValidator(0.0, 1000000.0, 5))
        self.dilution_input.textChanged.connect(self._update_results_table) # Connect signal
        dilution_layout.addWidget(self.dilution_input)
        input_layout.addLayout(dilution_layout)

        # Mass (mg) - New input field
        mass_mg_layout = QHBoxLayout()
        mass_mg_layout.addWidget(QLabel("Mass (mg):"))
        self.mass_mg_input = QLineEdit() # Input in mg
        self.mass_mg_input.setValidator(QDoubleValidator(0.01, 10000000.0, 5)) # Validator for mg
        self.mass_mg_input.textChanged.connect(self._update_results_table) # Connect signal
        mass_mg_layout.addWidget(self.mass_mg_input)
        input_layout.addLayout(mass_mg_layout)

        main_layout.addWidget(input_groupbox)

        # --- Analytes Table Section ---
        analyte_groupbox = QGroupBox("Analytes")
        analyte_layout = QVBoxLayout()
        analyte_groupbox.setLayout(analyte_layout)

        self.analytes_table = QTableWidget()
        self.analytes_table.setRowCount(len(ANALYTES))
        # Remove LOD column: now 6 columns
        self.analytes_table.setColumnCount(6)
        self.analytes_table.setHorizontalHeaderLabels([
            "Analyte Name", "Amount", "LOQ", "State Limit", "Final Result", "Status"
        ])
        self.analytes_table.setAlternatingRowColors(True)
        # Make specific columns read-only, allow editing Amount
        # self.analytes_table.setEditTriggers(QTableWidget.NoEditTriggers)

        double_validator = QDoubleValidator(0.0, 1000000.0, 5)

        for row, analyte_name in enumerate(ANALYTES):
            # Analyte Name (read-only)
            name_item = QTableWidgetItem(analyte_name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.analytes_table.setItem(row, 0, name_item)

            # Analyte Amount (Editable QLineEdit)
            amount_input = QLineEdit("0") # Default to 0
            amount_input.setValidator(double_validator)
            amount_input.textChanged.connect(self._update_results_table) # Connect signal
            self.analytes_table.setCellWidget(row, 1, amount_input)
            self.analyte_amount_inputs[analyte_name] = amount_input # Store reference

            # LOQ (read-only) - moves to column 2
            loq_item = QTableWidgetItem(str(LOQ))
            loq_item.setFlags(loq_item.flags() & ~Qt.ItemIsEditable)
            loq_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 2, loq_item) # LOQ is column 2

            # State Limit (read-only)
            state_limit = STATE_LIMITS.get(analyte_name, 0.0) # Get limit, default 0 if not found
            limit_item = QTableWidgetItem(str(state_limit))
            limit_item.setFlags(limit_item.flags() & ~Qt.ItemIsEditable)
            limit_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 3, limit_item) # State Limit is column 3

            # Final Result (read-only, initially ND)
            result_item = QTableWidgetItem("ND") # Default to ND
            result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)
            result_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 4, result_item) # Final Result is column 4

            # Status (read-only, initially '-')
            status_item = QTableWidgetItem("-") # Default to Pass since result is ND
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            status_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 5, status_item) # Status is column 5

        # Resize columns
        header = self.analytes_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch) # Analyte Name stretch
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents) # Amount content size
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents) # LOQ content size
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents) # State Limit content size
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents) # Final Result content size
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents) # Status content size
        self.analytes_table.verticalHeader().setVisible(False) # Hide row numbers
        self.analytes_table.verticalHeader().setDefaultSectionSize(22) # Set smaller row height

        analyte_layout.addWidget(self.analytes_table)
        main_layout.addWidget(analyte_groupbox)

        # --- Action Buttons (Export/Copy) ---
        self.export_button = QPushButton("Export Results to Excel")
        self.export_button.clicked.connect(self.export_results)
        self.copy_button = QPushButton("Copy Final Results")
        self.copy_button.clicked.connect(self.copy_final_results)
        self.copy_nd_button = QPushButton("Copy ND")
        self.copy_nd_button.clicked.connect(self.copy_nd_results)
        self.clear_button = QPushButton("Clear Inputs") # New button
        self.clear_button.clicked.connect(self.clear_inputs)   # Connect to new method

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.clear_button)  # Add clear button first
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.copy_button)
        button_layout.addWidget(self.copy_nd_button)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)
        # --- End of Left Panel ---

        # --- Right Panel (Saved Samples) ---
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_widget.setLayout(right_layout)
        main_splitter.addWidget(right_widget)

        right_layout.addWidget(QLabel("Saved Samples"))

        # Replace QListWidget with QTableWidget
        self.saved_samples_table = QTableWidget()
        self.saved_samples_table.setColumnCount(4)
        self.saved_samples_table.setHorizontalHeaderLabels(["Date", "Sample Number", "Dilution", "Mass (mg)"])
        self.saved_samples_table.setSelectionBehavior(QTableWidget.SelectRows) # Select whole rows
        self.saved_samples_table.setSelectionMode(QTableWidget.SingleSelection) # Only one row at a time
        self.saved_samples_table.setEditTriggers(QTableWidget.NoEditTriggers) # Read-only
        self.saved_samples_table.verticalHeader().setVisible(False) # Hide row numbers
        self.saved_samples_table.setAlternatingRowColors(True)
        self.saved_samples_table.itemSelectionChanged.connect(self.load_selected_sample) # Load on selection change
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.saved_samples_table.setSortingEnabled(True) # Allow sorting by column
        self.saved_samples_table.verticalHeader().setDefaultSectionSize(22) # Set smaller row height

        right_layout.addWidget(self.saved_samples_table)

        saved_button_layout = QHBoxLayout()
        self.save_button = QPushButton("Save Current")
        self.delete_button = QPushButton("Delete Selected")
        self.delete_button.setObjectName("delete_button") # Add object name for styling

        self.save_button.clicked.connect(self.save_current_sample)
        self.delete_button.clicked.connect(self.delete_selected_sample)

        saved_button_layout.addWidget(self.save_button)
        saved_button_layout.addWidget(self.delete_button)
        right_layout.addLayout(saved_button_layout)
        # --- End of Right Panel ---

        # Configure Splitter Ratio (optional)
        main_splitter.setSizes([700, 350]) # Adjust ratio if needed

        # Load initial data into table
        self.load_samples_table()
        self._update_results_table() # Initial calculation pass

        self.show()

    # --- Sample Management Methods ---
    def save_current_sample(self):
        if not self.db_conn:
            QMessageBox.warning(self, "Database Error", "Database connection not available.")
            return

        sample_number = self.sample_input.text().strip()
        if not sample_number:
            QMessageBox.warning(self, "Input Error", "Please enter a Sample Number before saving.")
            return

        # Construct the database key using the DATE FROM THE WIDGET
        sample_date_str = self.sample_date_input.date().toString(Qt.ISODate) # YYYY-MM-DD
        db_key = f"{sample_date_str.replace('-', '')}_{sample_number}" # yyyymmdd_sampleNumber

        client_name = self.client_name_input.text().strip()
        dilution_text = self.dilution_input.text().strip()
        mass_mg_text = self.mass_mg_input.text().strip()

        try:
            dilution_factor = float(dilution_text) if dilution_text else 0.0
            mass_mg = float(mass_mg_text) if mass_mg_text else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid number in Dilution Factor or Mass (mg). Cannot save.")
            return

        # Get analyte amounts
        analyte_amounts = {}
        for analyte_name, input_widget in self.analyte_amount_inputs.items():
            amount_text = input_widget.text().strip()
            try:
                analyte_amounts[analyte_name] = float(amount_text) if amount_text else 0.0
            except ValueError:
                # This shouldn't happen with validators, but handle defensively
                QMessageBox.warning(self, "Data Error", f"Invalid amount found for {analyte_name}. Cannot save.")
                return

        # Convert analyte data to JSON
        analyte_data_json = json.dumps(analyte_amounts)

        # Check if entry already exists
        overwrite = False
        try:
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT 1 FROM samples WHERE sample_number = ?", (db_key,))
            exists = cursor.fetchone()
            if exists:
                reply = QMessageBox.question(self, 'Confirm Overwrite',
                    f"A saved entry for sample '{sample_number}' on {sample_date_str} already exists.\nOverwrite it?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.No:
                    return # User chose not to overwrite
                else:
                    overwrite = True # User confirmed overwrite

        except sqlite3.Error as e:
             QMessageBox.critical(self, "Database Error", f"Error checking for existing sample: {e}")
             return

        # Save to DB (INSERT OR REPLACE handles both new and overwrite cases)
        try:
            cursor = self.db_conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO samples
                (sample_number, original_sample_number, client_name, sample_date, dilution_factor, mass_mg, analyte_data)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (db_key, sample_number, client_name, sample_date_str, dilution_factor, mass_mg, analyte_data_json))
            self.db_conn.commit()
            save_message = "updated" if overwrite else "saved"
            QMessageBox.information(self, "Success", f"Sample '{sample_number}' for {sample_date_str} {save_message} successfully.")
            self.load_samples_table() # Refresh the table
        except sqlite3.Error as e:
             QMessageBox.critical(self, "Database Error", f"Could not save sample '{db_key}': {e}")

    def load_selected_sample(self): # Removed arguments as itemSelectionChanged doesn't pass useful ones
        selected_row = self.saved_samples_table.currentRow() # Get selected row index
        if selected_row < 0: # No row selected or selection cleared
            return

        db_key_item = self.saved_samples_table.item(selected_row, 0) # Get item from first column
        if not db_key_item:
             QMessageBox.warning(self, "Error", "Could not retrieve item for selected row.")
             return

        db_key = db_key_item.data(Qt.UserRole) # Get the key from item data
        if not db_key:
            QMessageBox.warning(self, "Error", "Could not retrieve key for selected item.")
            return

        if not self.db_conn:
            QMessageBox.warning(self, "Database Error", "Database connection not available.")
            return

        try:
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT * FROM samples WHERE sample_number = ?", (db_key,)) # Query using the db_key
            sample_data = cursor.fetchone()

            if not sample_data:
                QMessageBox.warning(self, "Error", f"Sample '{db_key}' not found in database.")
                return

            # Unpack data (order depends on CREATE TABLE - note the added original_sample_number and sample_date)
            _db_key, original_sample_number, client_name, sample_date_str, dilution_factor, mass_mg, analyte_data_json = sample_data

            # Populate UI fields (block signals temporarily to avoid multiple updates)
            self.sample_input.blockSignals(True)
            self.client_name_input.blockSignals(True)
            self.sample_date_input.blockSignals(True) # Block date signals
            self.dilution_input.blockSignals(True)
            self.mass_mg_input.blockSignals(True)

            self.sample_input.setText(original_sample_number or "") # Use original sample number here
            self.client_name_input.setText(client_name or "")
            # Set Date
            loaded_date = QDate.fromString(sample_date_str, Qt.ISODate) if sample_date_str else QDate.currentDate()
            self.sample_date_input.setDate(loaded_date)
            self.dilution_input.setText(str(dilution_factor))
            self.mass_mg_input.setText(str(mass_mg))

            self.sample_input.blockSignals(False)
            self.client_name_input.blockSignals(False)
            self.sample_date_input.blockSignals(False) # Unblock date signals
            self.dilution_input.blockSignals(False)
            self.mass_mg_input.blockSignals(False)

            # Load analyte data
            analyte_amounts = json.loads(analyte_data_json)
            for analyte_name, amount_widget in self.analyte_amount_inputs.items():
                amount_widget.blockSignals(True)
                amount_widget.setText(str(analyte_amounts.get(analyte_name, 0.0)))
                amount_widget.blockSignals(False)

            # Trigger a single update after all fields are populated
            self._update_results_table()

        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Could not load sample '{db_key}': {e}")
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Data Error", f"Could not parse analyte data for sample '{db_key}'.")

    def delete_selected_sample(self):
        selected_row = self.saved_samples_table.currentRow() # Get selected row index
        if selected_row < 0:
            QMessageBox.warning(self, "Selection Error", "Please select a sample from the table to delete.")
            return

        db_key_item = self.saved_samples_table.item(selected_row, 0) # Get item from first column
        if not db_key_item:
            QMessageBox.warning(self, "Error", "Could not retrieve item data for deletion.")
            return

        db_key = db_key_item.data(Qt.UserRole) # Get the key from item data
        if not db_key:
            QMessageBox.warning(self, "Error", "Could not retrieve key for selected item to delete.")
            return

        # Use the display text for confirmation (more complex now, construct from row)
        date_text = self.saved_samples_table.item(selected_row, 0).text()
        num_text = self.saved_samples_table.item(selected_row, 1).text()
        display_text = f"{num_text} ({date_text})"

        reply = QMessageBox.question(self, 'Confirm Delete',
            f"Are you sure you want to delete sample \'{display_text}'?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            try:
                cursor = self.db_conn.cursor()
                cursor.execute("DELETE FROM samples WHERE sample_number = ?", (db_key,)) # Delete using the db_key
                self.db_conn.commit()
                QMessageBox.information(self, "Deleted", f"Sample '{display_text}' deleted.")
                self.load_samples_table() # Refresh table
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Database Error", f"Could not delete sample '{db_key}': {e}")

    # --- Calculation and Export Methods ---
    def _update_results_table(self):
        """Recalculates and updates the Final Result column in the table."""
        try:
            dilution_text = self.dilution_input.text().strip()
            mass_mg_text = self.mass_mg_input.text().strip()
            dilution_factor = float(dilution_text) if dilution_text else 0.0
            mass_in_grams = (float(mass_mg_text) / 1000.0) if mass_mg_text else 0.0
        except ValueError:
            return
        mass_is_valid = mass_in_grams > 0
        for row, analyte_name in enumerate(ANALYTES):
            final_result_str = "-"
            try:
                amount_input_widget = self.analyte_amount_inputs[analyte_name]
                amount_text = amount_input_widget.text().strip()
                analyte_amount = float(amount_text) if amount_text else 0.0
                result_numeric = 0.0 # Initialize numeric result

                if analyte_amount == 0.0:
                    final_result_str = "ND"
                else:
                    if mass_is_valid:
                        result_numeric = (analyte_amount * (dilution_factor / mass_in_grams)) / 1000
                        # With LOD/BQL removed: ND if below LOQ, else numeric
                        if result_numeric < LOQ:
                            final_result_str = "ND"
                        else:
                            final_result_str = "{:.3g}".format(result_numeric)
                    else:
                        final_result_str = "Invalid Mass"
            except ValueError:
                final_result_str = "Invalid Amt"
            except Exception as e:
                print(f"Error updating row {row} ({analyte_name}): {e}")
                final_result_str = "Error"
            result_item = QTableWidgetItem(final_result_str)
            result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)
            result_item.setTextAlignment(Qt.AlignCenter)

            # Color item green if ND
            if final_result_str == "ND":
                result_item.setBackground(QColor('lightgreen'))
            else:
                # Ensure background is cleared if result changes from ND to a number
                result_item.setBackground(QColor('white')) # Or match alternating color if needed

            # Final Result is column 4 after removing LOD
            self.analytes_table.setItem(row, 4, result_item)

            # --- Calculate and Set Status --- 
            status_str = "-"
            state_limit = STATE_LIMITS.get(analyte_name, float('inf')) # Default to infinity if no limit

            if final_result_str == "ND": # ND is Pass
                status_str = "Pass"
            elif final_result_str in ("-", "Invalid Mass", "Invalid Amt", "Error"):
                status_str = "-" # Keep as is for errors
            else:
                try:
                    result_numeric_val = float(final_result_str) # Convert formatted string back
                    if result_numeric_val > state_limit:
                        status_str = "Fail"
                    else:
                        status_str = "Pass"
                except ValueError:
                    status_str = "Error" # Should not happen, but catch anyway

            status_item = QTableWidgetItem(status_str)
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            status_item.setTextAlignment(Qt.AlignCenter)
            # Optionally color Pass/Fail
            if status_str == "Fail":
                 status_item.setForeground(QColor('red'))
            elif status_str == "Pass":
                 status_item.setForeground(QColor('darkgreen'))
            # Status is column 5 after removing LOD
            self.analytes_table.setItem(row, 5, status_item)

    def export_results(self):
        """Exports the current table data to an Excel file."""
        sample_number = self.sample_input.text().strip()
        dilution_text = self.dilution_input.text().strip()
        mass_mg_text = self.mass_mg_input.text().strip()
        if not sample_number:
            QMessageBox.warning(self, "Input Error", "Please enter a Sample Number for the export.")
            return
        try:
            dilution_factor = float(dilution_text) if dilution_text else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid Dilution Factor. Cannot export.")
            return
        try:
            mass_in_grams = (float(mass_mg_text) / 1000.0) if mass_mg_text else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid Mass (mg). Cannot export.")
            return
        export_data = []
        has_calculable_data = False
        for row, analyte_name in enumerate(ANALYTES):
            amount_widget = self.analyte_amount_inputs[analyte_name]
            amount_text = amount_widget.text().strip()
            analyte_amount = 0.0
            try:
                analyte_amount = float(amount_text) if amount_text else 0.0
            except ValueError:
                QMessageBox.warning(self, "Data Error", f"Invalid amount '{amount_text}' for {analyte_name} found in table. Skipping row in export.")
                continue
            # Read from table with updated indices (no LOD column)
            loq_item = self.analytes_table.item(row, 2)
            loq_text = loq_item.text() if loq_item else str(LOQ)
            final_result_item = self.analytes_table.item(row, 4)
            final_result_text = final_result_item.text() if final_result_item else "-"
            state_limit_item = self.analytes_table.item(row, 3)
            state_limit_text = state_limit_item.text() if state_limit_item else "N/A"
            status_item = self.analytes_table.item(row, 5)
            status_text = status_item.text() if status_item else "-"

            export_data.append({
                "Analyte Name": analyte_name,
                "Analyte Amount": analyte_amount,
                "LOQ": loq_text,
                "State Limit": state_limit_text,
                "Final Result": final_result_text,
                "Status": status_text
            })
            if final_result_text not in ("-", "Invalid Mass", "Invalid Amt", "Error"):
                has_calculable_data = True
        if not has_calculable_data:
            reply = QMessageBox.question(self, 'Export Warning', "No valid calculated results found. Export anyway?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return
        elif not export_data:
            QMessageBox.warning(self, "Export Error", "No data to export.")
            return
        df_results = pd.DataFrame(export_data)
        # Update columns for export (LOD removed)
        df_results = df_results[["Analyte Name", "Analyte Amount", "LOQ", "State Limit", "Final Result", "Status"]]
        today_date = datetime.date.today().strftime("%Y%m%d")
        safe_sample_number = "".join(c for c in sample_number if c.isalnum() or c in ('_', '-')).rstrip()
        if not safe_sample_number:
            safe_sample_number = "NoSampleNum"
        default_filename = f"{today_date}_{safe_sample_number}_PSQuants.xlsx"
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Save Export File", default_filename, "Excel Files (*.xlsx);;All Files (*)", options=options)
        if fileName:
            try:
                with pd.ExcelWriter(fileName, engine='openpyxl') as writer:
                    sample_info = {
                        "Sample Number:": sample_number,
                        "Client Name:": self.client_name_input.text().strip(),
                        "Sample Date:": self.sample_date_input.date().toString(Qt.ISODate), # Add date to export
                        "Dilution Factor:": dilution_factor,
                        "Mass (g):": mass_in_grams
                    }
                    df_info_keys = pd.DataFrame(list(sample_info.keys()), columns=["Parameter"])
                    df_info_vals = pd.DataFrame(list(sample_info.values()), columns=["Value"])
                    df_info_keys.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0, startcol=0)
                    df_info_vals.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0, startcol=1)
                    start_row_results = len(sample_info) + 1
                    df_results.to_excel(writer, sheet_name='Sheet1', index=False, startrow=start_row_results)
                    worksheet = writer.sheets['Sheet1']
                    for col_idx, column in enumerate(worksheet.columns):
                        max_length = 0
                        column_letter = get_column_letter(col_idx + 1)
                        if col_idx < len(df_results.columns) + (1 if col_idx > 1 else 0): # Adjust index check for separated info
                            if col_idx <= 1: # Sample info cols
                                header_cell_value = df_info_keys.iloc[0,0] if col_idx == 0 else df_info_vals.iloc[0,0] # Simplification, check keys/vals
                            else: # Results Table starts effectively at col A again, shift index
                                results_col_idx = col_idx - (0 if start_row_results > 1 else 0) # Simplified check
                                if results_col_idx < len(df_results.columns):
                                    header_cell_value = df_results.columns[results_col_idx]
                                else: header_cell_value = "" # Avoid index error
                            # Check length only if header_cell_value is not empty
                            header_length = len(str(header_cell_value)) if header_cell_value else 0
                            if header_length > max_length:
                                max_length = header_length
                        for cell in column:
                            try:
                                if cell.value:
                                    cell_length = len(str(cell.value))
                                    if cell_length > max_length:
                                        max_length = cell_length
                            except: pass
                        adjusted_width = max_length + 2
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                QMessageBox.information(self, "Success", f"Results successfully exported to:\n{fileName}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Could not save the file.\nError: {e}")
        else:
            QMessageBox.information(self, "Cancelled", "Export cancelled.")

    def copy_final_results(self):
        """Copies the 'Final Result' column to the clipboard."""
        final_results = []
        for row in range(self.analytes_table.rowCount()):
            item = self.analytes_table.item(row, 4) # Final Result column index after removing LOD
            if item:
                final_results.append(item.text())
            else:
                final_results.append("")
        if final_results:
            clipboard_text = "\n".join(final_results)
            QApplication.clipboard().setText(clipboard_text)
            QMessageBox.information(self, "Copied", "Final results copied to clipboard.")
        else:
            QMessageBox.warning(self, "Nothing to Copy", "The results table is empty.")

    def copy_nd_results(self):
        """Copies 'ND' for every analyte to the clipboard."""
        nd_results = ["ND"] * len(ANALYTES)
        clipboard_text = "\n".join(nd_results)
        QApplication.clipboard().setText(clipboard_text)
        QMessageBox.information(self, "Copied", f"Copied {len(ANALYTES)} 'ND' results to clipboard.")

    def clear_inputs(self):
        """Clears all input fields on the calculator panel."""
        # Clear main inputs
        self.sample_input.clear()
        self.client_name_input.clear()
        self.sample_date_input.setDate(QDate.currentDate()) # Reset date
        self.dilution_input.clear()
        self.mass_mg_input.clear()

        # Clear analyte amounts (setting to '0' triggers updates)
        for amount_widget in self.analyte_amount_inputs.values():
            # Block signals temporarily if setting '0' causes excessive updates,
            # but it should be fine as _update_results_table handles it.
            amount_widget.setText("0")

        # Optional: Scroll analyte table back to top if needed
        # self.analytes_table.scrollToTop()

        QMessageBox.information(self, "Cleared", "Input fields have been cleared.")

    # --- Cleanup ---
    def closeEvent(self, event):
        """Ensures DB connection is closed when the window closes."""
        if self.db_conn:
            self.db_conn.close()
            print("Database connection closed.")
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PSCalculatorApp()
    sys.exit(app.exec_()) 
