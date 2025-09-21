import sys
import os
import pandas as pd
import datetime
import sqlite3
import json
import math
import re


from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QMessageBox,
    QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QSplitter, QDateEdit, QMenu, QAction
)
from PyQt5.QtGui import QDoubleValidator, QFont, QColor, QKeySequence
from PyQt5.QtCore import Qt, QDate, QEvent
from openpyxl.utils import get_column_letter

# =========================
# CONFIGURACIÓN
# =========================
# Por defecto: procesa TODOS los samples con YES (puedes pasar un número para limitar)
DEFAULT_BATCH_LIMIT = None
RAW_SHEET_NAME = "raw results"        # hoja de entrada del Excel

# =========================
# LÓGICA EXISTENTE
# =========================
LOQ = 0.1  # Limit of Quantitation

STATE_LIMITS = {
    "Abamectin": 0.5, "Acephate": 0.4, "Acequinocyl": 2.0, "Acetamiprid": 0.2, "Aldicarb": 0.4,
    "Azoxystrobin": 0.2, "Bifenazate": 0.2, "Bifenthrin": 0.2, "Boscalid": 0.4, "Carbaryl": 0.2,
    "Carbofuran": 0.2, "Chlorantraniliprole": 0.2, "Chlorfenapyr": 1.0, "Chlorpyrifos": 0.2,
    "Clofentezine": 0.2, "Cyfluthrin": 1.0, "Cypermethrin": 1.0, "Daminozide": 1.0, "Diazinon": 0.2,
    "Dichlorvos": 1.0, "Dimethoate": 0.2, "Ethoprophos": 0.2, "Etofenprox": 0.4, "Etoxazole": 0.2,
    "Fenoxycarb": 0.2, "Fenpyroximate": 0.4, "Fipronil": 0.4, "Flonicamid": 1.0, "Fludioxonil": 0.4,
    "Hexythiazox": 1.0, "Imazalil": 0.2, "Imidacloprid": 0.4, "Kresoxim-methyl": 0.4, "Malathion A": 0.2,
    "Metalaxyl": 0.2, "Methiocarb": 0.2, "Methomyl": 0.4, "Methyl parathion": 0.2, "MGK 264": 0.2,
    "Myclobutanil": 0.2, "Naled": 0.5, "Oxamyl": 1.0, "Paclobutrazol": 0.4, "Permethrins*": 0.2,
    "Phosmet": 0.2, "Piperonyl butoxide": 2.0, "Prallethrin": 0.2, "Propiconazole": 0.4, "Propoxure": 0.2,
    "Pyrethrins*": 1.0, "Pyridaben": 0.2, "Spinosad*": 0.2, "Spiromesifen": 0.2, "Spirotetramat": 0.2,
    "Spiroxamine": 0.4, "Tebuconazole": 0.4, "Thiacloprid": 0.2, "Thiamethoxam": 0.2, "Trifloxystrobin": 0.2
}
ANALYTES = list(STATE_LIMITS.keys())
DB_NAME = "saved_samples.db"

STYLESHEET = """
QWidget { font-size: 10pt; background-color: #fcfcfc; color: #333333; }
QGroupBox { font-weight: bold; border: 1px solid #e0e0e0; border-radius: 6px; margin-top: 6px; background-color: #ffffff; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px 0 3px; }
QLabel { padding-top: 2px; background-color: transparent; }
QLineEdit, QDateEdit { padding: 4px; border: 1px solid #cccccc; border-radius: 4px; background-color: #ffffff; }
QPushButton { padding: 6px 15px; border: 1px solid #2a9d8f; border-radius: 4px; background-color: #2a9d8f; color: white; font-weight: bold; outline: none; }
QPushButton:hover { background-color: #268c80; border-color: #268c80; }
QPushButton:pressed { background-color: #217a70; }
QPushButton#delete_button { background-color: #e76f51; border-color: #e76f51; }
QPushButton#delete_button:hover { background-color: #d66041; border-color: #d66041; }
QPushButton#delete_button:pressed { background-color: #c55131; }
QTableWidget { border: 1px solid #e0e0e0; border-radius: 3px; gridline-color: #e0e0e0; background-color: #ffffff; alternate-background-color: #f8f8f8; }
QHeaderView::section { background-color: #f0f0f0; padding: 5px; border: none; border-bottom: 1px solid #e0e0e0; font-weight: bold; }
QSplitter::handle { background-color: #e0e0e0; height: 1px; }
QSplitter::handle:horizontal { width: 1px; margin: 0 4px; }
QSplitter::handle:vertical { height: 1px; margin: 4px 0; }
"""

class PSCalculatorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.analyte_amount_inputs = {}
        self.db_conn = None
        self.setup_database()
        self.initUI()

    # ---------------- DB ----------------
    def setup_database(self):
        try:
            self.db_conn = sqlite3.connect(DB_NAME)
            cursor = self.db_conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS samples (
                    sample_number TEXT PRIMARY KEY,      -- yyyymmdd_originalSampleNumber
                    original_sample_number TEXT,
                    client_name TEXT,
                    sample_date TEXT,                    -- YYYY-MM-DD
                    dilution_factor REAL,
                    mass_mg REAL,
                    analyte_data TEXT
                )
            """)
            self.db_conn.commit()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Could not initialize database: {e}")
            self.db_conn = None

    def _show_saved_table_context_menu(self, pos):
        """
        Menú contextual para la tabla de 'Saved Samples'.
        Opciones: Cargar muestra y Eliminar registro de la base de datos.
        """
        index = self.saved_samples_table.indexAt(pos)
        if not index.isValid():
            return

        row = index.row()
        self.saved_samples_table.selectRow(row)

        menu = QMenu(self.saved_samples_table)

        action_load = QAction("Cargar muestra", self)
        action_delete = QAction("Eliminar registro de la base de datos", self)

        action_load.triggered.connect(self.load_selected_sample)
        action_delete.triggered.connect(self.delete_selected_sample)

        menu.addAction(action_load)
        menu.addSeparator()
        menu.addAction(action_delete)

        menu.exec_(self.saved_samples_table.viewport().mapToGlobal(pos))

    def load_samples_table(self):
        if not self.db_conn:
            return
        current_selection_key = None
        selected_row = self.saved_samples_table.currentRow()
        if selected_row >= 0:
            item = self.saved_samples_table.item(selected_row, 0)
            if item:
                current_selection_key = item.data(Qt.UserRole)

        self.saved_samples_table.setRowCount(0)
        self.saved_samples_table.setSortingEnabled(False)
        try:
            cursor = self.db_conn.cursor()
            cursor.execute("""SELECT sample_number, original_sample_number, sample_date, dilution_factor, mass_mg
                              FROM samples ORDER BY sample_number DESC""")
            samples = cursor.fetchall()
            row_to_reselect = -1
            self.saved_samples_table.setRowCount(len(samples))
            for row, (db_key, original_num, sample_date, dilution, mass_mg) in enumerate(samples):
                item_date = QTableWidgetItem(sample_date or '(NoDate)')
                item_date.setData(Qt.UserRole, db_key)
                item_num = QTableWidgetItem(original_num or '(NoNum)')
                try:
                    item_dil = QTableWidgetItem(f"{float(dilution):g}")
                except Exception:
                    item_dil = QTableWidgetItem(str(dilution))
                item_mass = QTableWidgetItem(str(mass_mg))
                for it in (item_date, item_num, item_dil, item_mass):
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                self.saved_samples_table.setItem(row, 0, item_date)
                self.saved_samples_table.setItem(row, 1, item_num)
                self.saved_samples_table.setItem(row, 2, item_dil)
                self.saved_samples_table.setItem(row, 3, item_mass)
                if db_key == current_selection_key:
                    row_to_reselect = row
            if row_to_reselect >= 0:
                self.saved_samples_table.selectRow(row_to_reselect)
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Could not load samples table: {e}")
        finally:
            self.saved_samples_table.setSortingEnabled(True)

    def _format_sigfigs_no_sci(self, x: float, sig: int = 3) -> str:
        """
        Formatea 'x' con 'sig' cifras significativas SIN notación científica.
        - Redondea correctamente (usa round con decimales calculados por orden de magnitud).
        - No agrega ceros extra al final (quita ceros/punto sobrantes).
        """
        if x == 0 or not math.isfinite(x):
            return "0"
        power = math.floor(math.log10(abs(x)))
        decimals = sig - 1 - power
        # Redondeo a 'decimals' (si decimals < 0 redondea a decenas, centenas, etc.)
        rounded = round(x, decimals)
        if decimals > 0:
            s = f"{rounded:.{decimals}f}"
            s = s.rstrip("0").rstrip(".")  # no agregar ceros innecesarios
            return s if s else "0"
        else:
            # Sin decimales
            return f"{rounded:.0f}"


    # --------------- UI -----------------
    def initUI(self):
        self.setWindowTitle('PS Quants Calculator')
        self.setGeometry(100, 100, 1100, 780)
        self.setStyleSheet(STYLESHEET)

        main_splitter = QSplitter(Qt.Horizontal)
        top_layout = QVBoxLayout(self)
        top_layout.addWidget(main_splitter)
        self.setLayout(top_layout)

        # Left
        left_widget = QWidget()
        main_layout = QVBoxLayout()
        left_widget.setLayout(main_layout)
        main_splitter.addWidget(left_widget)

        title_label = QLabel("Pesticide Quantitation Calculator")
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        input_groupbox = QGroupBox("Sample Information")
        input_layout = QVBoxLayout(); input_groupbox.setLayout(input_layout)

        # Sample Number
        sample_layout = QHBoxLayout()
        sample_layout.addWidget(QLabel("Sample Number:"))
        self.sample_input = QLineEdit()
        sample_layout.addWidget(self.sample_input)
        input_layout.addLayout(sample_layout)

        # Client Name
        client_layout = QHBoxLayout()
        client_layout.addWidget(QLabel("Client Name:"))
        self.client_name_input = QLineEdit()
        client_layout.addWidget(self.client_name_input)
        input_layout.addLayout(client_layout)

        # Sample Date
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Sample Date:"))
        self.sample_date_input = QDateEdit()
        self.sample_date_input.setDate(QDate.currentDate())
        self.sample_date_input.setDisplayFormat("yyyy-MM-dd")
        self.sample_date_input.setCalendarPopup(True)
        date_layout.addWidget(self.sample_date_input)
        input_layout.addLayout(date_layout)

        # Dilution Factor
        dilution_layout = QHBoxLayout()
        dilution_layout.addWidget(QLabel("Dilution Factor:"))
        self.dilution_input = QLineEdit()
        self.dilution_input.setValidator(QDoubleValidator(0.0, 1000000.0, 5))
        self.dilution_input.textChanged.connect(self._update_results_table)
        dilution_layout.addWidget(self.dilution_input)
        input_layout.addLayout(dilution_layout)

        # Mass (mg)
        mass_mg_layout = QHBoxLayout()
        mass_mg_layout.addWidget(QLabel("Mass (mg):"))
        self.mass_mg_input = QLineEdit()
        self.mass_mg_input.setValidator(QDoubleValidator(0.01, 10000000.0, 5))
        self.mass_mg_input.textChanged.connect(self._update_results_table)
        mass_mg_layout.addWidget(self.mass_mg_input)
        input_layout.addLayout(mass_mg_layout)

        main_layout.addWidget(input_groupbox)

        # Analytes table
        analyte_groupbox = QGroupBox("Analytes")
        analyte_layout = QVBoxLayout(); analyte_groupbox.setLayout(analyte_layout)
        self.analytes_table = QTableWidget()
        self.analytes_table.setRowCount(len(ANALYTES))
        self.analytes_table.setColumnCount(6)
        self.analytes_table.setHorizontalHeaderLabels(["Analyte Name", "Amount", "LOQ", "State Limit", "Final Result", "Status"])
        self.analytes_table.setAlternatingRowColors(True)

        double_validator = QDoubleValidator(0.0, 1000000.0, 5)
        for row, analyte_name in enumerate(ANALYTES):
            name_item = QTableWidgetItem(analyte_name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.analytes_table.setItem(row, 0, name_item)

            amount_input = QLineEdit("0")
            amount_input.setValidator(double_validator)
            amount_input.textChanged.connect(self._update_results_table)
            amount_input.installEventFilter(self)  # multi-cell paste
            self.analytes_table.setCellWidget(row, 1, amount_input)
            self.analyte_amount_inputs[analyte_name] = amount_input

            loq_item = QTableWidgetItem(str(LOQ))
            loq_item.setFlags(loq_item.flags() & ~Qt.ItemIsEditable)
            loq_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 2, loq_item)

            state_limit = STATE_LIMITS.get(analyte_name, 0.0)
            limit_item = QTableWidgetItem(str(state_limit))
            limit_item.setFlags(limit_item.flags() & ~Qt.ItemIsEditable)
            limit_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 3, limit_item)

            result_item = QTableWidgetItem("ND")
            result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)
            result_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 4, result_item)

            status_item = QTableWidgetItem("-")
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            status_item.setTextAlignment(Qt.AlignCenter)
            self.analytes_table.setItem(row, 5, status_item)

        header = self.analytes_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.analytes_table.verticalHeader().setVisible(False)
        self.analytes_table.verticalHeader().setDefaultSectionSize(22)

        analyte_layout.addWidget(self.analytes_table)
        main_layout.addWidget(analyte_groupbox)

        # Botones
        self.export_button = QPushButton("Export Results to Excel")
        self.export_button.clicked.connect(self.export_results)
        self.copy_button = QPushButton("Copy Final Results")
        self.copy_button.clicked.connect(self.copy_final_results)
        self.copy_nd_button = QPushButton("Copy ND")
        self.copy_nd_button.clicked.connect(self.copy_nd_results)
        self.clear_button = QPushButton("Clear Inputs")
        self.clear_button.clicked.connect(self.clear_inputs)

        # Batch desde Excel (sin elegir carpeta de salida)
        self.batch_button = QPushButton("Batch desde Excel…")
        self.batch_button.clicked.connect(self._ui_batch_from_excel_dialog)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.copy_button)
        button_layout.addWidget(self.copy_nd_button)
        button_layout.addWidget(self.batch_button)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        # Right (Saved Samples)
        right_widget = QWidget()
        right_layout = QVBoxLayout(); right_widget.setLayout(right_layout)
        main_splitter.addWidget(right_widget)

        right_layout.addWidget(QLabel("Saved Samples"))
        self.saved_samples_table = QTableWidget()
        self.saved_samples_table.setColumnCount(4)
        self.saved_samples_table.setHorizontalHeaderLabels(["Date", "Sample Number", "Dilution", "Mass (mg)"])
        self.saved_samples_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.saved_samples_table.setSelectionMode(QTableWidget.SingleSelection)
        self.saved_samples_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.saved_samples_table.verticalHeader().setVisible(False)
        self.saved_samples_table.setAlternatingRowColors(True)
        self.saved_samples_table.itemSelectionChanged.connect(self.load_selected_sample)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.saved_samples_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.saved_samples_table.setSortingEnabled(True)
        self.saved_samples_table.verticalHeader().setDefaultSectionSize(22)

        # Menú contextual (click derecho)
        self.saved_samples_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.saved_samples_table.customContextMenuRequested.connect(self._show_saved_table_context_menu)

        right_layout.addWidget(self.saved_samples_table)

        saved_button_layout = QHBoxLayout()
        self.save_button = QPushButton("Save Current")
        self.delete_button = QPushButton("Delete Selected")
        self.delete_button.setObjectName("delete_button")
        self.save_button.clicked.connect(self.save_current_sample)
        self.delete_button.clicked.connect(self.delete_selected_sample)
        saved_button_layout.addWidget(self.save_button)
        saved_button_layout.addWidget(self.delete_button)
        right_layout.addLayout(saved_button_layout)

        main_splitter.setSizes([720, 380])

        self.load_samples_table()
        self._update_results_table()
        self.show()

    # =========================
    # CRUD Samples
    # =========================
    def save_current_sample(self):
        if not self.db_conn:
            QMessageBox.warning(self, "Database Error", "Database connection not available.")
            return

        sample_number = self.sample_input.text().strip()
        if not sample_number:
            QMessageBox.warning(self, "Input Error", "Please enter a Sample Number before saving.")
            return

        sample_date_str = self.sample_date_input.date().toString(Qt.ISODate)  # YYYY-MM-DD
        db_key = f"{sample_date_str.replace('-', '')}_{sample_number}"  # yyyymmdd_sampleNumber

        client_name = self.client_name_input.text().strip()
        dilution_text = self.dilution_input.text().strip()
        mass_mg_text = self.mass_mg_input.text().strip()

        try:
            dilution_factor = float(dilution_text) if dilution_text else 0.0
            mass_mg = float(mass_mg_text) if mass_mg_text else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid number in Dilution Factor or Mass (mg). Cannot save.")
            return

        analyte_amounts = {}
        for analyte_name, input_widget in self.analyte_amount_inputs.items():
            amount_text = input_widget.text().strip()
            try:
                analyte_amounts[analyte_name] = float(amount_text) if amount_text else 0.0
            except ValueError:
                QMessageBox.warning(self, "Data Error", f"Invalid amount found for {analyte_name}. Cannot save.")
                return

        analyte_data_json = json.dumps(analyte_amounts)

        overwrite = False
        try:
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT 1 FROM samples WHERE sample_number = ?", (db_key,))
            exists = cursor.fetchone()
            if exists:
                reply = QMessageBox.question(
                    self, 'Confirm Overwrite',
                    f"A saved entry for sample '{sample_number}' on {sample_date_str} already exists.\nOverwrite it?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return
                else:
                    overwrite = True
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Error checking for existing sample: {e}")
            return

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
            self.load_samples_table()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Could not save sample '{db_key}': {e}")

    def load_selected_sample(self):
        selected_row = self.saved_samples_table.currentRow()
        if selected_row < 0:
            return

        db_key_item = self.saved_samples_table.item(selected_row, 0)
        if not db_key_item:
            QMessageBox.warning(self, "Error", "Could not retrieve item for selected row.")
            return

        db_key = db_key_item.data(Qt.UserRole)
        if not db_key:
            QMessageBox.warning(self, "Error", "Could not retrieve key for selected item.")
            return

        if not self.db_conn:
            QMessageBox.warning(self, "Database Error", "Database connection not available.")
            return

        try:
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT * FROM samples WHERE sample_number = ?", (db_key,))
            sample_data = cursor.fetchone()

            if not sample_data:
                QMessageBox.warning(self, "Error", f"Sample '{db_key}' not found in database.")
                return

            _db_key, original_sample_number, client_name, sample_date_str, dilution_factor, mass_mg, analyte_data_json = sample_data

            self.sample_input.blockSignals(True)
            self.client_name_input.blockSignals(True)
            self.sample_date_input.blockSignals(True)
            self.dilution_input.blockSignals(True)
            self.mass_mg_input.blockSignals(True)

            self.sample_input.setText(original_sample_number or "")
            self.client_name_input.setText(client_name or "")
            loaded_date = QDate.fromString(sample_date_str, Qt.ISODate) if sample_date_str else QDate.currentDate()
            self.sample_date_input.setDate(loaded_date)
            try:
                self.dilution_input.setText(f"{float(dilution_factor):g}")
            except Exception:
                self.dilution_input.setText(str(dilution_factor))
            self.mass_mg_input.setText(str(mass_mg))

            self.sample_input.blockSignals(False)
            self.client_name_input.blockSignals(False)
            self.sample_date_input.blockSignals(False)
            self.dilution_input.blockSignals(False)
            self.mass_mg_input.blockSignals(False)

            analyte_amounts = json.loads(analyte_data_json)
            for analyte_name, amount_widget in self.analyte_amount_inputs.items():
                amount_widget.blockSignals(True)
                amount_widget.setText(str(analyte_amounts.get(analyte_name, 0.0)))
                amount_widget.blockSignals(False)

            self._update_results_table()

        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Could not load sample '{db_key}': {e}")
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Data Error", f"Could not parse analyte data for sample '{db_key}'.")

    def delete_selected_sample(self):
        selected_row = self.saved_samples_table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "Selection Error", "Please select a sample from the table to delete.")
            return

        db_key_item = self.saved_samples_table.item(selected_row, 0)
        if not db_key_item:
            QMessageBox.warning(self, "Error", "Could not retrieve item data for deletion.")
            return

        db_key = db_key_item.data(Qt.UserRole)
        if not db_key:
            QMessageBox.warning(self, "Error", "Could not retrieve key for selected item to delete.")
            return

        date_text = self.saved_samples_table.item(selected_row, 0).text()
        num_text = self.saved_samples_table.item(selected_row, 1).text()
        display_text = f"{num_text} ({date_text})"

        reply = QMessageBox.question(
            self, 'Confirm Delete',
            f"Are you sure you want to delete sample '{display_text}'?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
                cursor = self.db_conn.cursor()
                cursor.execute("DELETE FROM samples WHERE sample_number = ?", (db_key,))
                self.db_conn.commit()
                QMessageBox.information(self, "Deleted", f"Sample '{display_text}' deleted.")
                self.load_samples_table()
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Database Error", f"Could not delete sample '{db_key}': {e}")

    # ============================
    # Cálculo & Export
    # ============================
    def _update_results_table(self):
        """
        Recalcula y actualiza la columna 'Final Result' de la tabla.

        NUEVO: Final result = (Amount / Mass_mg) * DF
        - Mass (mg) se usa tal cual viene en la UI (y del Excel).
        - Se mantiene ND si resultado < LOQ (0.1).
        - Formato: 3 cifras significativas ({:.3g}).
        """
        try:
            dilution_text = self.dilution_input.text().strip()
            mass_mg_text = self.mass_mg_input.text().strip()
            dilution_factor = float(dilution_text) if dilution_text else 0.0
            mass_mg = float(mass_mg_text) if mass_mg_text else 0.0
        except ValueError:
            return

        mass_is_valid = mass_mg > 0

        for row, analyte_name in enumerate(ANALYTES):
            final_result_str = "-"
            try:
                amount_input_widget = self.analyte_amount_inputs[analyte_name]
                amount_text = amount_input_widget.text().strip()
                analyte_amount = float(amount_text) if amount_text else 0.0
                result_numeric = 0.0

                if analyte_amount == 0.0:
                    final_result_str = "ND"
                else:
                    if mass_is_valid:
                        # NUEVA FÓRMULA
                        result_numeric = (analyte_amount / mass_mg) * dilution_factor
                        # Lógica ND/LOQ se mantiene
                        if result_numeric < LOQ:
                            final_result_str = "ND"
                        else:
                            #final_result_str = "{:.3g}".format(result_numeric)
                            final_result_str = self._format_sigfigs_no_sci(result_numeric, sig=3)
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

            if final_result_str == "ND":
                result_item.setBackground(QColor('lightgreen'))
            else:
                result_item.setBackground(QColor('white'))

            self.analytes_table.setItem(row, 4, result_item)

            # Status
            status_str = "-"
            state_limit = STATE_LIMITS.get(analyte_name, float('inf'))

            if final_result_str == "ND":
                status_str = "Pass"
            elif final_result_str in ("-", "Invalid Mass", "Invalid Amt", "Error"):
                status_str = "-"
            else:
                try:
                    result_numeric_val = float(final_result_str)
                    status_str = "Fail" if result_numeric_val > state_limit else "Pass"
                except ValueError:
                    status_str = "Error"

            status_item = QTableWidgetItem(status_str)
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            status_item.setTextAlignment(Qt.AlignCenter)
            if status_str == "Fail":
                status_item.setForeground(QColor('red'))
            elif status_str == "Pass":
                status_item.setForeground(QColor('darkgreen'))
            self.analytes_table.setItem(row, 5, status_item)

    def eventFilter(self, obj, event):
        """Enable multi-cell paste into the Amount column."""
        try:
            if event.type() == QEvent.KeyPress and isinstance(obj, QLineEdit):
                if event.matches(QKeySequence.Paste):
                    if obj in self.analyte_amount_inputs.values():
                        text = QApplication.clipboard().text()
                        if ('\n' in text) or ('\t' in text) or ('\r' in text):
                            self._paste_values_into_amounts(obj, text)
                            return True
        except Exception as e:
            print(f"Paste event handling error: {e}")
        return super().eventFilter(obj, event)

    def _paste_values_into_amounts(self, start_widget, text):
        start_row = 0
        for r in range(self.analytes_table.rowCount()):
            if self.analytes_table.cellWidget(r, 1) is start_widget:
                start_row = r
                break

        lines = [line for line in text.splitlines() if line is not None]
        if not lines:
            return

        values = []
        for line in lines:
            if '\t' in line:
                cell = line.split('\t', 1)[0]
            else:
                cell = line
            cell = cell.strip()
            values.append('0' if cell == '' else cell)

        last_row = self.analytes_table.rowCount()
        for i, val in enumerate(values):
            row = start_row + i
            if row >= last_row:
                break
            w = self.analytes_table.cellWidget(row, 1)
            if isinstance(w, QLineEdit):
                w.blockSignals(True)
                w.setText(val)
                w.blockSignals(False)
        self._update_results_table()

    def export_results(self):
        """Export con diálogo (manual)."""
        sample_number = self.sample_input.text().strip()
        dilution_text = self.dilution_input.text().strip()
        mass_mg_text = self.mass_mg_input.text().strip()

        if not sample_number:
            QMessageBox.warning(self, "Input Error", "Please enter a Sample Number for the export.")
            return

        try:
            float(dilution_text) if dilution_text else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid Dilution Factor. Cannot export.")
            return
        try:
            float(mass_mg_text) if mass_mg_text else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Invalid Mass (mg). Cannot export.")
            return

        export_data, has_calculable_data = self._collect_export_rows()
        if not has_calculable_data:
            reply = QMessageBox.question(
                self, 'Export Warning', "No valid calculated results found. Export anyway?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        elif not export_data:
            QMessageBox.warning(self, "Export Error", "No data to export.")
            return

        df_results = pd.DataFrame(export_data)
        df_results = df_results[["Analyte Name", "Analyte Amount", "LOQ", "State Limit", "Final Result", "Status"]]

        today_date = datetime.date.today().strftime("%Y%m%d")
        safe_sample_number = "".join(c for c in sample_number if c.isalnum() or c in ('_', '-')).rstrip()
        if not safe_sample_number:
            safe_sample_number = "NoSampleNum"
        default_filename = f"{today_date}_{safe_sample_number}_PSQuants.xlsx"

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(
            self, "Save Export File", default_filename,
            "Excel Files (*.xlsx);;All Files (*)", options=options
        )
        if fileName:
            try:
                self._write_export_excel(fileName, df_results)
                QMessageBox.information(self, "Success", f"Results successfully exported to:\n{fileName}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Could not save the file.\nError: {e}")
        else:
            QMessageBox.information(self, "Cancelled", "Export cancelled.")

    # ============================
    # Export SIN diálogo (para batch)
    # ============================
    def export_results_to_path(self, file_path):
        sample_number = self.sample_input.text().strip()
        dilution_text = self.dilution_input.text().strip()
        mass_mg_text = self.mass_mg_input.text().strip()

        if not sample_number:
            raise ValueError("Sample Number is required for export.")

        try:
            float(dilution_text) if dilution_text else 0.0
        except ValueError:
            raise ValueError("Invalid Dilution Factor. Cannot export.")
        try:
            float(mass_mg_text) if mass_mg_text else 0.0
        except ValueError:
            raise ValueError("Invalid Mass (mg). Cannot export.")

        export_data, _ = self._collect_export_rows()
        if not export_data:
            raise ValueError("No data to export.")

        df_results = pd.DataFrame(export_data)
        df_results = df_results[["Analyte Name", "Analyte Amount", "LOQ", "State Limit", "Final Result", "Status"]]
        self._write_export_excel(file_path, df_results)

    def _collect_export_rows(self):
        export_data = []
        has_calculable_data = False
        for row, analyte_name in enumerate(ANALYTES):
            amount_widget = self.analyte_amount_inputs[analyte_name]
            amount_text = amount_widget.text().strip()
            try:
                analyte_amount = float(amount_text) if amount_text else 0.0
            except ValueError:
                continue

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
        return export_data, has_calculable_data

    def _write_export_excel(self, file_path, df_results):
        sample_number = self.sample_input.text().strip()
        dilution_text = self.dilution_input.text().strip()
        mass_mg_text = self.mass_mg_input.text().strip()
        dilution_factor = float(dilution_text) if dilution_text else 0.0
        mass_in_grams = (float(mass_mg_text) / 1000.0) if mass_mg_text else 0.0

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            sample_info = {
                "Sample Number:": sample_number,
                "Client Name:": self.client_name_input.text().strip(),
                "Sample Date:": self.sample_date_input.date().toString(Qt.ISODate),
                "Dilution Factor:": dilution_factor,    # sin forzar decimal
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
                for cell in column:
                    try:
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if cell_length > max_length:
                                max_length = cell_length
                    except:
                        pass
                adjusted_width = max_length + 2
                worksheet.column_dimensions[column_letter].width = adjusted_width

    # ============================
    # Utilidades UI
    # ============================
    def copy_final_results(self):
        final_results = []
        for row in range(self.analytes_table.rowCount()):
            item = self.analytes_table.item(row, 4)
            final_results.append(item.text() if item else "")
        if final_results:
            QApplication.clipboard().setText("\n".join(final_results))
            QMessageBox.information(self, "Copied", "Final results copied to clipboard.")
        else:
            QMessageBox.warning(self, "Nothing to Copy", "The results table is empty.")

    def copy_nd_results(self):
        nd_results = ["ND"] * len(ANALYTES)
        QApplication.clipboard().setText("\n".join(nd_results))
        QMessageBox.information(self, "Copied", f"Copied {len(ANALYTES)} 'ND' results to clipboard.")

    def clear_inputs(self):
        self.sample_input.clear()
        self.client_name_input.clear()
        self.sample_date_input.setDate(QDate.currentDate())
        self.dilution_input.clear()
        self.mass_mg_input.clear()
        for amount_widget in self.analyte_amount_inputs.values():
            amount_widget.setText("0")
        QMessageBox.information(self, "Cleared", "Input fields have been cleared.")

    def closeEvent(self, event):
        if self.db_conn:
            self.db_conn.close()
            print("Database connection closed.")
        event.accept()

    # ============================
    # NUEVAS FUNCIONES (Batch)
    # ============================
    @staticmethod
    def _extract_batch_date_from_path(file_path):
        """Devuelve la fecha YYYYMMDD embebida en el nombre del archivo o None."""
        base_name = os.path.basename(file_path)
        name_part, _ = os.path.splitext(base_name)
        match = re.search(r'([0-9]{8})', name_part)
        if match:
            return match.group(1)
        return None

    @staticmethod
    def _normalize_sample_id_text(x):
        """
        Normaliza el Sample Number:
        - Si viene como 14936.0 (float/string), devuelve '14936'
        - Si trae espacios, los recorta
        """
        if x is None:
            return ""
        s = str(x).strip()
        # quita .0 al final si es float-like
        if s.endswith(".0"):
            try:
                f = float(s)
                if f.is_integer():
                    return str(int(f))
            except:
                pass
        # intenta castear a num y detectar entero
        try:
            v = pd.to_numeric(s, errors='coerce')
            if pd.notna(v) and float(v).is_integer():
                return str(int(v))
        except:
            pass
        return s

    def _get_default_output_dir_today(self):
        """Devuelve ./Excel reports/<YYYYMMDD>, creándolo si no existe."""
        base_dir = os.path.join(os.getcwd(), "Excel reports")
        today_folder = datetime.date.today().strftime("%Y%m%d")
        out_dir = os.path.join(base_dir, today_folder)
        os.makedirs(out_dir, exist_ok=True)
        return out_dir

    def _ui_batch_from_excel_dialog(self):
        """Pide solo el Excel de entrada. La salida va a ./Excel reports/<YYYYMMDD>/"""
        xlsx_path, _ = QFileDialog.getOpenFileName(
            self, "Selecciona Excel de entrada",
            "", "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if not xlsx_path:
            return

        out_dir = self._get_default_output_dir_today()

        try:
            processed = self.batch_generate_reports_from_excel(
                xlsx_path=xlsx_path,
                output_dir=out_dir,
                limit_reports=DEFAULT_BATCH_LIMIT  # None => procesa todos
            )
            QMessageBox.information(
                self, "Batch completado",
                f"Se generaron {processed} reporte(s) en:\n{out_dir}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error en batch", str(e))

    def _read_raw_results_excel(self, xlsx_path):
        """
        Lee el Excel y normaliza columnas clave:
        A: sample, B: component, D: calc_conc, E: mass_mg, F: df, G: include (YES/NO)
        (OJO: 'RESULT' ya no existe; ahora G es el include)
        """
        try:
            df = pd.read_excel(xlsx_path, sheet_name=RAW_SHEET_NAME, engine='openpyxl')
        except Exception:
            df = pd.read_excel(xlsx_path, sheet_name=RAW_SHEET_NAME, header=None, engine='openpyxl')

        if df is None or df.empty:
            raise ValueError(f"No se pudo leer datos de la hoja '{RAW_SHEET_NAME}'.")

        # Necesitamos al menos hasta la columna G => 7 columnas
        if df.shape[1] < 7:
            raise ValueError("La hoja no tiene al menos 7 columnas (A..G). Verifica el formato.")

        norm = pd.DataFrame({
            'sample': df.iloc[:, 0],   # A
            'component': df.iloc[:, 1],# B
            'calc_conc': df.iloc[:, 3],# D (Amount)
            'mass_mg': df.iloc[:, 4],  # E
            'df': df.iloc[:, 5],       # F
            'include': df.iloc[:, 6],  # G (YES/NO)
        })

        # Normaliza
        norm['sample'] = norm['sample'].map(self._normalize_sample_id_text)
        norm['component'] = norm['component'].astype(str).str.strip()
        norm['calc_conc'] = pd.to_numeric(norm['calc_conc'], errors='coerce')
        norm['mass_mg'] = pd.to_numeric(norm['mass_mg'], errors='coerce')
        norm['df'] = pd.to_numeric(norm['df'], errors='coerce')
        norm['include'] = norm['include'].astype(str).str.strip().str.upper().isin(['YES', 'Y', 'TRUE', '1'])

        # Filtra filas con sample y component no vacíos
        norm = norm[(norm['sample'] != '') & (norm['component'] != '')]
        return norm

    @staticmethod
    def _map_component_to_analyte(component_name: str) -> str:
        """'Propiconazole 1' -> 'Propiconazole' (quita sufijo ' 1' si existe)."""
        component_name = component_name.strip()
        return component_name[:-2].strip() if component_name.endswith(" 1") else component_name

    def _fill_amounts_from_dict(self, analyte_to_amount):
        """Pone 0 en todos y luego llena los analitos presentes con sus Amounts."""
        for a_name, w in self.analyte_amount_inputs.items():
            w.blockSignals(True); w.setText("0"); w.blockSignals(False)
        for a_name, amount_val in analyte_to_amount.items():
            if a_name in self.analyte_amount_inputs:
                w = self.analyte_amount_inputs[a_name]
                w.blockSignals(True); w.setText(str(amount_val)); w.blockSignals(False)
        self._update_results_table()

    @staticmethod
    def _make_output_filename(sample_number, out_dir):
        date_str = datetime.date.today().strftime("%Y%m%d")
        safe_sample = "".join(c for c in str(sample_number) if c.isalnum() or c in ('_', '-')).rstrip() or "NoSampleNum"
        return os.path.join(out_dir, f"{date_str}_{safe_sample}_PSQuants.xlsx")

    def save_current_sample_silent(self):
        """Guarda la muestra actual en la BD sin diálogos (INSERT OR REPLACE)."""
        if not self.db_conn:
            return
        sample_number = self.sample_input.text().strip()
        if not sample_number:
            return
        sample_date_str = self.sample_date_input.date().toString(Qt.ISODate)
        db_key = f"{sample_date_str.replace('-', '')}_{sample_number}"
        client_name = self.client_name_input.text().strip()
        try:
            dilution_factor = float(self.dilution_input.text().strip() or 0.0)
            mass_mg = float(self.mass_mg_input.text().strip() or 0.0)
        except ValueError:
            dilution_factor, mass_mg = 0.0, 0.0

        analyte_amounts = {}
        for analyte_name, input_widget in self.analyte_amount_inputs.items():
            try:
                analyte_amounts[analyte_name] = float(input_widget.text().strip() or 0.0)
            except ValueError:
                analyte_amounts[analyte_name] = 0.0

        analyte_data_json = json.dumps(analyte_amounts)
        try:
            cursor = self.db_conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO samples
                (sample_number, original_sample_number, client_name, sample_date, dilution_factor, mass_mg, analyte_data)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (db_key, sample_number, client_name, sample_date_str, dilution_factor, mass_mg, analyte_data_json))
            self.db_conn.commit()
        except sqlite3.Error as e:
            print(f"[Batch save] DB error: {e}")

    def batch_generate_reports_from_excel(self, xlsx_path, output_dir, limit_reports=None):
        """
        Un reporte por sample (solo samples con al menos un componente include=YES):
        - Identifica de manera EXPLÍCITA la lista de samples únicos (sin inventar).
        - Si limit_reports es None -> procesa todos; si es un número, procesa hasta ese máximo.
        - Para cada sample:
            * Dentro del sample, filtra filas con include=YES (columna G).
            * Toma la PRIMERA ocurrencia por analito (si hay X 1 y X 2, se queda solo con el primero).
            * Amount = col D (calc_conc).
            * Mass (mg) = col E; DF = col F (tal cual).
            * Final result lo calcula _update_results_table() con la nueva fórmula.
            * Exporta y guarda en BD en silencio.
        """
        if not os.path.isfile(xlsx_path):
            raise FileNotFoundError(f"Archivo no encontrado: {xlsx_path}")
        if not os.path.isdir(output_dir):
            raise NotADirectoryError(f"Carpeta de salida inválida: {output_dir}")

        df = self._read_raw_results_excel(xlsx_path)

        batch_date = None
        date_str = self._extract_batch_date_from_path(xlsx_path)
        if date_str:
            parsed_date = QDate.fromString(date_str, "yyyyMMdd")
            if parsed_date.isValid():
                batch_date = parsed_date

        # 1) Filtra SOLO filas YES
        df_yes = df[df['include']].copy()
        if df_yes.empty:
            raise ValueError("No hay filas con include=YES en la hoja seleccionada.")

        # 2) Lista explícita de samples únicos (en orden de aparición), normalizados
        samples_unique = list(dict.fromkeys(df_yes['sample'].tolist()))
        if not samples_unique:
            self.load_samples_table()
            return 0

        # Si se especifica un límite, recorta la lista; si no, procesa todos
        if limit_reports is not None:
            try:
                nmax = int(limit_reports)
                samples_unique = samples_unique[:max(0, nmax)]
            except:
                pass  # si el límite no es válido, ignora y procesa todos

        processed = 0
        for sample in samples_unique:
            # Subconjunto del sample actual con YES
            sub = df_yes[df_yes['sample'] == sample]
            if sub.empty:
                continue

            # Mass (mg) y DF por muestra (primer no nulo)
            mass_val = sub['mass_mg'].dropna().iloc[0] if sub['mass_mg'].dropna().size > 0 else None
            df_val = sub['df'].dropna().iloc[0] if sub['df'].dropna().size > 0 else None

            # Mapea componente -> analito base y QUÉDATE con la PRIMERA ocurrencia por analito
            sub = sub.copy()
            sub['analyte_base'] = sub['component'].astype(str).map(self._map_component_to_analyte)
            dedup = sub.dropna(subset=['calc_conc']).drop_duplicates(subset=['analyte_base'], keep='first')

            analyte_to_amount = {}
            for _, row in dedup.iterrows():
                analyte_name = str(row['analyte_base'])
                val = row['calc_conc']
                if pd.isna(val):
                    continue
                if analyte_name in self.analyte_amount_inputs:
                    analyte_to_amount[analyte_name] = float(val)

            if not analyte_to_amount:
                continue

            # ======= RELLENAR UI PARA ESTE SAMPLE =======
            safe_sample = self._normalize_sample_id_text(sample)
            self.sample_input.setText(safe_sample)
            if batch_date is not None:
                self.sample_date_input.setDate(batch_date)
            else:
                self.sample_date_input.setDate(QDate.currentDate())

            # Mass (mg) tal cual
            if mass_val is not None and not pd.isna(mass_val):
                self.mass_mg_input.setText(f"{float(mass_val):g}")
            else:
                self.mass_mg_input.clear()

            # Dilution Factor tal cual
            if df_val is not None and not pd.isna(df_val):
                self.dilution_input.setText(f"{float(df_val):g}")
            else:
                self.dilution_input.clear()

            # Amounts por analito (desde col D)
            self._fill_amounts_from_dict(analyte_to_amount)

            # ======= EXPORTAR =======
            out_path = self._make_output_filename(sample_number=safe_sample, out_dir=output_dir)
            self.export_results_to_path(out_path)

            # ======= GUARDAR EN BD (silencioso) =======
            self.save_current_sample_silent()

            processed += 1

        self.load_samples_table()
        return processed


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PSCalculatorApp()
    sys.exit(app.exec_())
