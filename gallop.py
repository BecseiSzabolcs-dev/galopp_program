#!/usr/bin/env python3
"""
gallop.py
PyQt6 GUI for the Ügető program (titles + jockeys editing, load PDF placeholder, save CSV, make PPT)

Dependencies:
  pip install PyQt6 pandas python-pptx

Run:
  python gallop.py

This file implements a self-contained UI. PDF parsing and advanced PPT layout are left as clear hooks
so you can plug your existing extraction logic later.
"""

import os
import sys
from pathlib import Path

import requests
from PyQt6.QtCore import QSortFilterProxyModel, Qt
from PyQt6.QtGui import QAction, QKeySequence, QStandardItem, QStandardItemModel
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QTableView,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from modules import *

# Optional libraries
try:
    import pandas as pd
except Exception:
    pd = None
try:
    from pptx import Presentation
except Exception:
    Presentation = None


CSV_DIR = Path("csv")
PPT_DIR = Path("ppt")
CSV_DIR.mkdir(exist_ok=True)
PPT_DIR.mkdir(exist_ok=True)


class EditableTable(QWidget):
    """Reusable Excel-like editable table with search, add, delete, resizable columns"""

    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.columns = columns
        self.model = QStandardItemModel(0, len(columns), self)
        self.model.setHorizontalHeaderLabels(columns)

        # --- FILTER + SORT ---
        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)

        # --- TABLE ---
        self.table = QTableView()
        self.table.setModel(self.proxy)

        # Excel-like full manual resize (NOW WORKS)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(40)

        # Selection + editing behavior
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )
        self.table.verticalHeader().setVisible(True)  # show row numbers

        # --- SEARCH field ---
        self.search = QLineEdit()
        self.search.setPlaceholderText("Keresés…")
        self.search.textChanged.connect(self.on_search)

        # --- BUTTONS ---
        add_btn = QPushButton("Sor hozzáadása")
        del_btn = QPushButton("Sor törlése")
        add_btn.clicked.connect(lambda: self.add_row())
        del_btn.clicked.connect(self.delete_selected_row)

        # Layout top row (search + buttons)
        h = QHBoxLayout()
        h.addWidget(self.search)
        h.addWidget(add_btn)
        h.addWidget(del_btn)

        layout = QVBoxLayout(self)
        layout.addLayout(h)
        layout.addWidget(self.table)

    # -----------------------------------
    # FILTER
    # -----------------------------------
    def on_search(self, text):
        self.proxy.setFilterFixedString(text)

    # -----------------------------------
    # ADD ROW SAFELY (supports sorting!)
    # -----------------------------------
    def add_row(self, values=None):
        if values is None:
            values = ["" for _ in self.columns]

        # Insert into source model (not proxy)
        row_items = [QStandardItem(str(v)) for v in values]
        for it in row_items:
            it.setEditable(True)

        self.model.appendRow(row_items)

        # Scroll + focus new row
        source_idx = self.model.index(self.model.rowCount() - 1, 0)
        proxy_idx = self.proxy.mapFromSource(source_idx)
        self.table.scrollTo(proxy_idx)
        self.table.setCurrentIndex(proxy_idx)

    # -----------------------------------
    # DELETE ROW (works even when sorted)
    # -----------------------------------
    def delete_selected_row(self):
        sel = self.table.selectionModel().currentIndex()
        if not sel.isValid():
            QMessageBox.information(self, "Törlés", "Nincs kiválasztott sor.")
            return

        source_idx = self.proxy.mapToSource(sel)
        self.model.removeRow(source_idx.row())

    # -----------------------------------
    # EXPORT TABLE CONTENT
    # -----------------------------------
    def to_list_of_dicts(self):
        rows = []
        for r in range(self.model.rowCount()):
            d = {}
            for c, col_name in enumerate(self.columns):
                item = self.model.item(r, c)
                d[col_name] = item.text() if item else ""
            rows.append(d)
        return rows

    # -----------------------------------
    # LOAD OBJECTS AUTOMATICALLY
    # -----------------------------------
    def load_from_objects(self, data, mapping):
        self.model.removeRows(0, self.model.rowCount())

        for obj in data:
            values = [str(getattr(obj, attr, "")) for _, attr in mapping]
            self.add_row(values)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gallop program")
        self.resize(1100, 700)

        self.titles: list = []
        self.jockeys: list = []
        self.addons: list = []

        # Toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        # Load PDF
        load_pdf_act = QAction("Load PDF", self)
        load_pdf_act.setShortcut(QKeySequence("Ctrl+O"))
        load_pdf_act.triggered.connect(lambda: self.load_pdf(self))

        # Save to csv
        save_csv_act = QAction("Save Data to CSV", self)
        save_csv_act.setShortcut(QKeySequence("Ctrl+S"))
        save_csv_act.triggered.connect(lambda: self.save_csv(self))

        # Load previus CSV
        load_prev_act = QAction("Load previus data", self)
        load_prev_act.setShortcut(QKeySequence("Ctrl+T"))
        load_prev_act.triggered.connect(lambda: self.Load_previus(self))

        # Add Plus content for Powerpoints
        add_addon_act = QAction("Add Addon", self)
        add_addon_act.triggered.connect(lambda: self.Add_addon(self))
        add_addon_act.setShortcut(QKeySequence("A"))

        # create Porwerpoints
        make_ppt_act = QAction("Make PPT", self)
        make_ppt_act.triggered.connect(lambda: self.make_ppt(self))
        make_ppt_act.setShortcut(QKeySequence("M"))

        # Load toolbar buttons
        toolbar.addAction(load_pdf_act)
        toolbar.addAction(save_csv_act)
        toolbar.addAction(load_prev_act)
        toolbar.addAction(add_addon_act)
        toolbar.addAction(make_ppt_act)

        # Tabs for Titles and jockeys
        self.tabs = QTabWidget()

        self.titles_header = [
            "Id",
            "Daily",
            "Title",
            "Distance",
            "Start time",
            "Track",
            "Opinion",
        ]
        self.titles_columns = [
            ("Id", "id"),
            ("Daily", "daily"),
            ("Title", "title"),
            ("Distance", "dist"),
            ("Start time", "time"),
            ("Track", "track"),
            ("Opinion", "opinion"),
        ]
        self.jockeys_header = [
            "Start number",
            "Horse name",
            "Jockey name",
            "Weight",
            "Allowance",
            "Futam id",
            "IsRun",
        ]  # ["Start number", "Horse name", "Distance", "jockey name", "Futam id", "Run"]
        self.jockeys_columns = [
            ("Start number", "Hnum"),
            ("Horse name", "Hname"),
            ("Jockey name", "DJname"),
            ("Weight", "weight"),
            ("Allowance", "allowance"),
            ("Futam id", "Fnum"),
            ("IsRun", "isRun"),
        ]

        self.addons_header = ["Id", "Name", "Type", "Path"]
        self.addons_columns = [
            ("Id", "id"),
            ("Name", "name"),
            ("Type", "type"),
            ("Path", "path"),
        ]

        self.titles_widget = EditableTable(self.titles_header)
        self.jockeys_widget = EditableTable(self.jockeys_header)
        self.addons_widget = EditableTable(self.addons_header)

        self.tabs.addTab(self.titles_widget, "Titles")
        self.tabs.addTab(self.jockeys_widget, "Jockeys")
        self.tabs.addTab(self.addons_widget, "Addons")

        # Central layout
        central = QWidget()
        main_layout = QVBoxLayout(central)
        main_layout.addWidget(self.tabs)

        # Info label and status bar
        info = QLabel(
            "PDF betöltése után töltsd fel az adatokat, szerkeszd, majd mentsd CSV-be és készíts PPT-t."
        )
        main_layout.addWidget(info)

        self.setCentralWidget(central)
        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # Shortcuts: pressing any printable key will insert into selected cell
        self.table_widgets = [self.titles_widget.table, self.jockeys_widget.table]

    def load_pdf(self, mainWindow):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open PDF",
            str(Path.home() / "Downloads"),
            "PDF files (*.pdf);;All files (*)",
        )
        if not path:
            return
        self.status.showMessage(f"Betöltött PDF: {path}")

        pdf_data = ReadPDF(mainWindow=mainWindow, file_name=path)
        self.titles = pdf_data.futams
        self.jockeys = pdf_data.horses

        """ 
        for jockey in self.jockeys:
            print(jockey.__dict__())
        """

        # load_from_list_of_obj
        self.titles_widget.load_from_objects(self.titles, self.titles_columns)
        # load_from_list_of_obj
        self.jockeys_widget.load_from_objects(self.jockeys, self.jockeys_columns)

        self.titles_widget.table.resizeColumnsToContents()
        self.titles_widget.table.resizeRowsToContents()

        self.jockeys_widget.table.resizeColumnsToContents()
        self.jockeys_widget.table.resizeRowsToContents()
        QMessageBox.information(self, "PDF betöltve", "A PDF feldolgozása kész")

    def save_csv(self, mainWindow):
        # Ensure directories exist
        try:
            CSV_DIR.mkdir(exist_ok=True)
            if (
                not self.titles_widget.to_list_of_dicts()
                and not self.jockeys_widget.to_list_of_dicts()
            ):
                self.load_pdf(mainWindow)

            titles = self.titles_widget.to_list_of_dicts()
            jockeys = self.jockeys_widget.to_list_of_dicts()
            addons = self.addons_widget.to_list_of_dicts()

            if not (titles and jockeys):
                raise ValueError("The titles or jockeys tables are empty")

            self.addons = []
            self.titles = []
            self.jockeys = []

            for ln in jockeys:
                jockey = Horses()
                self.jockeys.append(jockey.load_dict(ln))

            for ln in titles:
                title = Futam()
                self.titles.append(title.load_dict(ln))

            for ln in addons:
                print(ln)
                self.addons.append(Addon(ln))

            title_header = ";".join(self.titles_header) + "\n"
            jockey_header = ";".join(self.jockeys_header) + "\n"

            addons_header = ";".join(self.addons_header) + "\n"

            with open("./csv/titles_data.csv", "w", encoding="utf-8") as f:
                f.write(title_header)
                for ln in self.titles:
                    f.write(str(ln) + "\n")

            with open("./csv/jockeys_data.csv", "w", encoding="utf-8") as f:
                f.write(jockey_header)
                for ln in self.jockeys:
                    f.write(str(ln) + "\n")
            if addons != []:
                with open("./csv/addons_data.csv", "w", encoding="utf-8") as f:
                    f.write(addons_header)
                    for ln in self.addons:
                        f.write(str(ln) + "\n")

            QMessageBox.information(self, "CSV fájlok", "A CSV fájlok lementése kész")

        except (ValueError, TypeError) as e:
            print("Somthing went wrong when tried to create CSV file!")
            print(f"Error {e}")

    def make_ppt(self, mainWindow):
        PPT_DIR.mkdir(exist_ok=True)

        ppt_path = QFileDialog.getExistingDirectory(
            None, "Select Folder", "/run/media/"
        )

        try:
            if (
                self.titles_widget.to_list_of_dicts()
                and self.jockeys_widget.to_list_of_dicts()
            ):
                titles = self.titles_widget.to_list_of_dicts()
                jockeys = self.jockeys_widget.to_list_of_dicts()
                addons = self.addons_widget.to_list_of_dicts()
                self.titles = []
                self.jockeys = []
                self.addons = []

                for ln in addons:
                    jockey = Horses()
                    self.addons.append(Addon(ln))

                for ln in jockeys:
                    jockey = Horses()
                    self.jockeys.append(jockey.load_dict(ln))

                for ln in titles:
                    title = Futam()
                    self.titles.append(title.load_dict(ln))

                MakePPT(self, self.jockeys, self.titles, self.addons, ppt_path)

            elif (
                os.path.exists("./csv/jockeys_data.csv")
                and os.path.exists("./csv/titles_data.csv")
                and os.path.exists("./csv/addons_data.csv")
            ):
                self.addons = []
                self.titles = []
                self.jockeys = []

                with open("./csv/addons_data.csv", "r", encoding="utf-8") as f:
                    fs = f.readline()
                    for ln in f:
                        # print(ln.strip())
                        self.addons.append(Addon(ln))

                with open("./csv/titles_data.csv", "r", encoding="utf-8") as f:
                    fs = f.readline()
                    for ln in f:
                        # print(ln.strip())
                        self.titles.append(Futam(ln))

                with open("./csv/jockeys_data.csv", "r", encoding="utf-8") as f:
                    fs = f.readline()
                    for ln in f:
                        self.jockeys.append(Horses(ln))

                MakePPT(self, self.jockeys, self.titles, self.addons, ppt_path)

            else:
                self.load_pdf(mainWindow)
                self.save_csv(mainWindow)
                titles = self.titles_widget.to_list_of_dicts()
                jockeys = self.jockeys_widget.to_list_of_dicts()
                self.titles = []
                self.jockeys = []

                for ln in jockeys:
                    jockey = Horses()
                    self.jockeys.append(jockey.load_dict(ln))

                for ln in titles:
                    title = Futam()
                    self.titles.append(title.load_dict(ln))

                MakePPT(self, self.jockeys, self.titles, self.addons, ppt_path)

        except Exception as e:
            QMessageBox.critical(
                self, "PPT file creation", f"The powerpoint files creation faild:\n{e}"
            )

    def Load_previus(self, mainWindow):
        if os.path.exists("./csv/jockeys_data.csv") and os.path.exists(
            "./csv/titles_data.csv"
        ):
            self.titles = []
            self.jockeys = []

            with open("./csv/titles_data.csv", "r", encoding="utf-8") as f:
                fs = f.readline()
                for ln in f:
                    # print(ln.strip())
                    self.titles.append(Futam(ln))

            with open("./csv/jockeys_data.csv", "r", encoding="utf-8") as f:
                fs = f.readline()
                for ln in f:
                    self.jockeys.append(Horses(ln))

            self.titles_widget.load_from_objects(self.titles, self.titles_columns)
            self.jockeys_widget.load_from_objects(self.jockeys, self.jockeys_columns)

            self.titles_widget.table.resizeColumnsToContents()
            self.titles_widget.table.resizeRowsToContents()

            self.jockeys_widget.table.resizeColumnsToContents()
            self.jockeys_widget.table.resizeRowsToContents()

            if os.path.exists("./csv/addons_data.csv_data.csv"):
                self.addons = []
                with open("./csv/addons_data.csv", "r", encoding="utf-8") as f:
                    fs = f.readline()
                    for ln in f:
                        self.addons.append(Horses(ln))

                self.titles_widget.load_from_objects(self.addons, self.addons_columns)

                self.addons_widget.table.resizeColumnsToContents()
                self.addons_widget.table.resizeRowsToContents()

            QMessageBox.information(
                self, "Previus files loading", "The previus data is loaded succesfully"
            )

        else:
            futam_path, _ = QFileDialog.getOpenFileName(
                self,
                "Open Titles data",
                "",
                "CSV files (*.csv);;TEXT files (*.txt);;All files (*)",
            )
            jockeys_path, _ = QFileDialog.getOpenFileName(
                self,
                "Open Jockeys data",
                "",
                "CSV files (*.csv);;TEXT files (*.txt);;All files (*)",
            )
            if os.path.exists(jockeys_path) and os.path.exists(futam_path):
                self.titles = []
                self.jockeys = []
                with open(futam_path, "r", encoding="utf-8") as f:
                    fs = f.readline()
                    for ln in f:
                        # print(ln.strip())
                        self.titles.append(Futam(ln))

                with open(jockeys_path, "r", encoding="utf-8") as f:
                    fs = f.readline()
                    for ln in f:
                        self.jockeys.append(Horses(ln))

                self.titles_widget.load_from_objects(self.titles, self.titles_columns)
                # load_from_list_of_obj
                self.jockeys_widget.load_from_objects(
                    self.jockeys, self.jockeys_columns
                )

                self.titles_widget.table.resizeColumnsToContents()
                self.titles_widget.table.resizeRowsToContents()

                self.jockeys_widget.table.resizeColumnsToContents()
                self.jockeys_widget.table.resizeRowsToContents()

                QMessageBox.information(
                    self,
                    "Previus files loading",
                    "The previus data is loaded succesfully",
                )
            else:
                QMessageBox.critical(
                    self, "Previus files loading", "The some off the files not exists"
                )

    def Add_addon(self, mainWindow):
        dialog = AddonSel(self)

        if dialog.exec():
            self.addons = dialog.get_files()

            if self.addons:
                self.addons_widget.load_from_objects(self.addons, self.addons_columns)
                self.addons_widget.table.resizeColumnsToContents()
                self.addons_widget.table.resizeRowsToContents()

                QMessageBox.information(
                    self, "Addons load", "The Addons loaded successfully."
                )

    # Optional: override keyPressEvent to insert text to the selected cell when typing
    def keyPressEvent(self, event):
        key = event.key()
        if event.text() and event.text().isprintable():
            for table in self.table_widgets:
                if table.hasFocus():
                    sel = table.selectionModel().currentIndex()
                    if sel.isValid():
                        src = (
                            table.model().mapToSource(sel)
                            if isinstance(table.model(), QSortFilterProxyModel)
                            else sel
                        )
                        item = (
                            self.titles_widget.model.itemFromIndex(src)
                            if table is self.titles_widget.table
                            else self.jockeys_widget.model.itemFromIndex(src)
                        )
                        if item is not None:
                            item.setText(item.text() + event.text())
                            return
        super().keyPressEvent(event)


def is_connected():
    try:
        response = requests.get("https://mla.kincsempark.hu/", timeout=5)
        return True
    except requests.ConnectionError:
        return False


def main():
    app = QApplication(sys.argv)

    w = MainWindow()
    if is_connected():
        if os.path.exists("./clock.jpeg"):
            w.show()
        else:
            QMessageBox.critical(w, "Error", 'Can\'t find "clock.jpeg"')

    else:
        QMessageBox.critical(w, "Error", "Can't connect to https://mla.kincsempark.hu/")
        return

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
