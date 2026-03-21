#!/usr/bin/env python3
"""
ugeto_ui.py
PyQt6 GUI for the Ügető program (titles + jockeys editing, load PDF placeholder, save CSV, make PPT)

Dependencies:
  pip install PyQt6 pandas python-pptx

Run:
  python ugeto_ui.py

This file implements a self-contained UI. PDF parsing and advanced PPT layout are left as clear hooks
so you can plug your existing extraction logic later.
"""

import sys
import os
import csv
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QMessageBox, QLabel, QLineEdit, QTabWidget,
    QTableView, QHeaderView, QAbstractItemView, QToolBar, QStatusBar
)
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtCore import Qt, QSortFilterProxyModel, QModelIndex
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from modules import *
import requests
# Optional libraries
try:
    import pandas as pd
except Exception:
    pd = None
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
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
            QAbstractItemView.EditTrigger.DoubleClicked |
            QAbstractItemView.EditTrigger.EditKeyPressed |
            QAbstractItemView.EditTrigger.AnyKeyPressed
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

        # Toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        load_pdf_act = QAction("Load PDF", self)
        load_pdf_act.setShortcut(QKeySequence("Ctrl+O"))
        load_pdf_act.triggered.connect(lambda: self.load_pdf(self))

        save_csv_act = QAction("Save Data to CSV", self)
        save_csv_act.setShortcut(QKeySequence("Ctrl+S"))
        save_csv_act.triggered.connect(lambda: self.save_csv(self))

        make_ppt_act = QAction("Make PPT", self)
        make_ppt_act.triggered.connect(lambda: self.make_ppt(self))

        toolbar.addAction(load_pdf_act)
        toolbar.addAction(save_csv_act)
        toolbar.addAction(make_ppt_act)

        # Tabs for Titles and jockeys
        self.tabs = QTabWidget()

        self.titles_header  = ["Id","Daily", "Title", "Distance", "Start time", "Track", "Opinion"]
        self.jockeys_header = ['Start number', 'Horse name', 'Jockey name', 'Weight', 'Allowance', 'Futam id', 'IsRun'] #["Start number", "Horse name", "Distance", "Driver name", "Futam id", "Run"]

        self.titles_widget  = EditableTable(self.titles_header)
        self.jockeys_widget = EditableTable(self.jockeys_header)

        self.tabs.addTab(self.titles_widget, "Titles")
        self.tabs.addTab(self.jockeys_widget, "Jockeys")

        # Central layout
        central = QWidget()
        main_layout = QVBoxLayout(central)
        main_layout.addWidget(self.tabs)

        # Info label and status bar
        info = QLabel("PDF betöltése után töltsd fel az adatokat, szerkeszd, majd mentsd CSV-be és készíts PPT-t.")
        main_layout.addWidget(info)

        self.setCentralWidget(central)
        self.status = QStatusBar()
        self.setStatusBar(self.status)

        # Shortcuts: pressing any printable key will insert into selected cell
        self.table_widgets = [self.titles_widget.table, self.jockeys_widget.table]

    def load_pdf(self,mainWindow):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF files (*.pdf);;All files (*)")
        if not path:
            return
        self.status.showMessage(f"Betöltött PDF: {path}")

        pdf_data = ReadPDF(mainWindow=mainWindow,file_name=path)
        self.titles = pdf_data.futams
        self.jockeys = pdf_data.horses

        """ 
        for jockey in self.jockeys:
            print(jockey.__dict__())
        """

        self.titles_columns = [
            ("Id", "id"),
            ("Daily", "daily"),
            ("Title", "title"),
            ("Distance", "dist"),
            ("Start time", "time"),
            ("Track", "track"),
            ("Opinion", "opinion")
        ]
        
        self.jockeys_columns = [
            ("Start number", "Hnum"),
            ("Horse name", "Hname"),
            ("Jockey name", "DJname"),
            ("Weight","weight"),
            ("Allowance","allowance"),
            ("Futam id", "Fnum"),
            ("IsRun", "isRun")
        ]
        
        #load_from_list_of_obj
        self.titles_widget.load_from_objects(self.titles,self.titles_columns)
        #load_from_list_of_obj
        self.jockeys_widget.load_from_objects(self.jockeys,self.jockeys_columns)

        self.titles_widget.table.resizeColumnsToContents()
        self.titles_widget.table.resizeRowsToContents()

        self.jockeys_widget.table.resizeColumnsToContents()
        self.jockeys_widget.table.resizeRowsToContents()
        QMessageBox.information(self, "PDF betöltve", "A PDF feldolgozása kész")

    def save_csv(self,mainWindow):
        # Ensure directories exist
        CSV_DIR.mkdir(exist_ok=True)
        if not self.titles_widget.to_list_of_dicts() and not self.jockeys_widget.to_list_of_dicts():
            self.load_pdf(mainWindow)

        titles = self.titles_widget.to_list_of_dicts()
        jockeys = self.jockeys_widget.to_list_of_dicts()
        self.titles = []
        self.jockeys = []

        for ln in jockeys: 
            driver = Horses()
            self.jockeys.append(driver.load_dict(ln))

        
        for ln in titles: 
            title = Futam()
            self.titles.append(title.load_dict(ln))
        
        title_header = ";".join(self.titles_header)+"\n"
        driver_header = ";".join(self.jockeys_header)+"\n"

        with open("./csv/titles_data.csv",'w',encoding="utf-8") as f:
            f.write(title_header)
            for ln in self.titles:
                f.write(str(ln)+"\n")

        with open("./csv/jockeys_data.csv",'w',encoding="utf-8") as f:
            f.write(driver_header)
            for ln in self.jockeys:
                f.write(str(ln)+"\n")
        

    def make_ppt(self,mainWindow):
        PPT_DIR.mkdir(exist_ok=True)
        
        if not self.titles_widget.to_list_of_dicts() and not self.jockeys_widget.to_list_of_dicts() and os.path.exists("./csv/jockeys_data.csv") and os.path.exists("./csv/titles_data.csv"):
            titles = []
            jockeys = []
            with open("./csv/titles_data.csv",'r',encoding="utf-8") as f:
                fs = f.readline()
                for ln in f: 
                    #print(ln.strip())
                    titles.append(Futam(ln))

            with open("./csv/jockeys_data.csv",'r',encoding="utf-8") as f:
                fs = f.readline()
                for ln in f: jockeys.append(Horses(ln))

            MakePPT(jockeys,titles)

        elif(self.titles_widget.to_list_of_dicts() and self.jockeys_widget.to_list_of_dicts()):
            titles = self.titles_widget.to_list_of_dicts()
            jockeys = self.jockeys_widget.to_list_of_dicts()
            self.titles = []
            self.jockeys = []

            for ln in jockeys: 
                driver = Horses()
                self.jockeys.append(driver.load_dict(ln))

            for ln in titles: 
                title = Futam()
                self.titles.append(title.load_dict(ln))

            MakePPT(self.jockeys,self.titles)
        else:
            self.load_pdf(mainWindow)
            self.save_csv(mainWindow)
            titles = self.titles_widget.to_list_of_dicts()
            jockeys = self.jockeys_widget.to_list_of_dicts()
            self.titles = []
            self.jockeys = []

            for ln in jockeys: 
                driver = Horses()
                self.jockeys.append(driver.load_dict(ln))

            for ln in titles: 
                title = Futam()
                self.titles.append(title.load_dict(ln))

            MakePPT(self.jockeys,self.titles)
            
    # Optional: override keyPressEvent to insert text to the selected cell when typing
    def keyPressEvent(self, event):
        key = event.key()
        if event.text() and event.text().isprintable():
            for table in self.table_widgets:
                if table.hasFocus():
                    sel = table.selectionModel().currentIndex()
                    if sel.isValid():
                        src = table.model().mapToSource(sel) if isinstance(table.model(), QSortFilterProxyModel) else sel
                        item = self.titles_widget.model.itemFromIndex(src) if table is self.titles_widget.table else self.jockeys_widget.model.itemFromIndex(src)
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
            if os.path.exists("./add_macro.xlsm"):
                w.show()
            else:
                QMessageBox.critical(w,"Error","Can't find \"add_macro.xlsm\"")
        else:
            QMessageBox.critical(w,"Error","Can't find \"clock.jpeg\"")

    else:
        QMessageBox.critical(w,"Error","Can't connect to https://mla.kincsempark.hu/")
        return

    sys.exit(app.exec())

if __name__ == '__main__':
    main()
