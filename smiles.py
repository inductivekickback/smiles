"""
An optimized GUI developed by Rebecca Medley for the purpose of making her colleagues' lives
easiser by simplifying the process of filling in a specific 'Mileage Reimbursement' form.


Disclaimer of Warranty

This software is provided "as is," without any warranties of any kind, either express or implied,
including but not limited to the implied warranties of merchantability, fitness for a particular
purpose, or non-infringement. The entire risk arising out of the use or performance of the software
remains with you. In no event shall the software provider be liable for any direct, indirect,
incidental, special, consequential, or punitive damages whatsoever arising out of or in connection
with the use or inability to use the software.
"""
import sys
import pickle
import os
import platform
import json
import random
from PyQt6.QtGui import (QAction, QIcon, QColor, QPainter, QPen, QDoubleValidator, QPixmap,
                                QFont, QDragEnterEvent, QDropEvent)
from PyQt6.QtCore import Qt, QDate, QEvent
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTableWidget, QDateEdit, QDialog,
                                QMessageBox, QLineEdit, QCompleter, QPushButton, QVBoxLayout,
                                QHBoxLayout, QWidget, QFormLayout, QLabel, QDialogButtonBox,
                                QFileDialog, QStyle, QStyleOption, QFrame)

from pdf_writer import fill_form


__version__ = "1.4.3"
__date__ = "Sept '24"

APP_NAME = "Smiles"
APP_EXT = "rlm"
BASE_DIR = os.path.dirname(__file__)
ARTEFACTS_DIR = "artefacts"
DATA_FILE = "data.pickle"
ICON_FILE = "mentor.png"
FORM_FILE = "mileage.pdf"
ABOUT_IMG_PATH = "support.png"


class CustomLineEdit(QLineEdit):
    """Overrides the default paintEvent to draw box outlines in a specified color."""

    COLORS = [QColor("#40FF0018"), QColor("#40FFA52C"), QColor("#40FFFF41"), QColor("#40008018"),
        QColor("#400000F9"), QColor("#4086007D")]

    def __init__(self, row_index, parent=None):
        super().__init__(parent)
        self.color = self.COLORS[row_index % len(self.COLORS)]

    def paintEvent(self, event):
        """Override the default event to use a custom color."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(self.color)
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawRect(self.rect().adjusted(0, 0, -1, -1))
        super().paintEvent(event)


class AboutDialog(QDialog):
    """A simple 'About' dialog."""

    QUOTES = (("Spread the weird!",),("We've got your back!",),
                    ("Stay curious!",), ("I shan't be doing that.",))
    QUOTE_AUTHOR = "-- The Mentor Team"

    def __init__(self, data_date=None):
        super().__init__()

        self.setWindowTitle(f"About {APP_NAME}")

        left_layout = QVBoxLayout()

        label = QLabel(self)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setFixedSize(166, 256)
        pixmap = QPixmap(os.path.join(BASE_DIR, ARTEFACTS_DIR, ABOUT_IMG_PATH))
        scaled_pixmap = pixmap.scaled(label.size(),
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation)
        label.setPixmap(scaled_pixmap)
        left_layout.addWidget(label)

        label = QLabel("medley_r@4j.lane.edu", alignment=Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(label)
        left_layout.addStretch(1)

        default_font_size = label.font().pointSize()

        font = QFont()
        font.setPointSize(default_font_size - 1)

        label.setFont(font)

        right_layout = QVBoxLayout()

        label = QLabel(f"{APP_NAME} v{__version__}",
                            alignment=Qt.AlignmentFlag.AlignCenter)
        font.setPointSize(default_font_size + 8)
        label.setFont(font)
        right_layout.addWidget(label)

        label = QLabel("by Rebecca Medley", alignment=Qt.AlignmentFlag.AlignCenter)
        font.setPointSize(default_font_size + 5)
        label.setFont(font)
        right_layout.addWidget(label)

        right_layout.addSpacing(30)

        hbox = QHBoxLayout()
        hbox.setSpacing(0)
        hbox.setContentsMargins(0, 0, 0, 0)

        vbox = QVBoxLayout()
        vbox.setSpacing(2)
        vbox.setContentsMargins(0, 0, 5, 0)
        label = QLabel("Application:", alignment=Qt.AlignmentFlag.AlignRight)
        vbox.addWidget(label)
        label = QLabel("Distance calculation:", alignment=Qt.AlignmentFlag.AlignRight)
        vbox.addWidget(label)
        widget = QWidget()
        widget.setLayout(vbox)
        hbox.addWidget(widget)

        vbox = QVBoxLayout()
        vbox.setSpacing(2)
        vbox.setContentsMargins(0, 0, 0, 0)
        label = QLabel("<a href='https://github.com/inductivekickback/smiles'>" +
            "github.com/inductivekickback/smiles</a>", alignment=Qt.AlignmentFlag.AlignLeft)
        label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction |
                                            Qt.TextInteractionFlag.LinksAccessibleByMouse)
        label.setOpenExternalLinks(True)
        vbox.addWidget(label)
        label = QLabel(f'(Built: {__date__})', alignment=Qt.AlignmentFlag.AlignLeft)
        vbox.addWidget(label)
        label = QLabel("<a href='https://github.com/inductivekickback/mileage'>" +
            "github.com/inductivekickback/mileage</a>", alignment=Qt.AlignmentFlag.AlignLeft)
        label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction |
                                            Qt.TextInteractionFlag.LinksAccessibleByMouse)
        label.setOpenExternalLinks(True)
        vbox.addWidget(label)
        date_text = "(Compiled: "
        if data_date:
            date_text += data_date.strftime("%b '%y") + ")"
        else:
            date_text += "---)"
        label = QLabel(f'{date_text}', alignment=Qt.AlignmentFlag.AlignLeft)
        vbox.addWidget(label)
        widget = QWidget()
        widget.setLayout(vbox)
        hbox.addWidget(widget)

        widget = QWidget()
        widget.setLayout(hbox)
        right_layout.addWidget(widget)

        quote_layout = QVBoxLayout(self)
        font.setPointSize(default_font_size + 10)
        quote = self.QUOTES[random.randrange(0, len(self.QUOTES))]
        for line in quote:
            label = QLabel(line , alignment=Qt.AlignmentFlag.AlignCenter)
            label.setFont(font)
            label.setStyleSheet("border: none;")
            quote_layout.addWidget(label)

        label = QLabel(self.QUOTE_AUTHOR, alignment=Qt.AlignmentFlag.AlignRight)
        font.setPointSize(default_font_size + 1)
        label.setFont(font)
        label.setStyleSheet("border: none;")
        quote_layout.addWidget(label)

        quote_container = QFrame()
        quote_container.setFrameShape(QFrame.Shape.Box)
        quote_container.setFrameShadow(QFrame.Shadow.Sunken)
        quote_container.setStyleSheet("QFrame {border: 1px dashed gray;padding: 1px;}")

        quote_container.setLayout(quote_layout)

        right_layout.addWidget(quote_container)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.button(QDialogButtonBox.StandardButton.Ok).setDefault(True)
        button_box.accepted.connect(self.close)

        right_layout.addWidget(button_box)

        right_container = QWidget()
        right_container.setLayout(right_layout)

        left_container = QWidget()
        left_container.setLayout(left_layout)

        layout = QHBoxLayout()
        layout.addWidget(left_container)
        layout.addWidget(right_container)

        self.setLayout(layout)

        self.adjustSize()
        self.setFixedSize(self.size())


class SettingsDialog(QDialog):
    """Maintains a simple 'settings' data structure across multiple platforms."""

    SETTINGS = ["Name", "Employee Number", "Building/Department", "In-District Mileage Account",
                "Out-of-District Mileage Account", "Administrator's Name"]

    def __init__(self, settings):
        super().__init__()
        self.setWindowTitle("User Info")
        layout = QFormLayout()
        self.edits = {}
        for setting in self.SETTINGS:
            edit = QLineEdit(settings[setting])
            edit.setFixedWidth(200)
            self.edits[setting] = edit
            layout.addRow(QLabel(f"{setting}:"), edit)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel)
        button_box.button(QDialogButtonBox.StandardButton.Ok).setDefault(True)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        self.setLayout(layout)
        self.adjustSize()
        self.setFixedSize(self.size())

    def get_settings(self):
        """Return a dict containing the current settings from this QDialog window."""
        settings = {}
        for setting in self.SETTINGS:
            settings[setting] = self.edits[setting].text()
        return settings

    @staticmethod
    def _get_settings_path():
        """Returns a reasonable place to put a settings file on different platforms."""
        system_name = platform.system()
        if system_name == 'Linux':
            return os.path.expanduser(f'~/.{APP_NAME}/settings.json')
        if system_name == 'Darwin':
            return os.path.expanduser(f'~/Library/Application Support/{APP_NAME}/settings.json')
        return os.path.join(os.getenv('LOCALAPPDATA'), APP_NAME, 'preferences.json')

    @staticmethod
    def load_settings():
        """
        Reads a dict containing settings from the settings file. Initializes an empty dict
        if a file isn't found.
        """
        pref_path = SettingsDialog._get_settings_path()
        if os.path.exists(pref_path):
            with open(pref_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            return {s:'' for s in SettingsDialog.SETTINGS}

    @staticmethod
    def save_settings(settings):
        """Saves the specified settings dict to file."""
        pref_path = SettingsDialog._get_settings_path()
        pref_dir = os.path.dirname(pref_path)
        if not os.path.exists(pref_dir):
            os.makedirs(pref_dir)
        with open(pref_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f)


class MainWindow(QMainWindow):
    """Presents a table of QDateEdit and QLineEdit widgets."""

    ROW_COUNT = 11
    COL_COUNT = 6
    MAX_ROWS = 45

    DATE_COL_INDEX = 0
    FROM_COL_INDEX = 1
    TO_COL_INDEX = 2
    PURPOSE_COL_INDEX = 3
    PARKING_COL_INDEX = 4
    MILES_COL_INDEX = 5

    PURPOSE = ["Mentoring meeting", "Meeting", "Return to office", "Drop off materials",
        "Classroom visit", "Office"]

    COLS = [("Date", 100), ("From Location", 120), ("To Location", 120), ("Purpose", 250),
        ("Parking", 60), ("Miles", 60)]

    ROW_COL_WIDTH = 25

    EMPTY_ROW_COLOR = 'gray'
    DATE_STR_FORMAT = "MM/dd/yyyy"

    def __init__(self, data, settings, initial_file=None):
        super().__init__()
        if len(data) == 2:
            self.data_date = None
            self.addresses, self.distances = data
        elif len(data) == 3:
            self.data_date, self.addresses, self.distances = data

        self.settings = settings
        self.doc_path = None

        self.setWindowTitle(self._get_window_title())
        self.setAcceptDrops(True)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        self.table_widget = QTableWidget(central_widget)
        self.table_widget.setColumnCount(self.COL_COUNT)
        self.table_widget.setHorizontalHeaderLabels([c[0] for c in self.COLS])

        for i, c in enumerate(self.COLS):
            self.table_widget.setColumnWidth(i, c[1])

        self.table_widget.verticalHeader().setFixedWidth(self.ROW_COL_WIDTH)

        self._create_completers_and_validators()

        for row in range(self.ROW_COUNT):
            self._add_table_row(row)

        self.table_widget.itemSelectionChanged.connect(self._cell_selection_changed)

        self.row_button = QPushButton("Add row")
        self.row_button.setShortcut("Ctrl+A")
        self.row_button.clicked.connect(self._grow_table)

        pdf_button = QPushButton("Create PDF")
        pdf_button.setShortcut("Ctrl+P")
        pdf_button.clicked.connect(self._create_pdf)

        layout = QVBoxLayout(central_widget)
        layout.addWidget(self.table_widget)
        bottom_layout = QHBoxLayout()
        bottom_layout.addWidget(self.row_button)
        bottom_layout.addStretch(1)  # Add stretchable space to push button to the right
        bottom_layout.addWidget(pdf_button)
        layout.addLayout(bottom_layout)
        central_widget.setLayout(layout)

        icon = QIcon(os.path.join(BASE_DIR, ARTEFACTS_DIR, ICON_FILE))
        self.setWindowIcon(icon)
        self._add_menubar()
        self._set_min_and_max_window_size()

        self.last_save = self._read_table()
        if initial_file:
            self.open_file(initial_file)

    def _set_min_and_max_window_size(self):
        style = self.table_widget.style()
        option = QStyleOption()
        option.initFrom(self.table_widget)

        extra_h_space = self.ROW_COL_WIDTH
        v = style.pixelMetric(QStyle.PixelMetric.PM_DefaultFrameWidth, option, self.table_widget)
        extra_h_space += 2 * v

        v = style.pixelMetric(QStyle.PixelMetric.PM_LayoutLeftMargin, option, self.table_widget)
        extra_h_space += 2 * v

        # Add extra space so the vertical scrollbar doesn't cause a horizontal scrollbar on Big Sur.
        scrollbar_width = style.pixelMetric(QStyle.PixelMetric.PM_ScrollBarExtent)
        extra_h_space += scrollbar_width

        preferred_width = sum(x[1] for x in self.COLS) + extra_h_space
        self.setMinimumSize(preferred_width,
                self.table_widget.rowHeight(0) * (self.ROW_COUNT + 3) + 3)
        self.setMaximumWidth(preferred_width)

    def _create_completers_and_validators(self):
        school_names = self.addresses.keys()
        self.school_completer = QCompleter(school_names)
        self.school_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.school_completer.activated.connect(self._enter_pressed)

        self.purpose_completer = QCompleter(self.PURPOSE)
        self.purpose_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.purpose_completer.activated.connect(self._enter_pressed)

        self.money_validator = QDoubleValidator()
        self.money_validator.setBottom(0.0)
        self.money_validator.setDecimals(2)

        self.mileage_validator = QDoubleValidator()
        self.mileage_validator.setBottom(0.0)
        self.mileage_validator.setDecimals(1)

    def _add_menubar(self):
        new_action = QAction("&New File", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self._new_file)

        open_action = QAction("&Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)

        self.save_action = QAction("&Save", self)
        self.save_action.setShortcut("Ctrl+S")
        self.save_action.triggered.connect(self._save_file)

        save_as_action = QAction("&Save As", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self._save_as)

        exit_action = QAction("&Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)

        settings_action = QAction("User Info", self)
        settings_action.setShortcut("Ctrl+,")
        settings_action.triggered.connect(self._show_settings)

        about_action = QAction("About", self)
        about_action.triggered.connect(self._show_about)

        menubar = self.menuBar()
        file_menu = menubar.addMenu("&File")
        pref_menu = menubar.addMenu("Settings")
        help_menu = menubar.addMenu("Help")

        file_menu.addAction(new_action)
        file_menu.addAction(open_action)
        file_menu.addSeparator()
        file_menu.addAction(self.save_action)
        file_menu.addAction(save_as_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)
        pref_menu.addAction(settings_action)
        help_menu.addAction(about_action)

    def _add_table_row(self, row):
        self.table_widget.insertRow(row)
        for col in range(self.COL_COUNT):
            if col == self.DATE_COL_INDEX:
                date_edit = QDateEdit()
                date_edit.setCalendarPopup(True)
                date_edit.setDate(QDate.currentDate())
                date_edit.setStyleSheet(f"color: {self.EMPTY_ROW_COLOR}")
                self.table_widget.setCellWidget(row, col, date_edit)
                # NOTE: Don't connect this until after setDate is called.
                date_edit.dateChanged.connect(self._on_date_changed)
            else:
                line_edit = CustomLineEdit(row)
                if col in (self.FROM_COL_INDEX, self.TO_COL_INDEX):
                    line_edit.setCompleter(self.school_completer)
                elif col == self.PURPOSE_COL_INDEX:
                    line_edit.setCompleter(self.purpose_completer)
                else:
                    line_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    if col == self.PARKING_COL_INDEX:
                        line_edit.setValidator(self.money_validator)
                    else:
                        line_edit.setValidator(self.mileage_validator)
                    line_edit.returnPressed.connect(self._return_pressed)

                self.table_widget.setCellWidget(row, col, line_edit)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """
        This, along with dropEvent(), allow documents to be opened by dropping them onto
        the main window.
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        """
        This, along with dragEnterEvent(), allow documents to be opened by dropping them onto
        the main window.
        """
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            self.open_file(file_path)

    def _get_window_title(self, doc_needs_saving=False):
        if self.doc_path:
            doc = os.path.basename(self.doc_path)
            if doc_needs_saving:
                return f"{APP_NAME} ({doc}*)"
            return f"{APP_NAME} ({doc})"
        if doc_needs_saving:
            return f"{APP_NAME}(*)"
        return f"{APP_NAME}"

    def closeEvent(self, event):
        """Override the default close function to prompt the user in case of unsaved changes."""
        data = self._read_table()
        if data != self.last_save:
            reply = QMessageBox.question(self,
                "Confirm Your Action",
                "Do you want to save your data before exiting?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                if not self._save_file(data):
                    event.ignore()
                    return
        event.accept()

    def _on_date_changed(self, date):
        """
        Most of the time the user updates the table in chronological order. To make this easier
        we propogate date changes down from the current line so (hopefully) the desired date will
        be closer when the user needs to select the next one. Lines with existing data are skipped.
        """
        coords = self._get_row_and_col()
        if coords:
            row, _ = coords
            for i in range(row + 1, self.table_widget.rowCount()):
                if self._row_is_empty(i):
                    self.table_widget.cellWidget(i, self.DATE_COL_INDEX).setDate(date)

    def _create_pdf(self):
        table = self._read_table(True)
        data = {"USER_INFO": self.settings, "TABLE": table}
        if not table:
            QMessageBox.warning(self,
                "Error",
                "The table doesn't contain any useful data.",
                QMessageBox.StandardButton.Ok)
            return
        err_col_str = None
        err_row = None
        err_col = None
        for i in range(0, len(table)):
            if not table[i][self.FROM_COL_INDEX]:
                err_col_str = 'From'
                err_row = i
                err_col = self.FROM_COL_INDEX
                break
            if not table[i][self.TO_COL_INDEX]:
                err_col_str = 'To'
                err_row = i
                err_col = self.TO_COL_INDEX
                break
            if not table[i][self.MILES_COL_INDEX]:
                err_col_str = 'Miles'
                err_row = i
                err_col = self.MILES_COL_INDEX
                break
        if err_col_str:
            QMessageBox.warning(self,
                "Warning",
                f"At least one line doesn't have a value for the '{err_col_str}' column.",
                QMessageBox.StandardButton.Ok,
                QMessageBox.StandardButton.Ok)
            self.table_widget.setCurrentCell(err_row, err_col)
            return

        file_dialog = QFileDialog(self)
        file_dialog.setDirectory(self._get_default_file_dialog_path())
        file_path, _ = file_dialog.getSaveFileName(self,
            "Save File",
            "",
            "PDF Files (*.pdf);;All Files (*)")
        if file_path:
            _, ext = os.path.splitext(file_path)
            if not ext:
                file_path = file_path + ".pdf"

            try:
                fill_form(os.path.join(BASE_DIR, ARTEFACTS_DIR, FORM_FILE), file_path, data)
            except:
                QMessageBox.critical(self,
                    "Error",
                    "The PDF could not be saved.",
                    QMessageBox.StandardButton.Ok)
            else:
                QMessageBox.information(self,
                    "PDF Created", "Remember to review the PDF and sign it before submitting!",
                    QMessageBox.StandardButton.Ok)

    def _read_table(self, strip_empty_rows=False):
        # Read the table starting from the bottom and skip empty rows until content is encountered.
        # Then only skip empty rows if strip_empty_rows.
        data = []
        found_non_empty_row = False
        for i in range(self.table_widget.rowCount() - 1, -1, -1):
            row = []
            for j in range(1, self.COL_COUNT):
                row.append(self.table_widget.cellWidget(i, j).text())
            # Don't preserve dates on lines that are otherwise empty.
            if row != ['', '', '', '', '']:
                found_non_empty_row = True
                d = self.table_widget.cellWidget(i, self.DATE_COL_INDEX).date()
                row.insert(0, d.toString(self.DATE_STR_FORMAT))
                data.append(row)
            elif not strip_empty_rows and found_non_empty_row:
                row.insert(0, None)
                data.append(row)
        data.reverse()
        return data

    def _write_table(self, data):
        today = QDate.currentDate()
        for i in range(self.table_widget.rowCount(), len(data)):
            self._add_table_row(i)
        for i, row in enumerate(data):
            for j in range(self.FROM_COL_INDEX, self.COL_COUNT):
                self.table_widget.cellWidget(i, j).setText(row[j])
            date_widget = self.table_widget.cellWidget(i, self.DATE_COL_INDEX)
            if row[self.DATE_COL_INDEX]:
                date = QDate.fromString(row[self.DATE_COL_INDEX], self.DATE_STR_FORMAT)
                date_widget.setDate(date)
            else:
                date_widget.setDate(today)

    def _get_default_file_dialog_path(self):
        """Get the path to the user's Desktop unless a file has been recently opened or saved"""
        if self.doc_path:
            return os.path.dirname(self.doc_path)
        return os.path.expanduser('~/Desktop')

    def _update_save_item_and_title(self, read_table=False):
        if not read_table:
            self.save_action.setEnabled(False)
            self.setWindowTitle(self._get_window_title(False))
            return
        data = self._read_table()
        if data != self.last_save:
            if not self.save_action.isEnabled():
                self.save_action.setEnabled(True)
                self.setWindowTitle(self._get_window_title(True))
        else:
            if self.save_action.isEnabled():
                self.save_action.setEnabled(False)
                self.setWindowTitle(self._get_window_title(False))

    def _save(self, file_path, data=None):
        if not data:
            data = self._read_table()
        try:
            with open(file_path, 'w', encoding='utf-8') as file:
                json.dump(data, file, indent=4)
                self.last_save = data
                self.doc_path = file_path
                self._update_save_item_and_title()
                return True
        except:
            QMessageBox.critical(self,
                "Error",
                "The file could not be saved.",
                QMessageBox.StandardButton.Ok)
        return False

    def _save_as(self, data=None):
        file_dialog = QFileDialog(self)
        file_dialog.setDirectory(self._get_default_file_dialog_path())
        file_path, _ = file_dialog.getSaveFileName(self,
            "Save File", "", f"{APP_NAME} Files (*.{APP_EXT});;All Files (*)")
        if file_path:
            _, ext = os.path.splitext(file_path)
            if not ext:
                file_path = file_path + f".{APP_EXT}"
            return self._save(file_path, data)
        return False

    def _save_file(self, data=None):
        if self.doc_path:
            return self._save(self.doc_path, data)
        return self._save_as(data)

    def _clear_table(self):
        self.table_widget.setCurrentCell(0, 0)
        self.table_widget.setRowCount(self.ROW_COUNT)
        today = QDate.currentDate()
        for i in range(0, self.table_widget.rowCount()):
            self.table_widget.cellWidget(i, self.DATE_COL_INDEX).setDate(today)
            for j in range(1, self.COL_COUNT):
                self.table_widget.cellWidget(i, j).clear()
        self.table_widget.setCurrentCell(0, 0)

    def open_file(self, file_path=None):
        """Opens the specified file path"""
        if not file_path:
            data = self._read_table()
            if data != self.last_save:
                reply = QMessageBox.question( self,
                    "Confirm Your Action",
                    "Do you want to save your data before continuing?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes)
                if reply == QMessageBox.StandardButton.Yes:
                    if not self._save_file(data):
                        return
            file_dialog = QFileDialog(self)
            file_dialog.setDirectory(self._get_default_file_dialog_path())
            file_path, _ = file_dialog.getOpenFileName(self,
                "Open File",
                "",
                f"{APP_NAME} Files (*.{APP_EXT});;All Files (*)")

        if file_path:
            try:
                if os.path.exists(file_path):
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        self._clear_table()
                        self._write_table(data)
                        self.last_save = self._read_table()
                        self.doc_path = file_path
                        self._update_save_item_and_title()
            except:
                QMessageBox.critical(self,
                    "Error",
                    "The file could not be opened.",
                    QMessageBox.StandardButton.Ok)
            self._update_table(False)

    def _new_file(self):
        """Could have also been called 'Clear Table'."""
        data = self._read_table()
        if data != self.last_save:
            reply = QMessageBox.question(self,
                "Confirm Your Action",
                "Do you want to save your data before starting a new file?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                if not self._save_file(data):
                    return
        self._clear_table()
        self.last_save = self._read_table()
        self.doc_path = None
        self._update_save_item_and_title()

    def _grow_table(self):
        count = self.table_widget.rowCount()
        if count < self.MAX_ROWS:
            self._add_table_row(count)
            self._update_table(False)
        if count == (self.MAX_ROWS - 1):
            self.row_button.setEnabled(False)

    def _show_settings(self):
        dialog = SettingsDialog(self.settings)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            settings = dialog.get_settings()
            if settings != self.settings:
                SettingsDialog.save_settings(settings)
                self.settings = settings

    def _show_about(self):
        about_dialog = AboutDialog(self.data_date)
        about_dialog.exec()

    def _get_row_and_col(self):
        selected_ranges = self.table_widget.selectedRanges()
        if selected_ranges:
            selected_range = selected_ranges[0]
            return (selected_range.topRow(), selected_range.leftColumn())
        return None

    def _select_next_cell(self):
        coords = self._get_row_and_col()
        if coords:
            row, col = coords
            while True:
                if col == self.COL_COUNT - 1:
                    row += 1
                    col = 0
                else:
                    col += 1
                cell = self.table_widget.cellWidget(row, col)
                if cell:
                    if not cell.isReadOnly():
                        break
                else:
                    row = 0
                    col = 0
                    break
            self.table_widget.setCurrentCell(row, col)

    def _set_custom_font(self, row, col, color=None, bold=False):
        item = self.table_widget.cellWidget(row, col)
        sheet = ""
        if color:
            sheet += f"color: {color}"
        if bold:
            sheet += "font-weight: bold;"
        item.setStyleSheet(sheet)

    def _clear_custom_font(self, row, col):
        self._set_custom_font(row, col)

    def _cell_selection_changed(self):
        coords = self._get_row_and_col()
        if coords:
            row, col = coords
            self._update_table()
            self._clear_custom_font(row, self.DATE_COL_INDEX)
            widget = self.table_widget.cellWidget(row, col)
            if widget.isReadOnly():
                self._select_next_cell()

    def _enter_pressed(self, _):
        self._select_next_cell()

    def _return_pressed(self):
        self._select_next_cell()

    def _row_is_empty(self, row):
        data = ''.join(self.table_widget.cellWidget(row, j).text()
                                for j in range(self.FROM_COL_INDEX, self.COL_COUNT))
        return data == ''

    def _update_table(self, update_save_and_title=True):
        """Iterates through the entire table. For each line:
         -If the line is blank set the 'Date' text to gray, else default color
         -If the From and To locations are valid then set the 'Miles' value (with default color)
         -If there is an error with the From or To locations then highlite 'Miles' to show error
        """
        for i in range(0, self.table_widget.rowCount()):
            if self._row_is_empty(i):
                self._set_custom_font(i, self.DATE_COL_INDEX, self.EMPTY_ROW_COLOR)
            else:
                self._clear_custom_font(i, self.DATE_COL_INDEX)
                origin = self.table_widget.cellWidget(i, self.FROM_COL_INDEX).text()
                dest = self.table_widget.cellWidget(i, self.TO_COL_INDEX).text()
                miles = self.table_widget.cellWidget(i, self.MILES_COL_INDEX)
                if origin and dest:
                    if origin.upper() == dest.upper():
                        miles.setText('0')
                        self._set_custom_font(i, self.MILES_COL_INDEX, None, True)
                        miles.setReadOnly(True)
                    else:
                        try:
                            dist = self.distances[origin][dest]
                            miles.setText(str(dist))
                            self._set_custom_font(i, self.MILES_COL_INDEX, None, True)
                            miles.setReadOnly(True)
                        except KeyError:
                            if miles.isReadOnly():
                                miles.clear()
                            self._clear_custom_font(i, self.MILES_COL_INDEX)
                            miles.setReadOnly(False)
                else:
                    miles.setReadOnly(False)
                    self._clear_custom_font(i, self.MILES_COL_INDEX)
        if update_save_and_title:
            self._update_save_item_and_title(True)


class SmileApp(QApplication):
    """Subclass to allow event interception."""

    def __init__(self, argv, data, settings, file_path):
        super().__init__(argv)
        self.main_window = MainWindow(data, settings, file_path)
        self.main_window.show()

    def event(self, event: QEvent):
        """React to FileOpen events to enable double-clicking on documents."""
        if event.type() == QEvent.Type.FileOpen:
            file_path = event.file()
            self.main_window.open_file(file_path)
            return True
        return super().event(event)


def main():
    """Loads settings and the mileage data structure from the disk and shows the GUI."""
    data = None
    settings = SettingsDialog.load_settings()
    with open(os.path.join(BASE_DIR, ARTEFACTS_DIR, DATA_FILE), 'rb') as file:
        data = pickle.load(file)

    file_path = None
    if len(sys.argv) > 1:
        file_path = sys.argv[1]

    app = SmileApp(sys.argv, data, settings, file_path)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
