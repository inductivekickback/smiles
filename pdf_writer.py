"""
Fills text field values into the 4j Mileage Reimbursement form.

Disclaimer of Warranty

This software is provided "as is," without any warranties of any kind, either express or implied,
including but not limited to the implied warranties of merchantability, fitness for a particular
purpose, or non-infringement. The entire risk arising out of the use or performance of the software
remains with you. In no event shall the software provider be liable for any direct, indirect,
incidental, special, consequential, or punitive damages whatsoever arising out of or in connection
with the use or inability to use the software.
"""
from datetime import datetime

import fitz


__version__ = "1.1.0"
__date__ = "Nov '24"

NUM_COLS_PER_LINE = 6
NUM_ROWS_PAGE_1 = 22
NUM_ROWS_PAGE_X = 34

PARKING_COL_INDEX = 4
MILES_COL_INDEX = 5

USER_INFO = [("Name", "txtEmpName"),
    ("Employee Number", "txtEmpNumber"),
    ("Building/Department", "txtSchool"),
    ("Account #", "txtInDistAcct"),
    ("Supervisor's Name", "Text1")]

FIELD_NAME_FMTS = ["txtP{0}Date.{1}", "txtP{0}FromLoc.{1}",
                        "txtP{0}ToLoc.{1}", "txtP{0}Purpose.{1}",
                        "txtP{0}AmtParking.{1}", "txtP{0}Miles.{1}"]

INPUT_STR_FORMAT = "%m/%d/%Y" # Equivalent to QDate's toString("MM/dd/yyyy")
OUTPUT_STR_FORMAT = "%m/%d/%y"


def _update_widget(widget, value):
    try:
        if 'Parking' in widget.field_name:
            value = f"{value:0.2f}"
        elif 'Miles' in widget.field_name:
            value = f"{value:0.1f}"
        elif 'Date' in widget.field_name:
            value = datetime.strptime(value, INPUT_STR_FORMAT).strftime(OUTPUT_STR_FORMAT)
    except ValueError as ex:
        pass
    widget.field_value = value
    widget.update()

def _parse_float(value):
    try:
        return float(value)
    except ValueError:
        return 0

def fill_form(path, save_path, data):
    user_info = data['USER_INFO']
    table = data['TABLE']

    pdf = fitz.open(path)
    page_parking_totals = [0 for x in range(0, len(pdf))]
    page_miles_totals = [0 for x in range(0, len(pdf))]

    values = {}

    for setting, field_name in USER_INFO:
        values[field_name] = user_info[setting]

    page_len = len(table) if len(table) < NUM_ROWS_PAGE_1 else NUM_ROWS_PAGE_1
    table_row = 0
    row = 0
    while table_row < page_len:
        for j in range(0, PARKING_COL_INDEX):
            field_name = FIELD_NAME_FMTS[j].format(1, row)
            values[field_name] = table[table_row][j]

        field_name = FIELD_NAME_FMTS[PARKING_COL_INDEX].format(1, row)
        value = _parse_float(table[table_row][PARKING_COL_INDEX])
        values[field_name] = value
        page_parking_totals[0] += value

        field_name = FIELD_NAME_FMTS[MILES_COL_INDEX].format(1, row)
        value = _parse_float(table[table_row][MILES_COL_INDEX])
        values[field_name] = value
        page_miles_totals[0] += value

        row += 1
        table_row += 1

    for page_num in range(1, len(pdf)):
        if table_row >= len(table):
            break
        row = 0
        for i in range(0, NUM_ROWS_PAGE_X):
            if table_row == len(table):
                break
            for j in range(0, PARKING_COL_INDEX):
                field_name = FIELD_NAME_FMTS[j].format(page_num+1, row)
                values[field_name] = table[table_row][j]

            field_name = FIELD_NAME_FMTS[PARKING_COL_INDEX].format(page_num+1, row)
            value = _parse_float(table[table_row][PARKING_COL_INDEX])
            values[field_name] = value
            page_parking_totals[page_num] += value
    
            field_name = FIELD_NAME_FMTS[MILES_COL_INDEX].format(page_num+1, row)
            value = _parse_float(table[table_row][MILES_COL_INDEX])
            values[field_name] = value
            page_miles_totals[page_num] += value

            table_row += 1
            row += 1

    for i in range(0, len(pdf)):
        values[f"txtP{i+1}TotParking"] = page_parking_totals[i]
        values[f"txtP{i+1}TotMiles"] = page_miles_totals[i]

    values["txtGrTotParking"] = sum(page_parking_totals)
    values["txtGrTotMiles"] = sum(page_miles_totals)

    for i in range(0, len(pdf)):
        page = pdf.load_page(i)
        for w in page.widgets():
            if w.field_type_string == "Text":
                if values.get(w.field_name):
                    _update_widget(w, values[w.field_name])

    pdf.save(save_path)
    pdf.close()
