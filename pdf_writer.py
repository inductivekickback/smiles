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
NUM_ROWS_PAGE_2 = 34
MAX_ROWS = NUM_ROWS_PAGE_1 + NUM_ROWS_PAGE_2

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
    except ValueError:
        pass
    widget.field_value = value
    widget.update()


def _parse_float(value):
    try:
        return float(value)
    except ValueError:
        return 0


def _write_pdf(form_path, additional_form_path, save_path, form_data):
    pdf = fitz.open(form_path)

    for w in pdf.load_page(0).widgets():
        if w.field_type_string == "Text":
            if w.field_name in form_data:
                _update_widget(w, form_data[w.field_name])

    if additional_form_path:
        # Widgets are not copied when using the insert_pdf method
        # so they will be copied manually.
        add_pdf = fitz.open(additional_form_path)
        pdf.insert_pdf(add_pdf)
        new_page = pdf.load_page(1)

        for w in add_pdf.load_page(0).widgets():
            if w.field_type_string == "Text":
                new_page.add_widget(w)
                if w.field_name in form_data:
                    _update_widget(w, form_data[w.field_name])
        add_pdf.close()

    pdf.save(save_path)
    pdf.close()


def fill_form(form_path, additional_form_path, save_path, data):
    user_info = data['USER_INFO']
    table = data['TABLE']

    page_parking_totals = [0]
    page_miles_totals = [0]

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

    if len(table) > NUM_ROWS_PAGE_1:
        page_parking_totals.append(0)
        page_miles_totals.append(0)
        row = 0
        while table_row < len(table):
            for j in range(0, PARKING_COL_INDEX):
                field_name = FIELD_NAME_FMTS[j].format(2, row)
                values[field_name] = table[table_row][j]

            field_name = FIELD_NAME_FMTS[PARKING_COL_INDEX].format(2, row)
            value = _parse_float(table[table_row][PARKING_COL_INDEX])
            values[field_name] = value
            page_parking_totals[1] += value

            field_name = FIELD_NAME_FMTS[MILES_COL_INDEX].format(2, row)
            value = _parse_float(table[table_row][MILES_COL_INDEX])
            values[field_name] = value
            page_miles_totals[1] += value

            row += 1
            table_row += 1

    for i in range(0, len(page_parking_totals)):
        values[f"txtP{i+1}TotParking"] = page_parking_totals[i]
        values[f"txtP{i+1}TotMiles"] = page_miles_totals[i]

    values["txtGrTotParking"] = sum(page_parking_totals)
    values["txtGrTotMiles"] = sum(page_miles_totals)

    if len(table) > NUM_ROWS_PAGE_1:
        _write_pdf(form_path, additional_form_path, save_path, values)
    else:
        _write_pdf(form_path, None, save_path, values)
