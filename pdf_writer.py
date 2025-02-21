"""
Fills text field values into the 4j 'Mileage Reimbursement Request' form.

The official PDF is very brittle and its checkboxes don't always render correctly in different
applications or on older macOS machines. To ensure compatibility, the original form's field_names
and widget geometries were extracted and this module now uses them to draw the forms.


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
from widget_data import FIELD_NAME_FMTS, PAGE_1_WIDGETS, PAGE_2_WIDGETS


__version__ = "1.2.0"
__date__ = "Feb '25"


NUM_ROWS_PAGE_1 = 22
NUM_ROWS_PAGE_2 = 34
MAX_ROWS = NUM_ROWS_PAGE_1 + NUM_ROWS_PAGE_2

PARKING_COL_INDEX = 4
MILES_COL_INDEX = 5
ROUND_TRIP_COL_INDEX = 6

DEFAULT_FONT_SIZE = 9.5
FONT_SIZE_DECREMENT = 0.5


USER_INFO = [("Name", "txtEmpName"),
    ("Employee Number", "txtEmpNumber"),
    ("Building/Department", "txtSchool"),
    ("Account Number", "txtInDistAcct")]

INPUT_STR_FORMAT = "%m/%d/%Y" # Equivalent to QDate's toString("MM/dd/yyyy")
OUTPUT_STR_FORMAT = "%m/%d/%y"


def _insert_text(page, widget_rect, value, alignment):
    """The bounding boxes of the widgets in the original PDF are not uniform."""
    font_size = DEFAULT_FONT_SIZE
    finished = False
    while not finished:
        unused_space = page.insert_textbox(widget_rect, value, fontsize=font_size, align=alignment)
        if unused_space < 0:
            font_size -= FONT_SIZE_DECREMENT
        else:
            finished = True


def _update_widget(page, field_name, widget_rect, value):
    alignment = fitz.TEXT_ALIGN_CENTER
    try:
        if 'Parking' in field_name:
            # Don't clutter the parking col with zeroes.
            if value:
                value = f"{value:0.2f}"
            else:
                value = ''
        elif 'Miles' in field_name:
            value = f"{value:0.1f}"
        elif 'Date' in field_name:
            value = datetime.strptime(value, INPUT_STR_FORMAT).strftime(OUTPUT_STR_FORMAT)
        elif 'Check' in field_name:
            if value == '1':
                value = 'X'
            else:
                value = ''
        else:
            alignment = fitz.TEXT_ALIGN_LEFT
            # Add a leading space because TEXT_ALIGN_LEFT is very aggressive.
            value = f" {value}"
    except ValueError:
        pass

    # Redraw the widget's outline (since we removed all of the actual widgets).
    page.draw_rect(widget_rect, color=(0, 0, 0), fill=(1, 1, 1), overlay=True)
    if value:
        _insert_text(page, widget_rect, value, alignment)


def _parse_float(value):
    try:
        return float(value)
    except ValueError:
        return 0


def _write_pdf(form_path, additional_form_path, save_path, form_data):
    pdf = fitz.open(form_path)
    page = pdf.load_page(0)
    page.clean_contents()
    for key, widget_rect in PAGE_1_WIDGETS.items():
        form_value = form_data.get(key, '')
        _update_widget(page, key, widget_rect, form_value)

    if additional_form_path:
        add_pdf = fitz.open(additional_form_path)
        pdf.insert_pdf(add_pdf)
        add_pdf.close()
        page = pdf.load_page(1)
        page.clean_contents()
        for key, widget_rect in PAGE_2_WIDGETS.items():
            form_value = form_data.get(key, '')
            _update_widget(page, key, widget_rect, form_value)

    pdf.save(save_path)
    pdf.close()


def fill_form(form_path, additional_form_path, save_path, data):
    """Compute totals while creating a dict for looking up values using the PDF's field_names"""
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
    data_has_rt_checkbox = len(table[table_row]) > ROUND_TRIP_COL_INDEX
    while table_row < page_len:
        for j in range(0, PARKING_COL_INDEX):
            field_name = FIELD_NAME_FMTS[j].format(1, row)
            values[field_name] = table[table_row][j]

        round_trip = False
        if data_has_rt_checkbox:
            field_name = FIELD_NAME_FMTS[ROUND_TRIP_COL_INDEX].format(table_row + 1)
            values[field_name] = table[table_row][ROUND_TRIP_COL_INDEX]
            round_trip = values[field_name] != '0'

        field_name = FIELD_NAME_FMTS[PARKING_COL_INDEX].format(1, row)
        value = _parse_float(table[table_row][PARKING_COL_INDEX])
        values[field_name] = value
        page_parking_totals[0] += value

        field_name = FIELD_NAME_FMTS[MILES_COL_INDEX].format(1, row)
        value = _parse_float(table[table_row][MILES_COL_INDEX])
        values[field_name] = value
        page_miles_totals[0] += value
        if round_trip:
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

            round_trip = False
            if data_has_rt_checkbox:
                field_name = FIELD_NAME_FMTS[ROUND_TRIP_COL_INDEX].format(table_row + 1)
                values[field_name] = table[table_row][ROUND_TRIP_COL_INDEX]
                round_trip = values[field_name] != '0'

            field_name = FIELD_NAME_FMTS[PARKING_COL_INDEX].format(2, row)
            value = _parse_float(table[table_row][PARKING_COL_INDEX])
            values[field_name] = value
            page_parking_totals[1] += value

            field_name = FIELD_NAME_FMTS[MILES_COL_INDEX].format(2, row)
            value = _parse_float(table[table_row][MILES_COL_INDEX])
            values[field_name] = value
            page_miles_totals[1] += value
            if round_trip:
                page_miles_totals[1] += value

            row += 1
            table_row += 1

    for i in range(0, len(page_parking_totals)):
        values[f"txtP{i+1}TotParking"] = page_parking_totals[i]
        values[f"txtP{i+1}TotMiles"] = page_miles_totals[i]

    values["txtP1&2TotParking"] = sum(page_parking_totals)
    values["txtP1&P2TotMiles"] = sum(page_miles_totals)

    if len(table) > NUM_ROWS_PAGE_1:
        _write_pdf(form_path, additional_form_path, save_path, values)
    else:
        _write_pdf(form_path, None, save_path, values)
