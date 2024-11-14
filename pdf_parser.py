"""
Parses 'official' driving distances between 4j school buildings from a PDF.

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


__version__ = "1.0.0"
__date__ = "Nov '24"


def _parse_float(value):
    try:
        return float(value)
    except ValueError:
        return 0

def parse_distance_table(path):
    pdf = fitz.open(path)
    page = pdf.load_page(0)
    table = page.find_tables().tables[0].extract()
    schools = [table[i][0] for i in range(1, len(table))]
    data = {}
    for i in range(1, len(table)):
        for j in range(i + 1, len(table)):
            o = f"{schools[i-1]}"
            d = f"{schools[j-1]}"
            val = _parse_float(table[i][j])
            if not data.get(o, None):
                data[o] = {}
            data[o][d] = val
    # The last row/col of the table doesn't get processed
    # but add a key for the school regardless.
    data[schools[-1]] = {}
    pdf.close()
    return data
