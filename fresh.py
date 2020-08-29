# Fresh to table
from docx import Document
import csv
import argparse

# Setup command line arguments
parser = argparse.ArgumentParser()
parser.add_argument("inputFile", help="File that contains the list")

parser.add_argument(
    "-w", "--width", help="Number of columns in the output table, 0 to use the width of the input file", type=int)
parser.add_argument("-d", "--delimiter",
                    help="Character that separates list items", default=",")
parser.add_argument("outputFile", help="Output file")
args = parser.parse_args()

# Parse
document = Document()
table = document.add_table(rows=1, cols=args.width)
with open(args.inputFile) as input_file:
    reader = csv.reader(input_file, delimiter=args.delimiter)
    col_counter = 0
    row_counter = 0
    for row in reader:  # TODO add ability to collapse many rows into one
        for item in row:
            if col_counter == args.width:
                row_cells = table.add_row()
                col_counter = 0
                row_counter += 1
            curr_cells = table.rows[row_counter].cells
            curr_cells[col_counter].text = item
            col_counter += 1

# Write
document.save(args.outputFile)
