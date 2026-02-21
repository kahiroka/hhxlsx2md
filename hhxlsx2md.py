import sys
import argparse
from openpyxl import load_workbook

MAX_ROW = 1000
MAX_COL = 1000

def parse_excel_with_hh(xlsx_path, geometry=[0,0,1,1], workaround=[]):
    wb = load_workbook(xlsx_path, data_only=True)
    for sheetname in wb.sheetnames:
        print(f"# {sheetname}")
        ws = wb[sheetname]
        data_col, data_row, header_width, header_height = geometry
        for row in range(data_row, MAX_ROW):
            term = True
            for hdr_col in range(data_col - header_width, data_col):
                hdr = ws.cell(row=row, column=hdr_col).value
                if hdr:
                    term = False
                    indent = hdr_col - (data_col - header_width) + 2
                    print("#"*indent + f" {hdr}")
            if term:
                break
                    
            for col in range(data_col, MAX_COL):
                value = ws.cell(row=row, column=col).value

                term = True
                for hdr_row in range(data_row - header_height, data_row):
                    hdr = ws.cell(row=hdr_row, column=col).value
                    if hdr:
                        term = False
                        indent = hdr_row - (data_row - header_height)
                        if hdr_row < data_row -1:
                            print(" "*indent*2 + "-" + f" {hdr}")
                        elif value or "dense" in workaround:
                            print(" "*indent*2 + "-" + f" {hdr}: ", end="")
                if term:
                    continue

                if value or "dense" in workaround:
                    print(f"{value}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="")
    parser.add_argument("filename", help="file name")
    parser.add_argument("--geometry", type=str, default="2x2+1+1", help="XxY[+W+H]: default data (X=1,Y=1), header (W=1,H=1)")
    parser.add_argument("--dense", action="store_true", help="display if empty")
    args = parser.parse_args()

    tmp = args.geometry.split("+")
    position = [int(x) for x in tmp[0].split("x")]
    layer = [int(tmp[1]), int(tmp[2])]
    geometry = position + layer
    workaround = []
    if args.dense:
        workaround.append("dense")
    if position[0] < layer[0] or position[1] < layer[1]:
        print(f"geometry mismatch: {geometry}", file=sys.stderr)
        sys.exit()

    parse_excel_with_hh(args.filename, geometry, workaround)