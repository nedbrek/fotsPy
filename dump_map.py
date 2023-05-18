#!/usr/bin/env python3

import fots_lib as fots
import json
import openpyxl
import sys

def dumpMap(wb_obj):
    summary_sheet = wb_obj["Summary"]
    row = 7
    col = 12
    current_turn = summary_sheet.cell(row, col)
    #print("Data for turn:", current_turn.value)

    map_sheet = wb_obj["Map"]
    data = fots.Fots()
    data.buildStarmap(map_sheet)
    print(json.dumps(data, default=vars))

def dumpHab(wb_obj):
    habs = fots.Habs()
    habs.readExcel(wb_obj)
    print(json.dumps(habs, default=vars))


### main
if __name__ == '__main__':
    import argparse
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("turnfile", nargs=1, help = "Excel turn file")
    arg_parser.add_argument("--hab", help = "Dump hab info", action="store_true")
    args = arg_parser.parse_args()

    wb_obj = openpyxl.load_workbook(args.turnfile[0])

    if args.hab:
        dumpHab(wb_obj)
        sys.exit(0)

    dumpMap(wb_obj)

