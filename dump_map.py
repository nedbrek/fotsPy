#!/usr/bin/env python3

import json
import openpyxl
import sys

### excel utils
def resolveFormula(sheet, cell):
    val = cell.value
    if type(val) is int:
        return val

    if val[0] == '=':
        plus_idx = val.find("+")
        idx = val[1:plus_idx]
        v2 = sheet[idx].value
        const = int(val[plus_idx+1:])
        v2 = v2 + const
        cell.value = v2

    return v2

### classes
class System:
    def __init__(self):
        self.xoff = -1
        self.yoff = -1
        self.key = "<error>"
        self.star_type = "<error>"

class FotsData:
    def __init__(self):
        self.stars = list()

    def buildStarmap(self, map_sheet):
        self.first_x, last_x = findX(map_sheet)
        self.first_y, last_y = findY(map_sheet)

        # walk the map
        for col in range(3, (last_x - self.first_x + 1) * 10):
            for row in range(4, (last_y - self.first_y + 1) * 10):
                cell = map_sheet.cell(row, col)
                val = cell.value
                if val == None:
                    continue

                s = System()

                s.xoff = self.first_x * 10 + col - 3
                s.yoff = self.first_y * 10 + row - 4

                key = makeKey(s.xoff, s.yoff)
                s.key = key

                comment = cell.comment
                if type(comment) is not openpyxl.comments.comments.Comment:
                    s.star_type = val
                    self.stars.append(s)
                else:
                    comments = comment.text.split("\n")
                    com_cols = comments[0].split()
                    if key != com_cols[0]:
                        print("Ned error", key, com_cols[0], file=sys.stderr)
                    s.star_type = com_cols[1]

# find x dimensions
def findX(map_sheet):
    row = 2
    col = 7
    first_x = map_sheet.cell(row, col).value
    while 1:
        cell = map_sheet.cell(row, col)
        val = cell.value
        if val == None:
            break

        last_x = resolveFormula(map_sheet, cell)

        col = col + 10
    #end while
    return (first_x, last_x)

# find grid y
def findY(map_sheet):
    row = 8
    col = 1
    first_y = map_sheet.cell(row, col).value
    while 1:
        cell = map_sheet.cell(row, col)
        val = cell.value
        if val == None:
            break

        last_y = resolveFormula(map_sheet, cell)

        row = row + 10
    #end while
    return (first_y, last_y)

def makeKey(grid_x, grid_y):
    dec_x = int(grid_x / 10)
    sec_x = (grid_x % 10) + 1

    dec_y = int(grid_y / 10)
    sec_y = grid_y % 10
    key = "{}x{}/{:02d}".format(dec_x, dec_y, sec_x + sec_y * 10)
    return key

def dumpStarmap(map_sheet):
    data = FotsData()
    data.buildStarmap(map_sheet)
    print(json.dumps(data, default=vars))

def dumpMap(wb_obj):
    map_sheet = wb_obj["Map"]
    dumpStarmap(map_sheet)

### main
if __name__ == '__main__':
    import argparse
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("turnfile", nargs=1, help = "Excel turn file")
    args = arg_parser.parse_args()

    wb_obj = openpyxl.load_workbook(args.turnfile[0])
    dumpMap(wb_obj)

