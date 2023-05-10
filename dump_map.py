#!/usr/bin/env python3

import openpyxl

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
    sec_x = grid_x % 10

    dec_y = int(grid_y / 10)
    sec_y = grid_y % 10
    key = "{}x{}/{:02d}".format(dec_x, dec_y, sec_x + sec_y * 10)
    return key

def dumpStarmap(map_sheet):
    first_x, last_x = findX(map_sheet)
    first_y, last_y = findY(map_sheet)

    # walk the map
    for col in range(3, (last_x - first_x + 1) * 10):
        for row in range(4, (last_y - first_y + 1) * 10):
            cell = map_sheet.cell(row, col)
            val = cell.value
            if val == None:
                continue

            off_x = first_x * 10 + col - 3
            off_y = first_y * 10 + row - 4

            print("{} {}".format(makeKey(off_x, off_y), val))


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

