#!/usr/bin/env python3

import json
import tkinter as tk
from tkinter import ttk

zoom_levels = [
   10,
   20,
   40,
   60,
   90,
   150
]

data = dict()
side = 10
canvas = None
first_x = 0
first_y = 0

def drawData():
    global side
    global canvas

    canvas.delete('all')

    for s in data["stars"]:
        x = (s["xoff"] - first_x) * side
        y = (s["yoff"] - first_y) * side
        t = s["star_type"]
        fill_color = 'black'
        if s["grid"]:
            fill_color = 'dark blue'
        canvas.create_rectangle(x, y, x + side, y + side, fill=fill_color)

        if side >= 40:
            star_type = s["star_type"]
            canvas.create_text(x + side / 2, y + side / 2, text=star_type, fill='white')

    canvas.configure(scrollregion = canvas.bbox('all'))

def zoomOut(event):
    global side
    global zoom_levels

    idx = zoom_levels.index(side)
    if idx == 0:
        return
    side = zoom_levels[idx - 1]
    drawData()

def zoomIn(event):
    global side
    global zoom_levels

    idx = zoom_levels.index(side)
    if idx + 1 == len(zoom_levels):
        return
    side = zoom_levels[idx + 1]
    drawData()

### main
if __name__ == '__main__':
    import argparse
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("turnfile", nargs=1, help = "JSON turn file")
    args = arg_parser.parse_args()

    with open(args.turnfile[0], 'r') as turn_file:
        data = json.loads(turn_file.read())

    # length of the side of a square
    print(data.keys())

    root = tk.Tk()

    hsb = ttk.Scrollbar(root, orient=tk.HORIZONTAL)
    vsb = ttk.Scrollbar(root, orient=tk.VERTICAL)

    canvas = tk.Canvas(root, yscrollcommand = vsb.set, xscrollcommand = hsb.set)

    first_x = data["first_x"] * 10
    first_y = data["first_y"] * 10

    drawData()

    hsb['command'] = canvas.xview
    vsb['command'] = canvas.yview

    hsb.pack(fill='x', side='bottom')
    vsb.pack(fill='y', side='right')
    canvas.pack(expand=True, fill='both', side='right')

    # bind zoom
    root.bind('<minus>',       zoomOut)
    root.bind('<KP_Subtract>', zoomOut)
    root.bind('<plus>',        zoomIn)
    root.bind('<KP_Add>',      zoomIn)

    root.mainloop()

