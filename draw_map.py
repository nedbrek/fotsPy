#!/usr/bin/env python3

from enum import Enum, auto
import json
import tkinter as tk
from tkinter import ttk

###
class EHabs(Enum):
    NO_HABS   = auto()  # nothing worth colonizing
    ALL_HABS  = auto()  # all habitable worlds colonized
    GOOD_HABS = auto()  # good colony site
    BAD_HABS  = auto()  # marginal colony site

###
zoom_levels = [
   10,
   20,
   40,
   60,
   90,
   150
]

data = dict()
hab_data = dict()
side = 10 # length of the side of a square
canvas = None
first_x = 0
first_y = 0

def drawData():
    global data
    global hab_data
    global side
    global canvas
    global first_x
    global first_y

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
            fill_color = 'grey'

            if 'last_survey' in s:
                state = EHabs.NO_HABS
                for w in s["worlds"]:
                    hab = hab_data[w["type"]]
                    own = 'owner' in w
                    if own:
                        # TODO show room for expansion
                        state = EHabs.ALL_HABS
                        fill_color = 'white'
                    elif state == EHabs.NO_HABS and (hab == 'M' or hab == 'N'):
                        state = EHabs.GOOD_HABS
                        fill_color = '#00C000'
                    elif state == EHabs.NO_HABS and hab == 'P':
                        state = EHabs.BAD_HABS
                        fill_color = 'red'

            canvas.create_text(x + side / 2, y + side / 2, text=star_type, fill=fill_color)

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
    arg_parser.add_argument("habfile", nargs=1, help = "JSON hab file")
    args = arg_parser.parse_args()

    with open(args.turnfile[0], 'r') as turn_file:
        data = json.loads(turn_file.read())

    with open(args.habfile[0], 'r') as hab_file:
        hab_data = json.loads(hab_file.read())["habs"]

    # build the GUI
    root = tk.Tk()

    pane = ttk.PanedWindow(root, orient='horizontal')
    pane.pack(expand=True, fill='both')

    right_frame = ttk.Frame()
    text = tk.Text()

    pane.add(text)
    pane.add(right_frame)

    # star map
    hsb = ttk.Scrollbar(right_frame, orient=tk.HORIZONTAL)
    vsb = ttk.Scrollbar(right_frame, orient=tk.VERTICAL)

    canvas = tk.Canvas(right_frame, yscrollcommand = vsb.set, xscrollcommand = hsb.set)

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

    # parse data
    first_x = data["first_x"] * 10
    first_y = data["first_y"] * 10

    drawData()

    root.mainloop()

