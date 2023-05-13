#!/usr/bin/env python3

import json
import tkinter as tk

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
    side = 10

    root = tk.Tk()
    canvas = tk.Canvas(root)

    first_x = data["first_x"] * 10
    first_y = data["first_y"] * 10
    for s in data["stars"]:
        x = (s["xoff"] - first_x) * side
        y = (s["yoff"] - first_y) * side
        t = s["star_type"]
        fill_color = 'black'
        if s["grid"]:
            fill_color = 'dark blue'
        canvas.create_rectangle(x, y, x + side, y + side, fill=fill_color)

    canvas.configure(scrollregion = canvas.bbox('all'))
    canvas.pack(expand=True, fill='both')
    root.mainloop()

