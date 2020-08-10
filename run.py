#!/usr/bin/env python3

import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, '-m', "pip", "install", package])

import tkinter as tk
from tkinter import filedialog
from csv import reader
try:
    import xlrd
except ImportError:
    install("xlrd")

root = tk.Tk()
canvas1 = tk.Canvas(root, width=350, height=300, bg='lightsteelblue')
canvas1.pack()
present = []

late, onTime, total = [], [], []
def attend():
    global path_late 
    path_late = filedialog.askopenfilename()

def absentee():
    global path_total 
    path_total = filedialog.askopenfilename()

def save():
    if 'path_late' not in globals() or 'path_total' not in globals(): return
    files = [('Excel Sheet', '*.xlsx')]
    dest = filedialog.asksaveasfile(mode='w', filetypes=files)
    if dest is None: return
    df = xlrd.open_workbook(path_total)
    sheet = df.sheet_by_index(0)
    for r in range(1, sheet.nrows): total.append(sheet.row_values(r)[3])
    with open(path_late, 'r', encoding='utf-16', errors='ignore') as csvfile:
        i = 0
        data = reader(csvfile)
        for line in data:
            if i==0: i+=1;continue;
            if "Left" in line[0]: continue
            if "Dr. " in line[0]: continue
            if line[0] in onTime or line[0] in late: continue
            name = line[0][:line[0].index('\t')]
            if name not in present: present.append(name)
            time = line[1].split()[0].replace(':','')
            limit = int(time[:2]+"3100")
            if int(time)<limit: onTime.append(line[0])
            else: late.append(line[0][:line[0].index('\t')])
    dest.write("Late: \n")
    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        if row[3] in late: 
            for data in row: 
                dest.write(str(data) + '\t')
            dest.write('\n')
    
    absent = set(total).difference(set(present))
    dest.write("\nAbsentees:\n")
    for r in range(1, sheet.nrows):
        row = sheet.row_values(r)
        if row[3] in absent: 
            for data in row: 
                dest.write(str(data) + '\t')
            dest.write('\n')
    dest.close()

bb_l = tk.Button(text="Import Attendance", command=attend, bg="green", fg="white")
bb_a = tk.Button(text="Import Total List", command=absentee, bg="green", fg="white") 
enter = tk.Button(text="Save", command=save, bg="red", fg="white") 
canvas1.create_window(100,150, window=bb_l)
canvas1.create_window(250,150, window=bb_a)
canvas1.create_window(180,200, window=enter)


root.mainloop()
