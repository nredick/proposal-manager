import os
import tkinter as tk
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog

import proposal_creator

clicked = False


def info():
    global clicked
    if not clicked:
        clicked = True
        labelframe = LabelFrame(root, text="How to Use the Proposal Manager")
        labelframe.pack(fill="both", expand="no")
        left = Label(labelframe, text="\nThe program accepts an input file of the *.xlsx type using the \'Create Proposal\' button.\nData will be pulled from the sheet in the first/0 position within the Excel workbook.\nA budget proposal will be created and inserted in the second/1st position and the\nfile will automatically be opened in Excel when the program finishes.\n")
        left.pack()


def create():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=[("Excel files", "*.xlsx")])
    proposal_creator.create_proposal(filename)  # create proposal
    os.system(f'open \"{filename}\" -a \"Microsoft Excel\"')  # open the excel workbook


root = Tk()
root.configure(background='#abd3be')


# set window title
root.title("Redbud Proposal Manager")

# window styling
root.geometry('600x400')

#bottom
bottom = tk.Label(root, text="August 2020. Created by Nathalie Redick", bg='#abd3be', foreground='#2c543e')
bottom.pack(side='bottom')

# button for proposal creation
buttonframe = tk.Frame(root, bg='#abd3be')
buttonframe.pack(fill='both', expand='no')

create_button = tk.Button(buttonframe, text="Create Proposal", command=create, bg='#abd3be', foreground='#2c543e', padx=6, pady=6)
create_button.pack(padx=5, pady=5)

# button for more info/directions
info_button = tk.Button(buttonframe, text="More Info", command=info, bg='#abd3be', foreground='#2c543e', padx=6, pady=6)
info_button.pack(padx=5, pady=5)

# show window
root.mainloop()
