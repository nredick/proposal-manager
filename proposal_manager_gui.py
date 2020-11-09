import os
import time
import tkinter as tk
from tkinter import *
from tkinter import filedialog

import proposal_creator

clicked = False


def click():
    global clicked
    more_info.pack(fill="both", expand="no")
    if not clicked and more_info.winfo_exists():
        more_info.pack_forget()  # destroys the label
    clicked = not clicked


def create():
    more_info.pack_forget()
    start_time = time.time()
    out.insert(INSERT, f'\n{time.ctime()} Starting proposal creation...\n')
    root.update()
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=[("Excel files", "*.xlsx")])
    result = proposal_creator.create_proposal(filename)  # create proposal
    out.pack(fill="both", expand="no")
    for i, line in enumerate(result):
        out.insert(END, f'{line}\n')
        #time.sleep(.15)
        root.update()
    end_time = time.time()
    out.insert(END, f'{time.ctime()} Finished proposal creation. Elapsed time: {end_time - start_time}')
    out.insert(END, f'{time.ctime()} Opening proposal in Microsoft Excel...')
    root.update()
    #time.sleep(3)
    os.system(f'open \"{filename}\" -a \"Microsoft Excel\"')  # open the excel workbook
    #root.destroy()


# create tkinter window root
root = Tk()
root.configure(background='#abd3be')

# set window title
root.title("Redbud Proposal Manager")

# window styling
root.geometry('625x500')

# more info label
more_info = tk.Text(root, bg='#abd3be', bd=0, height=7, yscrollcommand=True, xscrollcommand=True)
more_info.insert(INSERT, "How to Use the Proposal Manager:\n\nThe program accepts an input file of the *.xlsx type using the \'Create Proposal\' button.\nData will be pulled from the sheet in the first/0 position within the Excel workbook.\nA budget proposal will be created and inserted in the second/1st position and the\nfile will automatically be opened in Excel when the program finishes.\n")

# output label
out = tk.Text(root, bg='#abd3be', bd=0, highlightthickness=0, yscrollcommand=True, xscrollcommand=True)

# bottom
bottom = tk.Label(root, text="August 2020. Created by Nathalie Redick", bg='#abd3be', foreground='#2c543e')
bottom.pack(side='bottom')

# button for proposal creation
buttonframe = tk.Frame(root, bg='#abd3be')
buttonframe.pack(fill='both', expand='no')

create_button = tk.Button(buttonframe, text="Create Proposal", command=create, bg='#abd3be', foreground='#2c543e', padx=6, pady=6)
create_button.pack(padx=5, pady=5)

# button for more info/directions
info_button = tk.Button(buttonframe, text="More Info", command=click, bg='#abd3be', foreground='#2c543e', padx=6, pady=6)
info_button.pack(padx=5, pady=5)

# show window
root.mainloop()

