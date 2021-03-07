import os
import time
from tkinter import *
from tkinter import filedialog
import rough_budget
import actual_budget

clicked = False


def user_guide():
    if not clicked:
        string = StringVar()
        msg = "How to Use the Proposal Manager" \
              "\n\nThe program accepts an input file of the type *.xlsx where the first " \
              "sheet of the workbook matches the budget proposal format. Data " \
              "will be taken from the first sheet in the file. When the program is done, " \
              "the file with the budget in the second position will be opened in Excel.\n" \
              "Note that the Rough and Actual budget proposal creators run independently (i.e. you can create one " \
              "without needing to create the other.\n" \
              "A new template file can be created in the File dropdown menu."
        string.set(msg)

        label = Message(root, textvariable=string, relief='raised')
        label.pack()
    else:
        pass


def create_rough():
    start_time = time.time()
    out.insert(INSERT, f'\n{time.ctime()} Starting rough budget creation...\n')
    root.update()
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=[("Excel files", "*.xlsx")])
    result = rough_budget.create_proposal(filename)  # create proposal
    out.pack(fill="both", expand="no")
    for i, line in enumerate(result):
        out.insert(END, f'{line}\n')
        #time.sleep(.15)
        root.update()
    end_time = time.time()
    out.insert(END, f'{time.ctime()} Finished rough budget creation. Elapsed time: {end_time - start_time}')
    out.insert(END, f'{time.ctime()} Opening proposal in Microsoft Excel...')
    root.update()
    #time.sleep(3)
    os.system(f'open \"{filename}\" -a \"Microsoft Excel\"')  # open the excel workbook
    #root.destroy()


def create_actual():
    start_time = time.time()
    out.insert(INSERT, f'\n{time.ctime()} Starting actual budget creation...\n')
    root.update()
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=[("Excel files", "*.xlsx")])
    result = actual_budget.create_proposal(filename)  # create proposal
    out.pack(fill="both", expand="no")
    for i, line in enumerate(result):
        out.insert(END, f'{line}\n')
        #time.sleep(.15)
        root.update()
    end_time = time.time()
    out.insert(END, f'{time.ctime()} Finished actual budget creation. Elapsed time: {end_time - start_time}')
    out.insert(END, f'{time.ctime()} Opening proposal in Microsoft Excel...')
    root.update()
    #time.sleep(3)
    os.system(f'open \"{filename}\" -a \"Microsoft Excel\"')  # open the excel workbook
    #root.destroy()


def create_new():
    fn = "NEW client estimating template $(date +\"%b-%r\").xlsx"
    os.system(f'cp "client estimating template.xlsx" {fn}')
    os.system(f'open {fn}')


# colours & styles
fg = '#2f374a'
bg = '#566a99'
body_font = ('calibri', 10, 'bold', 'underline')

# set up the tk window
root = Tk()  # create tkinter window root
root.configure(background=bg)
root.title("Redbud Proposal Manager")  # set window title

# set up the menu bar
menubar = Menu(root)

file_menu = Menu(menubar, tearoff=0)
file_menu.add_command(label="New", command=create_new)
file_menu.add_separator()
file_menu.add_command(label="Quit", command=root.quit)
menubar.add_cascade(label="File", menu=file_menu)

help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="User Guide", command=user_guide)
menubar.add_cascade(label="Help", menu=help_menu)

# window styling
root.geometry('625x500')

# output label
out = Text(root, bg=bg, bd=0, highlightthickness=0, yscrollcommand=True, xscrollcommand=True)

# footer
footer = Label(root, text="March 2021. Created by Nathalie Redick", bg=bg, foreground=fg)
footer.pack(side='bottom')

buttonframe = Frame(root, bg=bg)
buttonframe.pack(fill='both', expand='no')

# button for rough budget creation
creater_button = Button(buttonframe,
                        relief='groove',
                        text="Create Rough Budget",
                        command=create_rough,
                        bg=bg, foreground=fg,
                        padx=6, pady=6)
creater_button.pack(padx=5, pady=5)

# button for actual budget creation
createa_button = Button(buttonframe, compound=CENTER,
                        relief='groove',
                        text="Create Actual Budget",
                        command=create_actual,
                        bg=bg, foreground=fg,
                        padx=6, pady=6)
createa_button.pack(padx=5, pady=5)

# show window
root.config(menu=menubar)
root.mainloop()
