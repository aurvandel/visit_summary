# from tkinter import *
#
#
# def donothing():
#     filewin = Toplevel(root)
#     button = Button(filewin, text="Do nothing button")
#     button.pack()
#
#
# root = Tk()
# menubar = Menu(root)
# filemenu = Menu(menubar, tearoff=0)
# filemenu.add_command(label="New", command=donothing)
# filemenu.add_command(label="Open", command=donothing)
# filemenu.add_command(label="Save", command=donothing)
# filemenu.add_command(label="Save as...", command=donothing)
# filemenu.add_command(label="Close", command=donothing)
#
# filemenu.add_separator()
#
# filemenu.add_command(label="Exit", command=root.quit)
# menubar.add_cascade(label="File", menu=filemenu)
# editmenu = Menu(menubar, tearoff=0)
# editmenu.add_command(label="Undo", command=donothing)
#
# editmenu.add_separator()
#
# editmenu.add_command(label="Cut", command=donothing)
# editmenu.add_command(label="Copy", command=donothing)
# editmenu.add_command(label="Paste", command=donothing)
# editmenu.add_command(label="Delete", command=donothing)
# editmenu.add_command(label="Select All", command=donothing)
#
# menubar.add_cascade(label="Edit", menu=editmenu)
# helpmenu = Menu(menubar, tearoff=0)
# helpmenu.add_command(label="Help Index", command=donothing)
# helpmenu.add_command(label="About...", command=donothing)
# menubar.add_cascade(label="Help", menu=helpmenu)
#
# root.config(menu=menubar)
# root.mainloop()

from tkinter import *

root = Tk()
root.title("Tk dropdown example")

# Add a grid
mainframe = Frame(root)
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)
mainframe.pack(pady=100, padx=100)

# Create a Tkinter variable
tkvar = StringVar(root)

# Dictionary with options
choices = {'Pizza', 'Lasagne', 'Fries', 'Fish', 'Potatoe'}
tkvar.set('Pizza')  # set the default option

popupMenu = OptionMenu(mainframe, tkvar, *choices)
Label(mainframe, text="Choose a dish").grid(row=1, column=1)
popupMenu.grid(row=2, column=1)


# on change dropdown value
def change_dropdown(*args):
    print(tkvar.get())
    return tkvar.get()


# link function to change dropdown
tkvar.trace('w', change_dropdown)

root.mainloop()
root.destroy()