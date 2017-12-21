# import gi
# gi.require_version('Gtk', '3.0')
# import gi.repository.Gtk as Gtk

# class MyWindow(Gtk.Window):

#     def __init__(self):
#         Gtk.Window.__init__(self, title="Hello World")

#         self.button = Gtk.Button()
#         label = Gtk.Label(label="Hello World", angle=25, halign=Gtk.Align.END)
#         self.button.connect("clicked", self.on_button_clicked)
#         self.add(self.button)

#     def on_button_clicked(self, widget):
#         print("Hello World")

# win = MyWindow()
# win.connect("delete-event", Gtk.main_quit)
# win.show_all()
# Gtk.main()
import os
import string
import tkinter
import tkinter.messagebox
import time
import re
import operator

import docx
import docx.enum.text
import docx.shared
import pywinauto

import win32api


'''Automatically select the PowerChart window then select all and copy to clipboard'''

clipboard = []
w_handle = pywinauto.findwindows.find_window(title_re="PowerChart Organizer for")
app = pywinauto.application.Application().connect(handle=w_handle)
window = app.window(handle=w_handle)
window.Maximize()
window.SetFocus()
window.set_keyboard_focus()
window.ClickInput(coords=(640, 330))
#window.DoubleClick(coords=(141, 316))      # coords(141, 316) click on date picker
#window.DoubleClick(coords=(141, 316))
window.Wait('active').TypeKeys('^a')
window.Wait('active').TypeKeys('^c')
time.sleep(1)                               # have to wait for the clipboard to fill up
clipboard = tkinter.Tk().clipboard_get()     # copy contents of clipboard to get date
#window.ClickInput(coords=(640, 339))
#window.TypeKeys('^a')
#window.TypeKeys('^c')
#time.sleep(1)
#clipboard += tkinter.Tk().clipboard_get()   # copy contents of clipboard to get patients

date = ''
# table and months are used to get the date and reformat it
table = str.maketrans({key: None for key in string.punctuation})
months = {
    'January' : '1',
    'February' : '2',
    'March': '3',
    'April' : '4',
    'May' : '5',
    'June' : '6',
    'July' : '7',
    'August' : '8',
    'September' : '9',
    'October' : '10',
    'November' : '11',
    'December' : '12',
    }
lst = clipboard.split()
new_lst = []
time_re = re.compile(r'^([0-1]?[0-9]|[2][0-3]):([0-5][0-9])$')
final_lst = []
line_list = []
fin = open('schedule.txt', 'w+')
for i in range(len(lst)):
    line = ''
    if time_re.match(lst[i]): 
        #and (lst[i + 1] == 'AM' or lst[i + 1] == 'PM'):
        new_lst = lst[i:]
        for item in new_lst:
            if item == 'Years,':
                break
            else:
                line += item + ' '
        line += '\n'
        if 'Block' not in line:
            fin.write(line)
        if 'DX Sleep' not in line:
            line_list = line.split()
            for j in range(len(line_list)):
                provider, patient, appt_time = '', '', ''
                if line_list[j] == 'APRN,':
                    provider = "Mark Boyer, FNP"
                    patient = str(line_list[j + 5]) + str(line_list[j + 6])
                    appt_time = line_list[0]
                    final_lst.append((appt_time, provider, patient))
                elif line_list[j] == 'PA-C,':
                    provider = 'Quinn Ranson, PA-C'
                    patient = str(line_list[j + 5]) + str(line_list[j + 6])
                    appt_time = line_list[0]
                    final_lst.append((appt_time, provider, patient))
                elif line_list[j] == 'DNP,':
                    provider = 'Jennifer Fisher, DNP'
                    patient = str(line_list[j + 5]) + str(line_list[j + 6])
                    appt_time = line_list[0]
                    final_lst.append((appt_time, provider, patient))
                elif line_list[j] == 'MD,':
                    provider = 'Kirk Watkins, MD'
                    patient = str(line_list[j + 5]) + str(line_list[j + 6])
                    appt_time = line_list[0]
                    final_lst.append((appt_time, provider, patient))
                elif line_list[j] == 'DXSD': 
                    provider = 'Josh Conner, CRT, RPSGT'
                    patient = str(line_list[j + 6]) + str(line_list[j + 7])
                    appt_time = line_list[0]
                    final_lst.append((appt_time, provider, patient))
#             line_list = []
# sortedlst = sorted(final_lst, key=operator.itemgetter(1))
# for t in sortedlst:
#     print(t)
fin.close()

import tkinter  # Python 3

def center(win):
    """
    centers a tkinter window
    :param win: the root or Toplevel window to center
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

if __name__ == '__main__':
    root = tkinter.Tk()
    root.attributes('-alpha', 0.0)
    menubar = tkinter.Menu(root)
    filemenu = tkinter.Menu(menubar, tearoff=0)
    filemenu.add_command(label="Exit", command=root.destroy)
    menubar.add_cascade(label="File", menu=filemenu)
    root.config(menu=menubar)
    frm = tkinter.Frame(root, bd=4, relief='raised')
    frm.pack(fill='x')
    lab = tkinter.Label(frm, text='Hello World!', bd=4, relief='sunken')
    lab.pack(ipadx=4, padx=4, ipady=4, pady=4, fill='both')
    center(root)
    root.attributes('-alpha', 1.0)
    root.mainloop()