from tkinter import Tk, BOTH
from tkinter import ttk
from tkinter.ttk import Frame, Button, Style


class Window(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent)

        self.parent = parent

        self.initUI()


    def initUI(self):

        self.parent.title("Quit button")
        self.style = Style()
        self.style.theme_use("default")
        self.pack(fill=BOTH, expand=1)
        quitButton = Button(self, text="Quit",
            command=self.quit)
        quitButton.place(x=50, y=50)
        w = 290
        h = 150

        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()

        x = (sw - w) / 2
        y = (sh - h) / 2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))


def main():

    root = Tk()
    root.geometry("250x150+300+300")
    app = Window(root)
    root.mainloop()


if __name__ == '__main__':
    main()