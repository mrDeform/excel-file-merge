import datetime
import xlrd
from os import path
from pandas import read_excel
from pandas import ExcelWriter
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo


class App:
    def __init__(self, window):
        self.window = window
        self.window.title("concatenatio")

        self.btn = Button(window, text="Merge Excel files", width=20, command=self.running)
        self.btn.pack(side=TOP, pady=20)

        self.exit_button = Button(window, text='Exit Program', width=20, command=self.window.destroy)
        self.exit_button.pack(side=TOP, pady=20)

        self.window.mainloop()

    def running(self):
        filenames = fd.askopenfilenames(filetypes=[("Excel files", ".xlsx")])

        if filenames != "":
            week = (datetime.datetime.now().isocalendar()[1]) - 1
            year = (datetime.datetime.now().isocalendar()[0]) % 2000
            final_filename = 'Alarm MBH_H-W{0}{1}.xlsx'.format(year, week)

            with ExcelWriter(final_filename) as writer:
                for cell in filenames:
                    read_excel(xlrd.open_workbook(cell), sheet_name=0, header=None)\
                        .to_excel(writer, sheet_name=path.splitext(path.basename(cell))[0], index=False, header=0)

            showinfo(title='SAVED', message='Success')
            self.window.destroy()


def main():
    root = Tk()
    App(root)


if __name__ == '__main__':
    main()

# developed by mrDeform
