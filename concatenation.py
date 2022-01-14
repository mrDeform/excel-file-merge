import xlrd
import time
import datetime
from os import path
from pandas import read_excel
from pandas import ExcelWriter
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from threading import Thread


def get_final_filename():
    # final filename with last week and current year
    week = (datetime.datetime.now().isocalendar()[1]) - 1
    year = (datetime.datetime.now().isocalendar()[0]) % 2000
    return 'Alarm MBH_H-W{0}{1}.xlsx'.format(year, week)


class App:
    def __init__(self, window):
        self.window = window
        self.window.title("concatenation")

        self.files_count = 1
        self.start_time = 0
        self.file_upload_speed = 0
        self.filenames = tuple()
        self.thread = Thread(target=self.merge_files)

        self.progress = Progressbar(window, orient=HORIZONTAL, length=130, mode='determinate')

        self.btn_merge = Button(window, text="Merge Excel files", width=20, command=self.load_files)
        self.btn_merge.pack(side=TOP, pady=20, padx=60)

        self.exit_button = Button(window, text='Exit Program', width=20, command=self.window.destroy)
        self.exit_button.pack(side=BOTTOM, pady=20, padx=60)

        self.window.mainloop()

    def progressbar_state(self):
        self.progress['value'] += (47 / self.files_count)

        if round(self.progress['value']) < 93:
            self.window.update()

            if round(self.progress['value']) == 47:
                self.file_upload_speed = int(((time.time() - self.start_time) / self.files_count) * 1150)
                self.window.after(self.file_upload_speed, self.progressbar_state)

            if round(self.progress['value']) > 47:
                self.window.after(self.file_upload_speed, self.progressbar_state)

    def load_files(self):
        self.filenames = fd.askopenfilenames(filetypes=[("Excel files", ".xlsx")])
        self.files_count = len(self.filenames)
        if self.filenames != "":
            self.collect_final_file()

    def finish_thread(self):
        self.progress['value'] = 100
        self.window.update()
        showinfo(title='SAVED', message='Success')
        self.exit_button['state'] = 'enable'

    def merge_files(self):
        final_filename = get_final_filename()
        # combine files into one file
        with ExcelWriter(final_filename) as writer:
            for cell in self.filenames:
                # read the first sheet of the open file and add it to the final file
                read_excel(xlrd.open_workbook(cell), sheet_name=0, header=None) \
                    .to_excel(writer, sheet_name=path.splitext(path.basename(cell))[0], index=False, header=0)
                self.progressbar_state()
        self.finish_thread()

    def collect_final_file(self):
        self.btn_merge.pack_forget()
        self.progress.pack(side=TOP, pady=20, padx=20)
        self.exit_button['state'] = 'disabled'
        self.start_time = time.time()
        self.thread.start()


def main():
    root = Tk()
    App(root)


if __name__ == '__main__':
    main()

# developed by mrDeform