# Imports:
import pandas as pd
from tkinter import *
from tkinter.ttk import Progressbar
import time
from tkinter.filedialog import askdirectory, askopenfilename

# Init:

column_list = ['column1',
               'column2',
               'column3',
               'column4',
               'column5',
               'column6',
               'column7',
               'column8']


# Methods:
def select_dir_open():
    global read_dir
    read_dir = askopenfilename()  # Get the directory of the required file
    get_dir.destroy()


def select_dir_save():
    global save_dir
    save_dir = askdirectory()  # Set the directory to save the results
    save_to.destroy()


def assign_month(value_m):
    global month
    month = value_m
    set_month.destroy()


def assign_part(value_p):
    global part
    part = value_p
    set_part.destroy()


def open_file():
    global get_dir
    get_dir = Tk()
    get_dir.geometry("500x500")
    """ 'get_dir' loop: """
    # Labels:
    Label(get_dir, text="Welcome, please select the relevant Recon Data file.", pady=50).pack()
    # Buttons:
    Button(get_dir, text="OK", padx=50, pady=50, command=select_dir_open).pack()
    get_dir.mainloop()


def file_pre_proccess():
    wait1 = Tk()
    wait1.geometry("500x500")
    Label(wait1, text="Processing files, please wait.", pady=50).pack()

    def step():
        global df_pb
        global df_hf
        global df_el
        global df_el_ou
        my_progres['value'] += (100 / 7)
        wait1.update_idletasks()
        time.sleep(1)
        df = pd.read_excel(read_dir)  # Load the original file to copy the data
        wait1.update_idletasks()
        my_progres['value'] += (100 / 7)
        time.sleep(1)
        df.drop(columns=df.columns.difference(column_list), inplace=True)
        wait1.update_idletasks()
        my_progres['value'] += (100 / 7)
        time.sleep(1)
        df.rename(columns={"column1": "diffrent name"}, inplace=True)
        wait1.update_idletasks()
        my_progres['value'] += (100 / 7)
        time.sleep(1)
        df_pb = df[df['column2'].isin(['column1',column1])]
        wait1.update_idletasks()
        my_progres['value'] += (100 / 7)
        time.sleep(1)
        df_hf = df[df['LP'] == 'LP3']
        wait1.update_idletasks()
        my_progres['value'] += (100 / 7)
        time.sleep(1)
        df_el = df[df['LP'].isin(['a', 'b'])]
        wait1.update_idletasks()
        my_progres['value'] += (100 / 7)
        df_el_ou = df[df['LP'] == 'c']
        time.sleep(1)
        wait1.destroy()

    my_progres = Progressbar(wait1, orient=HORIZONTAL, length=300, mode='determinate')
    my_progres.pack(expand=True)
    wait1.after(2000, step)
    wait1.mainloop()


def save_locaiton():
    global save_to
    save_to = Tk()
    save_to.geometry("500x500")
    """ 'save_to' loop: """
    # Labels:
    Label(save_to, text="Please Choose where to save the Trx data", pady=50).pack()
    # Buttons:
    Button(save_to, text="OK", padx=50, pady=50, command=select_dir_save).pack()
    save_to.mainloop()


def select_month():
    global set_month
    set_month = Tk()
    set_month.geometry("500x500")
    set_month.title("Month Selection Window")
    set_month.columnconfigure(0, weight=1)
    set_month.columnconfigure(1, weight=1)
    set_month.columnconfigure(2, weight=1)
    set_month.columnconfigure(3, weight=1)
    set_month.rowconfigure(0, weight=1)
    set_month.rowconfigure(1, weight=1)
    set_month.rowconfigure(2, weight=1)
    set_month.rowconfigure(3, weight=1)
    """ 'set_month' loop: """
    # Buttons:
    Button(set_month, text="January", width=5, height=2, padx=25, pady=25,
           command=lambda: assign_month("January")).grid(row=1, column=0, padx=1)
    Button(set_month, text="February", width=5, height=2, padx=25, pady=25,
           command=lambda: assign_month("February")).grid(row=1, column=1, padx=1)
    Button(set_month, text="March", width=5, height=2, padx=25, pady=25, command=lambda: assign_month("March")).grid(
        row=1, column=2, padx=1)
    Button(set_month, text="April", width=5, height=2, padx=25, pady=25, command=lambda: assign_month("April")).grid(
        row=2, column=0, padx=1)
    Button(set_month, text="May", width=5, height=2, padx=25, pady=25, command=lambda: assign_month("May")).grid(
        row=2, column=1, padx=1)
    Button(set_month, text="June", width=5, height=2, padx=25, pady=25, command=lambda: assign_month("June")).grid(
        row=2, column=2, padx=1)
    Button(set_month, text="July", width=5, height=2, padx=25, pady=25, command=lambda: assign_month("July")).grid(
        row=3, column=0, padx=1)
    Button(set_month, text="August", width=5, height=2, padx=25, pady=25, command=lambda: assign_month("August")).grid(
        row=3, column=1, padx=1)
    Button(set_month, text="September", width=5, height=2, padx=25, pady=25,
           command=lambda: assign_month("September")).grid(row=3, column=2, padx=1)
    Button(set_month, text="October", width=5, height=2, padx=25, pady=25,
           command=lambda: assign_month("October")).grid(row=4, column=0, padx=1)
    Button(set_month, text="November", width=5, height=2, padx=25, pady=25,
           command=lambda: assign_month("November")).grid(row=4, column=1, padx=1)
    Button(set_month, text="December", width=5, height=2, padx=25, pady=25,
           command=lambda: assign_month("December")).grid(row=4, column=2, padx=1)
    # Labels:
    Label(set_month, text="Please select Month of Reconciliation").grid(row=0, column=1, padx=1)

    set_month.mainloop()


def close():
    global clos
    close = Tk()
    close.geometry("500x500")
    Label(close, text="Task Completed", pady=50).pack()
    Button(close, text="Close", padx=50, pady=50, command=close.destroy).pack()
    close.mainloop()


def select_part():
    global set_part
    set_part = Tk()
    set_part.geometry("500x500")
    set_part.title("Part Selection Window")
    set_part.columnconfigure(0, weight=1)
    set_part.columnconfigure(1, weight=1)
    set_part.columnconfigure(2, weight=1)
    set_part.columnconfigure(3, weight=1)
    set_part.rowconfigure(0, weight=1)
    set_part.rowconfigure(1, weight=1)
    set_part.rowconfigure(2, weight=1)
    set_part.rowconfigure(3, weight=1)

    ''' 'set_part' loop: '''
    # Buttons:
    Button(set_part, text="Part 1", command=lambda: assign_part("Part 1")).grid(row=1, column=1, ipadx=10, ipady=10)
    Button(set_part, text="Part 2", command=lambda: assign_part("Part 2")).grid(row=1, column=2, ipadx=10, ipady=10)
    # Labels:
    Label(set_part, text="Please select part of Reconciliation").grid(row=0, column=1)

    set_part.mainloop()


def save_files():
    wait2 = Tk()
    wait2.geometry("500x500")
    Label(wait2, text="Saving files, please wait.", pady=50).pack()

    def step():
        my_progress2['value'] += (100 / 5)
        wait2.update_idletasks()
        time.sleep(1)
        df_hf.to_excel(save_dir + '/' + 'Trx data LP1 ' + month + ' ' + part + '.xlsx',
                       index=False,
                       header=True)
        wait2.update_idletasks()
        my_progress2['value'] += (100 / 5)
        time.sleep(1)
        df_pb.to_excel(save_dir + '/' + 'Trx data LP2 ' + ' ' + month + ' ' + part + '.xlsx',
                       index=False,
                       header=True)
        wait2.update_idletasks()
        my_progress2['value'] += (100 / 5)
        time.sleep(1)
        df_el.to_excel(save_dir + '/' + 'Trx data LP3 ' + ' ' + month + ' ' + part + '.xlsx',
                       index=False,
                       header=True)
        wait2.update_idletasks()
        my_progress2['value'] += (100 / 5)
        time.sleep(1)
        df_el_ou.to_excel(save_dir + '/' + 'Trx data LP4 ' + ' ' + month + ' ' + part + '.xlsx',
                          index=False,
                          header=True)
        wait2.update_idletasks()
        my_progress2['value'] += (100 / 5)
        time.sleep(1)
        wait2.update_idletasks()
        wait2.destroy()

    my_progress2 = Progressbar(wait2, orient=HORIZONTAL, length=300, mode='determinate')
    my_progress2.pack(expand=True)
    wait2.after(2000, step)
    wait2.mainloop()


if __name__ == '__main__':
    open_file()
    file_pre_proccess()
    save_locaiton()
    select_month()
    select_part()
    save_files()
    close()
