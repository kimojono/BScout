import tkinter.filedialog
from tkinter import *
import re

import os
import json
import openpyxl
from openpyxl.utils.cell import get_column_letter

MATH_OPERATORS = ['+','-','*','/','(',')','^','=']
KEYCODES_LIST = [49, 50, 51, 52, 53, 54, 55, 56, 57, 48, 81, 87, 69, 82, 84, 89, 85, 73, 79, 80, 65,
                 83, 68, 70, 71, 72, 74, 75, 76, 90, 88, 67, 86, 66, 78, 77]
BAD_FILE_NAME_CHARS = ['\\', '/', '*', ':', '"', '?', '|', '<', '>']
global root, counters_list, box_score, add_or_subtract, CONFIG
HEADER_FONT = ("Calibri Light", 14)
TEXT_FONT = ("Calibri Light", 10)

def get_formula(formula_string: str, table_index: str, player_row: int = 2):
    formula = formula_string.split("=")
    name = formula[0]
    formula = "=" + formula[1]
    for counter in CONFIG["player counters names list"] + CONFIG["general counters names list"]:
        counter_indexes_list = [m.start() for m in re.finditer(counter, formula)]
        for substring_index in counter_indexes_list[::-1]:
            if formula[substring_index-1] in MATH_OPERATORS and formula[substring_index+len(counter)] in MATH_OPERATORS:
                formula = formula[:substring_index] + f'HLOOKUP("{counter}",{table_index},{player_row},FALSE)' + formula[substring_index+len(counter):]
    return [name, formula]


class Counter:
    def __init__(self, window, row: int, column: int, name="empty", count=0):
        self.name = name
        self.count = count
        self.row = row
        self.column = column
        self.window = window


class EntryCounter(Counter):
    def __init__(self, window, row: int, column: int, name="empty", count=0):
        super().__init__(window, row, column, name, count)
        self.entry = Entry(window, width=8, borderwidth=2)


class KeysCounter(Counter):
    def __init__(self, window, key: int, row: int, column: int, name="empty", count=0):
        super().__init__(window, row, column, name, count)
        self.key = key
        self.label = None
        self.display()

    def add_or_subtract_one(self, add_or_subtract_flag):
        if add_or_subtract_flag:
            self.count -= 1
        else:
            self.count += 1
        self.display()

    def display(self):
        Label(self.window, text=f"{chr(self.key).lower()}", font=TEXT_FONT, pady=1, bg="lightblue1") \
            .grid(row=self.row, column=self.column)
        Label(self.window, text=f"{self.name}", font=TEXT_FONT, pady=1, bg="lightblue1") \
            .grid(row=self.row, column=self.column + 1)
        self.label = Label(self.window, text=f"{self.count}", font=TEXT_FONT, pady=1, bg="lightblue1")
        self.label.grid(row=self.row, column=self.column + 2)


def init_root_screen(filename: str):
    global counters_list, add_or_subtract
    counters_per_column = 22
    '''
    background_image = ImageTk.PhotoImage(file="background_image.jpg")
    background_label = Label(root,image=background_image)
    background_label.image=background_image
    background_label.grid(row=0,column=0)
    '''  # adds background image, currently not working
    counters_list = []
    add_or_subtract = False
    for i in range(len(CONFIG["player counters names list"])):
        counters_list.append(
            KeysCounter(root, KEYCODES_LIST[i], int(i - int(i / counters_per_column) * counters_per_column) + 3, int(i / counters_per_column) * 3,
                        CONFIG["player counters names list"][i]))
    for i in range(int((len(counters_list)-1) / counters_per_column) + 1):
        Label(root, text="Key", font=TEXT_FONT, bg='lightblue1', fg="black", bd=3) \
            .grid(row=2, column=i * 3)
        Label(root, text="Content", font=TEXT_FONT, bg='lightblue1', fg="black", bd=3) \
            .grid(row=2, column=i * 3 + 1)
        Label(root, text="Count", font=TEXT_FONT, bg='lightblue1', fg="black", bd=3) \
            .grid(row=2, column=i * 3 + 2)
    Label(root, text=filename, font=HEADER_FONT, bg='lightblue1', padx=30, bd=3) \
        .grid(row=0, column=0, columnspan=(int(len(counters_list) / counters_per_column) + 1) * 3)
    Label(root, text="Press '-' To Subtract", font=TEXT_FONT, bg='lightblue1', padx=30, bd=3) \
        .grid(row=1, column=0, columnspan=(int(len(counters_list) / counters_per_column) + 1) * 3)
    export_button = Button(root, text="Export", font=TEXT_FONT, bg='snow', padx=30, bd=3,
                           command=export_to_excel)
    import_button = Button(root, text="Import", font=TEXT_FONT, bg='snow', padx=30, bd=3,
                           command=import_from_excel)
    export_button.grid(row=counters_per_column+5, column=int((len(counters_list)-1) / counters_per_column) * 3, sticky=E, columnspan=3)
    import_button.grid(row=counters_per_column+5, column=0, sticky=W, columnspan=2)
    root.resizable(False, False)
    root.eval('tk::PlaceWindow . center')


def key_pressed(event):
    global add_or_subtract
    if event.char == '-':
        add_or_subtract = not add_or_subtract
    for i in range(len(counters_list)):
        if event.keycode == KEYCODES_LIST[i]:
            counters_list[i].add_or_subtract_one(add_or_subtract)


def save_and_close(file_name, player_name, window_to_destroy):
    flag = True
    for counter in box_score:
        if counter.entry.get().isdecimal():
            counter.count = counter.entry.get()
        else:
            flag = False
            break
    if file_name != '' and not any([c in BAD_FILE_NAME_CHARS for c in file_name]) and flag:
        file_full_location = rf"{CONFIG['SAVE LOCATION']}\{file_name}.xlsx"
        export_results = openpyxl.Workbook()
        data_sheet = export_results.active
        data_sheet.title = "Data"
        columns = ["player name"] + [i.name for i in box_score if i.name != 'empty'] + \
                  [i.name for i in counters_list if i.name != 'empty']
        data = [player_name] + [i.count for i in box_score if i.name != 'empty'] + \
               [i.count for i in counters_list if i.name != 'empty']
        data_sheet.append(columns)
        data_sheet.append(data)
        statistics_sheet = export_results.create_sheet("statistics")
        statistics_sheet["A1"] = "player name"
        statistics_sheet["A2"] = player_name
        for formula_string_index in range(len(CONFIG["ADVANCED STATS"])):
            temp_formula = get_formula(CONFIG["ADVANCED STATS"][formula_string_index], "Data!" + data_sheet.dimensions)
            statistics_sheet[f"{get_column_letter(formula_string_index + 2)}1"] = temp_formula[0]
            statistics_sheet[f"{get_column_letter(formula_string_index + 2)}2"] = temp_formula[1]
        export_results.save(file_full_location)
        window_to_destroy.destroy()


def export_to_excel():
    global box_score
    export_window = Toplevel(root)
    export_window.resizable(False, False)
    export_window.title = "name the file"
    Label(export_window, text="file name:", font=TEXT_FONT).grid(row=0, column=0, columnspan=2)
    file_name_entry = Entry(export_window, width=35, borderwidth=5)
    file_name_entry.grid(row=0, column=2, columnspan=3)
    file_name_entry.focus()
    Label(export_window, text="player name:", font=TEXT_FONT).grid(row=1, column=0, columnspan=2)
    player_name_entry = Entry(export_window, width=20, borderwidth=3)
    player_name_entry.grid(row=1, column=2, columnspan=2)
    Label(export_window, text="", font=TEXT_FONT).grid(row=2, column=0)
    box_score = []
    for i in range(len(CONFIG["general counters names list"])):
        box_score.append(
            EntryCounter(export_window, i % 16 + 3, 1 + int(i / 16) * 2, CONFIG["general counters names list"][i]))
        box_score[i].entry.grid(row=box_score[i].row, column=box_score[i].column)
        box_score[i].entry.insert(-1, "0")
        Label(export_window, text=f'{CONFIG["general counters names list"][i]}:', font=("Calibri Light", 12)) \
            .grid(row=i % 16 + 3, column=+int(i / 16) * 2)
    Button(export_window, text="done", font=TEXT_FONT,
           command=lambda: save_and_close(file_name_entry.get(), player_name_entry.get(), export_window)) \
        .grid(row=1200, column=2, columnspan=2)
    export_window.bind('<Return>',
                       lambda event: save_and_close(file_name_entry.get(), player_name_entry.get(), export_window))


def import_from_excel():
    filename = tkinter.filedialog.askopenfilename(filetypes=[("Excel file", "*.xlsx")],
                                                  initialdir=CONFIG["SAVE LOCATION"])
    new_window = Toplevel(root)
    new_window.title = "calculation finished"
    new_window.resizable(False, False)
    new_window.geometry(f"+{int(root.winfo_screenwidth() / 2)}+{int(root.winfo_screenheight() / 2)}")
    if filename != '':
        book = openpyxl.load_workbook(filename)
        if "statistics" not in book:
            import_results = openpyxl.load_workbook(filename)
            data_sheet = import_results.active
            statistics_sheet = import_results.create_sheet("statistics")
            for row_index in range(1, data_sheet.max_row + 1):
                statistics_sheet[f"A{row_index}"] = f'=HLOOKUP("player name",{data_sheet.title}!{data_sheet.dimensions},{row_index},FALSE)'
            for formula_string_index in range(len(CONFIG["ADVANCED STATS"])):
                temp_formula = ["", ""]
                for row_index in range(2, data_sheet.max_row + 1):
                    temp_formula = get_formula(CONFIG["ADVANCED STATS"][formula_string_index],
                                               f"{data_sheet.title}!{data_sheet.dimensions}",row_index)
                    statistics_sheet[f"{get_column_letter(formula_string_index + 2)}{row_index}"] = temp_formula[1]
                statistics_sheet[f"{get_column_letter(formula_string_index + 2)}1"] = temp_formula[0]

            import_results.save(filename)
            Label(new_window, text="Finished! everything worked", padx=30, pady=15).pack()
        else:
            Label(new_window, text="statistics already exists", padx=30, pady=15).pack()
    else:
        Label(new_window, text="didn't choose a file", padx=30, pady=15).pack()

    Button(new_window, text="close", command=new_window.destroy).pack()
    new_window.focus()
    new_window.bind('<Return>', lambda event: new_window.destroy)


def main():
    global root, CONFIG
    filename = tkinter.filedialog.askopenfilename(filetypes=[("Json File", "*.json")],
                                                  initialdir=os.getcwd() + "/configurations")
    root = Tk()
    root.configure(bg="lightblue1")
    with open(filename) as file:
        CONFIG = json.load(file)
        init_root_screen(filename.split("/")[-1].split(".")[0])
        root.bind('<Key>', key_pressed)
        root.focus()
        root.mainloop()


if __name__ == "__main__":
    main()
