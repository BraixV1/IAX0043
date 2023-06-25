import csv
import os.path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import tkinter as tk

class Data:

    def __init__(self, code: str, postision: str, version: str = None) -> None:
        self.code = code
        self.version = version
        self.positsion = postision

    def __repr__(self) -> str:
        return f"Kood: {self.code} \n kogus: {self.amount} \n Versioon: {self.version} \n Positsioon: {self.positsion}"

    def getCode(self) -> str:
        return self.code
    
    
    def getVersion(self):
        return self.version
    
    def getPosition(self) -> str:
        return self.positsion
    
    def __eq__(self, other) -> bool:
        if isinstance(other, Data):
            return self.code == other.getCode  and self.version == other.getVersion and self.positsion == other.getPosition
        return False

    
class DataBase:


    def __init__(self) -> None:
        self.info = []


    def  __repr__(self) -> str:
        return str(self.info)
    
    def add(self, item: Data):
        self.info.append(item)

    def getDataBase(self):
        return self.info
    
    def __add__(self, other):
        if isinstance(other, DataBase):
            self.info.extend(other.getDataBase())
        return self
    
    def getPositions(self) -> list:
        result = []
        for item in self.info:
            if item.getPosition() not in result:
                result.append(item.getPosition())
        return result


    # For When working with CSV files
    def getData(self, NameOfTheFIle: str) -> None:
        # Opens the file.
        with open(NameOfTheFIle, encoding='utf-8-sig',) as csv_file:

            # Saves the file as a csv.reader object and separates the lines in file to lists of strings which were separated by the delimiter.
            csv_reader = list(csv.reader(csv_file, delimiter=";"))
            self.DatabaseFiller(csv_reader)
    # When working for XLSX files
    def getDataWithXlsx(self, NameOfTheFile: str) -> None:

        workbook = load_workbook(NameOfTheFile)
        sheet = workbook.active

        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(list(row))
        self.DatabaseFiller(data)


    def DatabaseFiller(self, items: list) -> None:
        if "Versioon" in items[0]:
            versionIndex =  items[0].index("Versioon")
        else:
            versionIndex =  None
        koodIndex = items[0].index("Kood")
        positsioonIndex = items[0].index("Positsioon")
        for row in items[1:]:
            positsions = row[positsioonIndex].split(", ")
            for positsion in positsions:
                if versionIndex == None:
                    item = Data(row[koodIndex], positsion)    
                else:
                    item = Data(row[koodIndex], positsion, row[versionIndex])
                self.add(item)


def magic(file_first: str, file_second: str, output: str) -> int:


    if not os.path.isfile(file_first):
        return -1
    if not os.path.isfile(file_second):
        return 0
    

    workbook = Workbook()
    sheet = workbook.active


    red = Font(color="FF0000")
    black = Font(color="000000")
    green = PatternFill(start_color="5bba75", end_color="5bba75", fill_type="solid")
    purple = PatternFill(start_color="c95da0", end_color="c95da0", fill_type="solid")
    
    if file_first.split(".")[-1] == "csv":
        file1 = DataBase()
        file1.getData(file_first)
        fileUnchanged1 = file1.getDataBase()


        file2 = DataBase()
        file2.getData(file_second)
        fileUnchanged2 = file2.getDataBase()
    if file_first.split(".")[-1] == "xlsx":
        file1 = DataBase()
        file1.getDataWithXlsx(file_first)
        fileUnchanged1 = file1.getDataBase()


        file2 = DataBase()
        file2.getDataWithXlsx(file_second)
        fileUnchanged2 = file2.getDataBase()

    positions = list(set((file1.getPositions() + file2.getPositions())))
    positions.sort()
    rows = [["Kood", "Versioon", "Kood", "Versioon", "Positsioon"]]


    for position in positions:
        foundFile1 = list(filter(lambda x: x.getPosition() == position, fileUnchanged1))
        foundFile2 = list(filter(lambda x: x.getPosition() == position, fileUnchanged2))
        if len(foundFile1) > 0 and len(foundFile2) > 0:
            row = [foundFile1[0].getCode(), foundFile1[0].getVersion()
                    , foundFile2[0].getCode(), foundFile2[0].getVersion(), foundFile2[0].getPosition()]
        if len(foundFile1) > 0 and len(foundFile2) == 0:
            row = [foundFile1[0].getCode(), foundFile1[0].getVersion(), "", "", foundFile1[0].getPosition()]
        if len(foundFile1) == 0 and len(foundFile2) > 0:
            row = ["", "", foundFile2[0].getCode(), foundFile2[0].getVersion(), foundFile2[0].getPosition()]
        rows.append(row)


    header_row = rows[0]
    for col_num, value in enumerate(header_row, start=1):
        cell = sheet.cell(row=1, column=col_num)
        if col_num == 1 or col_num == 2:
            cell.fill = green
        if col_num == 3 or col_num == 4:
            cell.fill = purple
        cell.value = value

    data = rows[1:]
    for row_num, row in enumerate(data, start=2):
        for col_num, value in enumerate(row, start=1):
            try:
                value = int(value)
            except ValueError:
                pass
            cell = sheet.cell(row=row_num, column=col_num)
            cell.value = value
            if col_num in [1, 3]:  # Columns 1 and 3
                if col_num == 1:
                    cell.fill = green
                    if(row[col_num - 1] != row[col_num + 1]):
                        cell.font = red
                if col_num == 3:
                    cell.fill = purple
                    if(row[col_num - 1] != row[col_num - 3]):
                        cell.font = red
            elif col_num in [2, 4]:  # Columns 2 and 4
                if col_num == 2:
                    cell.fill = green
                    if(row[col_num - 1] != row[col_num + 1]):
                        cell.font = red
                if col_num == 4:
                    cell.fill = purple
                    if(row[col_num - 1] != row[col_num - 3]):
                        cell.font = red
            elif col_num == 5:  # Column 5
                cell.font = black
    
    workbook.save(output)
    return 1


def on_entry_click(event):
    if entry1.get() == "Enter file name":
        entry1.delete(0, tk.END)


def on_entry_leave(event):
    if entry1.get() == "":
        entry1.insert(0, "Enter file name")


def on_entry2_click(event):
    if entry2.get() == "Enter file name":
        entry2.delete(0, tk.END)


def on_entry2_leave(event):
    if entry2.get() == "":
        entry2.insert(0, "Enter file name")


def on_entry3_click(event):
    if entry3.get() == "Enter output file name":
        entry3.delete(0, tk.END)


def on_entry3_leave(event):
    if entry3.get() == "":
        entry3.insert(0, "Enter output file name")

def CheckFileTypes() -> bool:
    return entry1.get().split(".")[-1] == entry2.get().split(".")[-1]

def CheckFileReadble() -> bool:
    types = ["csv", "xlsx"]
    return entry1.get().split(".")[-1] in types and entry2.get().split(".")[-1] in types



def button_command():
    def close_popup():
        if popup.winfo_exists():  # Check if the popup exists before destroying it
            popup.destroy()

    def close_popup_and_main_window():
        close_popup()
        window.destroy()

    

    result = magic(entry1.get(), entry2.get(), entry3.get())

    popup = tk.Toplevel()

    if not CheckFileTypes():
        popup.title("Type Error")

        label = tk.Label(popup, text="PLease make sure that boths files are of same types")
        label.pack(padx=10, pady=10)

        close_button = tk.Butoon(popup, text="Close", command=close_popup)

    if not CheckFileReadble:
        popup.title("Type Error")

        label = tk.Label(popup, text="Files have to be either csv or xlsx")
        label.pack(padx=10, pady=10)

        close_button = tk.Butoon(popup, text="Close", command=close_popup)

    if result == 1:
        popup.title("Success")

        label = tk.Label(popup, text=f"Operation was completed successfully and the result was saved to {entry3.get()}")
        label.pack(padx=10, pady=10)

        close_button = tk.Button(popup, text="Close", command=close_popup_and_main_window)
        close_button.pack(pady=10)

    elif result == -1:
        popup.title("Error")

        label = tk.Label(popup, text=f"File {entry1.get()} not found!")
        label.pack(padx=10, pady=10)

        close_button = tk.Button(popup, text="Close", command=close_popup)
        close_button.pack(pady=10)

    elif result == 0:
        popup.title("Error")

        label = tk.Label(popup, text=f"File {entry2.get()} not found")
        label.pack(padx=10, pady=10)

        close_button = tk.Button(popup, text="Close", command=close_popup)
        close_button.pack(pady=10)

    popup.deiconify()



if __name__ == "__main__":
    # Create the main window
 # Create the main window
    window = tk.Tk()
    window.title("Difference finder v1.2")
    window.geometry("300x200")

    # Create the first text field (Entry widget) with default text
    entry1 = tk.Entry(window)
    entry1.insert(0, "Enter file name")  # Set default text

    # Bind the entry field events
    entry1.bind("<FocusIn>", on_entry_click)
    entry1.bind("<FocusOut>", on_entry_leave)

    # Create the second text field (Entry widget) with default text
    entry2 = tk.Entry(window)
    entry2.insert(0, "Enter file name")  # Set default text

    entry2.bind("<FocusIn>", on_entry2_click)
    entry2.bind("<FocusOut>", on_entry2_leave)

    # Filed 3

    entry3 = tk.Entry(window)
    entry3.insert(0, "Enter output file name")

    entry3.bind("<FocusIn>", on_entry3_click)
    entry3.bind("<FocusOut>", on_entry3_leave)


    # message box for completing the job

    button = tk.Button(window, text="Show differences", command=button_command)

    entry1.pack()
    entry2.pack()
    entry3.pack()
    button.pack()

    window.mainloop()
    
    
            





