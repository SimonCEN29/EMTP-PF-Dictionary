import tkinter
from tkinter import filedialog
from pathlib import Path
import csv
import os

class Directory:
    def __init__(self, saving_path: str=None):
        if saving_path:
            self.saving_path = Path(saving_path)
        else:
            root = tkinter.Tk()
            root.withdraw()  # Manually closing tkinter window.

            loop_condition = True
            while loop_condition:
                selected_path = filedialog.askdirectory()
                if not selected_path == "":
                    loop_condition = False
                else:
                    print("No folder was selected. Please select a folder.\n")
            print("Files will be saved at: \"" + selected_path + "\".")
            self.saving_path = Path(selected_path)
    
    def print_headers(self, csv_headers: dict, encoding="utf-8"):
        csv_files = []
        for key, headers in csv_headers.items():
            csv_path = self.saving_path / Path(key + ".csv")
            csv_files.append(csv_path)

            try:
                with csv_path.open("w", newline="", encoding=encoding) as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
            except PermissionError:
                print("One or more files are open. Close them and run the routine again.")

        return csv_files

    
if __name__ == "__main__":
    os.system('cls')
    dir = Directory(r"C:\Git Codes\EMTP-FP Dic Output")

    csv_files = {
        "PowerSystemResource": ["idx_psr", "FullName", "ShortName", "Extension"],
        "Terminal": ["idx_term", "idx_psr", "TerminalNumber", "MW", "Mvar", "u_mag", "u_deg"],
        "Load": ["idx_load", "idx_psr"],
        "Src": ["idx_src", "idx_psr"]
    }

    csv_files = dir.print_headers(csv_files, "latin-1")
    for csv_file in csv_files:
        print(csv_file)