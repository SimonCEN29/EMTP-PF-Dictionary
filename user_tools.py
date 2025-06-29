import tkinter
from tkinter import filedialog
import sys
import os
import csv
from pathlib import Path

class User:
    api_paths = {
        "Simón": r"C:\Program Files\DIgSILENT\PowerFactory 2022 SP9\Python\3.10"
    }

    def __init__(
        self, user_name=None
    ):
        self.user_name = user_name
        self.saving_path = None

        loop_condition = True
        while loop_condition:
            if user_name:
                try:
                    self.api_path = self.api_paths[user_name]
                    self.user_name = user_name
                    print('\nUser name selected: "' + self.user_name + '".')
                    print('API path at: "' + self.api_path + '".\n')
                    loop_condition = False

                except KeyError:
                    print("\nAn invalid user name was entered.\n")
                    user_name = self.prompt_user_selection()
            else:
                print("\nA user name was not entered.\n")
                user_name = self.prompt_user_selection()

    def prompt_user_selection(self):
        for key in self.api_paths.keys():
            print(
                'User name: "' + key + '", API path at: "' + self.api_paths[key] + '".'
            )

        user_name = input("\nEnter a valid user name from the list above: ")

        return user_name

    def start_powerfactory(self):
        sys.path.append(self.api_path)

        import powerfactory as pf

        print("Starting DIgSILENT PowerFactory...")
        self.app = []
        try:
            self.app = pf.GetApplicationExt()
        except pf.ExitError as error:
            print(error)
            print("error.code = " + str(error.code))
            sys.exit(0)
        print("DIgSILENT PowerFactory started correctly.\n")

        return None

    def prompt_path_selection(self, saving_path=None):
        """If this method is given a path as argument, it will skip the window prompt.
        It is suggested the user calls this method without a path argument, so an error
        is avoided for assining an invalid path."""

        root = (
            tkinter.Tk()
        )  # To manually close tkinter window afther this methods is finished.

        self.saving_path = saving_path

        loop_condition = True
        while loop_condition:
            if self.saving_path:
                loop_condition = False
            else:
                self.saving_path = tkinter.filedialog.askdirectory()
                self.saving_path = self.saving_path.replace(
                    "/", "\\"
                )  # askdirectory() uses slash instead of backslash
                if self.saving_path == "":
                    print("No folder was selected. Please select a folder.\n")

        print('Files will be saved at: "' + self.saving_path + '".')

        root.withdraw()  # Manually closing tkinter window.

        return None

    def activate_project(self, project_pf_path=None):
        projects = self.app.GetCurrentUser().GetContents("*.IntPrj")

        loop_condition = True
        while loop_condition:
            if project_pf_path:
                err = self.app.ActivateProject(project_pf_path)
            else:
                idx_project = 0
                for project in projects:
                    print(str(idx_project) + ". " + project.GetAttribute("loc_name"))
                    idx_project += 1

                idx_project = int(
                    input("\nEnter index number of project to activate: ")
                )
                print(
                    'Activating project "'
                    + projects[idx_project].GetAttribute("loc_name")
                    + '"...'
                )

                try:
                    err = self.app.ActivateProject(projects[idx_project].GetFullName())
                except IndexError:
                    print("Index entered out of index list.")

            if err == 0:
                try:
                    print(
                        'Project "'
                        + projects[idx_project].GetAttribute("loc_name")
                        + '" activated.\n'
                    )
                except UnboundLocalError:
                    print('Project "' + project_pf_path + '" activated.\n')
                loop_condition = False

        return None

    def get_name(self, obj):
        full_name = obj.loc_name + "." + obj.GetClassName()
        parent = obj.fold_id

        while parent.loc_name != "Network Data":
            full_name = parent.loc_name + "\\" + full_name
            parent = parent.fold_id
        
        return full_name

    def get_elms(self, psr_file_name, term_file_name, encoding="utf-8"):
        psr_csv_path = Path(self.saving_path + "\\" + psr_file_name)
        term_csv_path = Path(self.saving_path + "\\" + term_file_name)

        network_data = self.app.GetProjectFolder("netdat")
        loads = network_data.GetContents('*.ElmLod', 1)
        tlines = network_data.GetContents('*.ElmLne', 1)
        tr2s = network_data.GetContents('*.ElmTr2', 1)
        tr3s = network_data.GetContents('*.ElmTr3', 1)
        gens = network_data.GetContents('*.ElmSym', 1)
        ibrs = network_data.GetContents('*.ElmGenstat', 1)
        cers = network_data.GetContents('*.ElmSvs', 1)

        elms = loads
        elms.extend(tlines)
        elms.extend(tr2s)
        elms.extend(tr3s)
        elms.extend(gens)
        elms.extend(ibrs)
        elms.extend(cers)

        with psr_csv_path.open("a", newline="", encoding=encoding) as f1:
            psr_writer = csv.writer(f1)
            idx_elm = 0
            idx_term = 0
            
            for elm in elms:
                class_name = elm.GetClassName()
                full_name = self.get_name(elm)
                row = [idx_elm, full_name, elm.loc_name, class_name]
                psr_writer.writerow(row)

                with term_csv_path.open("a", newline="", encoding=encoding) as f2:
                    term_writer = csv.writer(f2)

                    terms = self.get_terminals(elm)
                    for term in terms:
                        row = [idx_term, idx_elm, term["TerminalNumber"], term["MW"], term["Mvar"]]
                        term_writer.writerow(row)
                        idx_term += 1
                idx_elm += 1

        return None
    
    def print_header(self, file_name, header, encoding="utf-8"):
        csv_path = Path(self.saving_path + "\\" + file_name)
        
        with csv_path.open("w", newline="", encoding=encoding) as f:
            writer = csv.writer(f)
            writer.writerow(header)

    def get_terminals(self, pf_object):
        terms = []
        class_name = pf_object.GetClassName()

        if class_name == "ElmLod":
            try:
                MW = str(pf_object.GetAttribute("m:P:bus1"))
                Mvar = str(pf_object.GetAttribute("m:Q:bus1"))
                u_mag = str(pf_object.GetAttribute("n:u:bus1"))
                u_deg = str(pf_object.GetAttribute("n:phiu:bus1"))
                terms = [{"TerminalNumber": "bus1", "MW": MW, "Mvar": Mvar, "u_mag": u_mag, "u_deg": u_deg}]
            except AttributeError:
                pass

        return terms
    
    def run_load_flow(self):
        ldf = self.app.GetFromStudyCase("ComLdf")
        ldf.Execute()

if __name__ == '__main__':
    os.system('cls')

    pf_user = User("Simón")
    pf_user.start_powerfactory()
    pf_user.activate_project(r"\simon.veloso\2505-BD-OP-COORD-DMAP")
    pf_user.run_load_flow()
    pf_user.prompt_path_selection(r"C:\Git Codes")

    psr_file_name = "PowerSystemResource.csv"
    term_file_name = "Terminal.csv"

    pf_user.print_header(psr_file_name, ["idx_elm", "FullName", "ShortName", "Extension"])
    pf_user.print_header(term_file_name, ["idx_term", "idx_elm", "TerminalNumber", "MW", "Mvar", "u_mag", "u_deg"])
    pf_user.get_elms(psr_file_name, term_file_name)