import sys
from api_paths import api_paths

class PfHandler:
    def __init__(self, user_name=None):
        self.api_paths = api_paths
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
        app = []
        try:
            app = pf.GetApplicationExt()
        except pf.ExitError as error:
            print(error)
            print("error.code = " + str(error.code))
            sys.exit(0)
        print("DIgSILENT PowerFactory started correctly.\n")

        self.app = app
        return app

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

if __name__ == '__main__':
    pass