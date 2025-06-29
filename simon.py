import os
import user_tools

if __name__ == '__main__':
    os.system('cls')

    pf_user = user_tools.User("Simón")
    pf_user.prompt_path_selection(r"C:\Git Codes")
    pf_user.start_powerfactory()
    pf_user.activate_project(r"\simon.veloso\2505-BD-OP-COORD-DMAP")
