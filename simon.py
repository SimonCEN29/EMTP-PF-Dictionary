import os
import user_tools
import powerfactory_api
import pf_tools
from pathlib import Path

if __name__ == '__main__':
    os.system('cls')

    
    pf_user = powerfactory_api.PfHandler("Simón")
    app = pf_user.start_powerfactory()
    pf_user.activate_project(r"\simon.veloso\2505-BD-OP-COORD-DMAP")

    data = pf_tools.PfData(app)
    data.run_load_flow()

    csv_headers = {
        "PowerSystemResource": ["idx_psr", "FullName", "ShortName", "Extension"],
        "Terminal": ["idx_term", "idx_psr", "TerminalNumber", "MW", "Mvar", "u_mag", "u_deg"],
        "Load": ["idx_load", "idx_psr"],
        "Src": ["idx_src", "idx_psr"]
    }

    dir = user_tools.Directory(r"C:\Git Codes\EMTP-FP Dic Output")
    csv_files = dir.print_headers(csv_headers, "latin-1")

    data.get_elms(csv_files, csv_headers, "latin-1")