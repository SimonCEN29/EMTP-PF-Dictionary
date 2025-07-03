import os
import user_tools
import powerfactory_api
import pf_tools
from pathlib import Path

if __name__ == '__main__':
    os.system('cls')

    
    pf_user = powerfactory_api.PfHandler("Kelly PF 2022")
    app = pf_user.start_powerfactory()
    pf_user.activate_project(r"\kelly.allendes\2505-BD-OP-COORD-DMAP")

    data = pf_tools.PfData(app)
    data.run_load_flow()

    csv_headers = {
        "PowerSystemResource": ["PsrID", "FullName", "ShortName", "Extension"],
        "Terminal": ["TermID", "PsrID", "TerminalSideID", "MW", "Mvar", "u_mag", "u_deg"],
        "Load": ["LoadID", "PsrID"],
        "Src": ["SrcID", "PsrID"]
    }

    dir = user_tools.Directory(r"C:\Users\kelly.allendes\Downloads")
    csv_files = dir.print_headers(csv_headers, "utf-8-sig")

    data.get_elms(csv_files, csv_headers, "utf-8-sig")