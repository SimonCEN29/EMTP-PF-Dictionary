import csv
import unicodedata

class PfData:
    def __init__(self, app):
        self.app = app

    def run_load_flow(self):
        ldf = self.app.GetFromStudyCase("ComLdf")
        ldf.Execute()

    def get_full_name(self, obj):
        ''' It gets the full name but only up to the "Network Data" folder '''
        full_name = obj.loc_name + "." + obj.GetClassName()
        parent = obj.fold_id

        while parent.loc_name != "Network Data":
            full_name = parent.loc_name + "\\" + full_name
            parent = parent.fold_id
        
        return full_name

    def get_elms(self, file_paths, file_headers, encoding="utf-8"):
        psr_csv_path = file_paths[0]
        term_csv_path = file_paths[1]
        load_csv_path = file_paths[2]
        src_csv_path = file_paths[3]

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

        PsrID, TermID, LoadID, SrcID = 1, 1, 1, 1

        with psr_csv_path.open("a", newline="", encoding=encoding) as f1:
            psr_writer = csv.writer(f1)
            
            for elm in elms:
                # print(elm.loc_name)
                class_name = elm.GetClassName()
                full_name = self.get_full_name(elm)
                short_name = elm.loc_name

                full_name = full_name.replace("\u2013", "-").replace("\u2014", "-")
                short_name = short_name.replace("\u2013", "-").replace("\u2014", "-")

                row = [PsrID, full_name, short_name, class_name]
                psr_writer.writerow(row)

                if class_name == "ElmLod":
                    with load_csv_path.open("a", newline="", encoding=encoding) as f2:
                        load_writer = csv.writer(f2)
                        row = [LoadID, PsrID]
                        load_writer.writerow(row)
                        LoadID += 1
                elif class_name in ("ElmSym", "ElmGenstat", "ElmSvs"):
                    with src_csv_path.open("a", newline="", encoding=encoding) as f2:
                        src_writer = csv.writer(f2)
                        row = [SrcID, PsrID]
                        src_writer.writerow(row)
                        SrcID += 1

                with term_csv_path.open("a", newline="", encoding=encoding) as f2:
                    term_writer = csv.writer(f2)

                    terms = self.get_terminals(elm)
                    for term in terms:
                        row = [
                            TermID, 
                            PsrID, 
                            term["TerminalSideID"], 
                            term["MW"], 
                            term["Mvar"],
                            term["u_mag"],
                            term["u_deg"]
                            ]
                        term_writer.writerow(row)
                        TermID += 1
                PsrID += 1

        return None
    
    def get_terminals(self, pf_object):
        terms = []
        class_name = pf_object.GetClassName()

        MW1, Mvar1, u_mag1, u_deg1 = "0", "0", "0", "0"
        MW2, Mvar2, u_mag2, u_deg2 = "0", "0", "0", "0"
        MW3, Mvar3, u_mag3, u_deg3 = "0", "0", "0", "0"
        MW4, Mvar4, u_mag4, u_deg4 = "0", "0", "0", "0"
        MW5, Mvar5, u_mag5, u_deg5 = "0", "0", "0", "0"

        if class_name in ("ElmLod", "ElmSym", "ElmGenstat", "ElmSvs"):
            try:
                MW1 = str(pf_object.GetAttribute("m:Psum:bus1"))
                Mvar1 = str(pf_object.GetAttribute("m:Qsum:bus1"))
                u_mag1 = str(pf_object.GetAttribute("m:u1:bus1"))
                u_deg1 = str(pf_object.GetAttribute("m:phiu1:bus1"))
            except AttributeError:
                pass

            terms = [{"TerminalSideID": "1", "MW": MW1, "Mvar": Mvar1, "u_mag": u_mag1, "u_deg": u_deg1}]

        elif class_name == "ElmLne":
            try:
                MW1 = str(pf_object.GetAttribute("m:Psum:bus1"))
                Mvar1 = str(pf_object.GetAttribute("m:Qsum:bus1"))
                u_mag1 = str(pf_object.GetAttribute("m:u1:bus1"))
                u_deg1 = str(pf_object.GetAttribute("m:phiu1:bus1"))

                MW2 = str(pf_object.GetAttribute("m:Psum:bus2"))
                Mvar2 = str(pf_object.GetAttribute("m:Qsum:bus2"))
                u_mag2 = str(pf_object.GetAttribute("m:u1:bus2"))
                u_deg2 = str(pf_object.GetAttribute("m:phiu1:bus2"))    
            except AttributeError:
                pass

            terms = [{"TerminalSideID": "1", "MW": MW1, "Mvar": Mvar1, "u_mag": u_mag1, "u_deg": u_deg1}]
            terms += [{"TerminalSideID": "2", "MW": MW2, "Mvar": Mvar2, "u_mag": u_mag2, "u_deg": u_deg2}]

        elif class_name in ("ElmTr2", "ElmTr3"):
            try:
                MW3 = str(pf_object.GetAttribute("m:Psum:bushv"))
                Mvar3 = str(pf_object.GetAttribute("m:Qsum:bushv"))
                u_mag3 = str(pf_object.GetAttribute("m:u1:bushv"))
                u_deg3 = str(pf_object.GetAttribute("m:phiu1:bushv"))

                MW5 = str(pf_object.GetAttribute("m:Psum:buslv"))
                Mvar5 = str(pf_object.GetAttribute("m:Qsum:buslv"))
                u_mag5 = str(pf_object.GetAttribute("m:u1:buslv"))
                u_deg5 = str(pf_object.GetAttribute("m:phiu1:buslv"))    
            except AttributeError:
                pass
            
            terms = [{"TerminalSideID": "3", "MW": MW3, "Mvar": Mvar3, "u_mag": u_mag3, "u_deg": u_deg3}]
            terms += [{"TerminalSideID": "5", "MW": MW5, "Mvar": Mvar5, "u_mag": u_mag5, "u_deg": u_deg5}]

            if class_name == "ElmTr3":
                try:
                    MW4 = str(pf_object.GetAttribute("m:Psum:busmv"))
                    Mvar4 = str(pf_object.GetAttribute("m:Qsum:busmv"))
                    u_mag4 = str(pf_object.GetAttribute("m:u1:busmv"))
                    u_deg4 = str(pf_object.GetAttribute("m:phiu1:busmv"))
                except AttributeError:
                    pass
                
                terms += [{"TerminalSideID": "4", "MW": MW4, "Mvar": Mvar4, "u_mag": u_mag4, "u_deg": u_deg4}]

        return terms
    

if __name__ == '__main__':
    pass