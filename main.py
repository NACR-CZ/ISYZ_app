import os
import os.path
import csv
import xml.etree.ElementTree as ET
import sys
import traceback

def get_exe_folder():
    if getattr(sys, "frozen", False):
        # path where the running exe is located
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

def exe_folder():
    try:
        exe_folder = get_exe_folder()
        export_folder = os.path.join(exe_folder, "export")
        os.makedirs(export_folder, exist_ok=True)
        return export_folder
    except Exception:
        traceback.print_exc()
    input("Press Enter to exit...")



# app_path = os.path.dirname(os.path.abspath(__file__))
# new_path = os.path.normpath(app_path)
# try:
#         os.makedirs(new_path + '/export/')
# except FileExistsError:
#     pass

# output_folder = new_path + '/export/'

def proarch_rejstrik_csv(xmlFile, file_name, output_folder):
    """Writes CSV file for ProArchiv software"""
    output_file = file_name.strip('.xml')
    with open(output_folder + "/proArch_rejstrik_" + output_file + ".csv",
             "w",
             newline="",
             encoding="UTF-8",
    ) as csvfile:
      writer = csv.writer(csvfile, delimiter="|", quoting=csv.QUOTE_ALL)
      writer.writerow(
       [
        "Úroveň popisu",
        "1. úroveň série",
        "Kategorie",
        "Evidenční jednotka Typ",
        "Evidenční jednotka Číslo",
        "Ukládací jednotka Typ",
        "Ukládací jednotka Číslo",
        "Původní/jiné označení Typ1",
        "Původní/jiné označení Hodnota1",
        "Původní/jiné označení Typ2",
        "Původní/jiné označení Hodnota2",
        "Obsah",
        "Datace vzniku",
        "Fyzický stav",
        "Možnost zveřejnění",
        "Přístupnost",
        "Služební poznámka",
        "Rozhodnutí archivu"
       ]
      )
      writer.writerow(
       ["série", "ano", "", "", "", "", "", "", "", "", "", "", "", "Rejstřík", "", ""]
      )
      csvfile.close()
    xml = ET.parse(xmlFile)
    for row in xml.findall("vec"):
        # Iterate over the XML file
        datace_raw = list()
        obsah = list()
        osoby = list()
        zprac = row.find("druh_stav_veci").text
        if "ODSKRTNUTA" == str(zprac):
            sp_zn_list = list()
            druh = row.find("druh_vec").text
            rocnik = row.find("rocnik").text
            bc_vec = row.find("bc_vec").text
            sp_zn_list.append(bc_vec)
            sp_zn_list.append(rocnik)
            predmet_rizeni = row.find("predmet_rizeni").text
            obsah.append(predmet_rizeni)
            raw_druh_vysledek = row.find("druh_vysledek").text
            druh_vysledek = list()
            if raw_druh_vysledek is not None:
                druh_vysledek.append(raw_druh_vysledek.lower())
            else:
                druh_vysledek.append('Nevyplněno')
            obsah.append('Vyřízení: ' + " ".join(druh_vysledek))
            # id_osoby_vyridil = row.find("id_osoby_vyridil").text
            try:
                senat_row = row.find("cislo_senatu").text
                if senat_row and int(senat_row) >= 0:
                    senat = senat_row
                    sp_zn_list.append(senat)
                else:
                    pass
            except AttributeError:
                pass
            sp_zn = "/".join(sp_zn_list)
            datum_doslo = row.find("datum_doslo").text
            datum_od = datum_doslo[0:4]
            datum_odskrtnuti = row.find("datum_odskrtnuti").text
            datum_do = datum_odskrtnuti[0:4]
            if datum_od == datum_do:
                datace_raw.append(datum_od)
            else:
                datace_raw.append(datum_od + "-" + datum_do)
            persons = row.findall("data_o_osobe_v_rizeni")
            for person in persons:
                osoba = list()
                role = person.find("druh_role_v_rizeni").text
                if role is not None:
                    osoba.append(role)
                else:
                    osoba.append('Role osoby nevyplněna')
                prijmeni = person.find("nazev_osoby").text
                if prijmeni is not None:
                    osoba.append(prijmeni)
                else:
                    pass
                jmeno = person.find("jmeno").text
                if jmeno is not None:
                    osoba.append(jmeno)
                else:
                    pass
                vek = person.find("priznak_mladistvy_dospely").text
                if vek == "M":
                    osoba.append("Mladistvý")
                else:
                    pass
                umrti = person.find("datum_umrti").text
                if umrti is not None:
                    osoba.append("Datum úmrtí: + " + umrti)
                else:
                    pass
                osoba_full = ", ".join(osoba)
                osoby.append(osoba_full)
        elif "MYLNÝ" in zprac:
            obsah = list()
            druh = row.find("druh_vec").text
            rocnik = row.find("rocnik").text
            bc_vec = row.find("bc_vec").text
            sp_zn_roc_list = (bc_vec, rocnik)
            sp_zn_roc = "/".join(sp_zn_roc_list)
            sp_zn_list = (druh, sp_zn_roc)
            sp_zn = " ".join(sp_zn_list)
            obsah.append("Mylný zápis")
            datace_raw.append("")
        elif "NEUKONČENO" in zprac:
            druh = row.find("druh_vec").text
            rocnik = row.find("rocnik").text
            bc_vec = row.find("bc_vec").text
            sp_zn_roc_list = (bc_vec, rocnik)
            sp_zn_roc = "/".join(sp_zn_roc_list)
            sp_zn_list = (druh, sp_zn_roc)
            sp_zn = " ".join(sp_zn_list)
            obsah.append("Neukončený spis")
            datace_raw.append("")
            # print(rocnik, bc_vec, sp_zn, obsah)
        else:
            druh = row.find("druh_vec").text
            rocnik = row.find("rocnik").text
            bc_vec = row.find("bc_vec").text
            sp_zn_roc_list = (bc_vec, rocnik)
            sp_zn_roc = "/".join(sp_zn_roc_list)
            sp_zn_list = (druh, sp_zn_roc)
            sp_zn = "/".join(sp_zn_list)
            datace_raw.append("")
            obsah = "chybný vstup pro spisovou značku"
            # print(rocnik, bc_vec, sp_zn, obsah)
        obsah_raw = " / ".join(obsah)
        # print(obsah_raw)
        obsah_final = obsah_raw.replace(' / ', '; ')
        # print(obsah_final)
        datace = "-".join(datace_raw)
        osoby_str = "; ".join(osoby)
        csv_line = [
            "složka",
            "",
            """Všeobecné, spisy""",
            "karton",
            "",
            "karton",
            "",
            "Spisová značka",
            sp_zn,
            "",
            "",
            obsah_final,
            datace,
            "dobrý",
            "nepublikovat",
            "osobní údaje - omezená přístupnost",
            osoby_str,
            ""]
        with open(output_folder + "/proArch_rejstrik_" + output_file + ".csv", "a",
                  newline="",
                  encoding="UTF-8") as csvfile:
                writer = csv.writer(csvfile,
                                    delimiter="|",
                                    quoting=csv.QUOTE_ALL)
                writer.writerow(csv_line)
                csvfile.close()


def proarch_isyz_csv(xmlFile, file_name, output_folder):
    """Produces CSV from parsed XML file"""
    output_file = file_name.strip('.xml')
    with open(output_folder + "/proArch_isyz_" + output_file + ".csv",
             "w",
             newline="",
             encoding="UTF-8",
    ) as csvfile:
      writer = csv.writer(csvfile,
                          delimiter="|",
                          quoting=csv.QUOTE_ALL)
      writer.writerow(
       [
        "Úroveň popisu",
        "1. úroveň série",
        "Kategorie",
        "Evidenční jednotka Typ",
        "Evidenční jednotka Číslo",
        "Ukládací jednotka Typ",
        "Ukládací jednotka Číslo",
        "Původní/jiné označení Typ1",
        "Původní/jiné označení Hodnota1",
        "Původní/jiné označení Typ2",
        "Původní/jiné označení Hodnota2",
        "Obsah",
        "Datace vzniku",
        "Fyzický stav",
        "Možnost zveřejnění",
        "Přístupnost",
        "Služební poznámka",
        "Rozhodnutí archivu"

       ]
      )
      writer.writerow(
       ["série", "ano", "", "", "", "", "", "", "", "", "", "", "", "ISYZ", "", ""]
      )
      csvfile.close()
    xml = ET.parse(xmlFile)

    for row in xml.findall("vec"):
        # Iterate over the XML file
        osoby = list()
        sp_zn_list = list()
        cislo_rejstriku = row.find("cislo_rejstrik").text
        druh = row.find("druh_vec").text
        rocnik = row.find("rocnik").text
        bc_vec = row.find("bc_vec").text
        sp_zn_roc_list = (bc_vec, rocnik)
        sp_zn_roc = "/".join(sp_zn_roc_list)
        sp_zn_list.append(cislo_rejstriku)
        sp_zn_list.append(druh)
        sp_zn_list.append(sp_zn_roc)
        try:
            senat_row = row.find("cislo_senatu").text
            if senat_row and int(senat_row) >= 0:
                senat = senat_row
                sp_zn_list.append(senat)
            else:
                pass
        except AttributeError:
            pass
        sp_zn = " ".join(sp_zn_list)
        datum_od_raw = row.find("datum_doslo").text
        datum_od = datum_od_raw[0:4]
        datum_do_raw = row.find("datum_vyrizeni").text
        datum_do = datum_do_raw[0:4]
        datace_raw = list()
        if datum_od == datum_do:
            datace_raw.append(datum_od)
        else:
            datace_raw.append(datum_od + "-" + datum_do)
        # c_rejstrik = row.find("cislo_rejstrik").text
        # prezkoumani = row.find("priznak_an_prezkoumani_zakonn").text
        # skartace = row.find("vyloucen_ze_skartace").text
        try:
            cizi_organizace_raw = row.find("cizi_organizace").text
            if cizi_organizace_raw is None:
                cizi_oganizace = ""
            else:
                cizi_oganizace = cizi_organizace_raw
        except AttributeError:
            cizi_oganizace = ""
        try:
            cizi_spisova_znacka = row.find("cizi_spisova_znacka").text
            if cizi_spisova_znacka is None:
                cj_cizi = ""
                cizi_spisova_znacka = ""
            else:
                cizi_spisova_znacka = ""
                # cj_cizi = "Spisová značka"
                cj_cizi = ""
        except AttributeError:
            cj_cizi = ""
            cizi_spisova_znacka = ""
        obsah_raw = row.find("popis_predmet_rizeni")
        if obsah_raw.text is not None:
            popis = str(obsah_raw)
        else:
            popis = "N/A"
        persons = row.findall("osoba_v_rizeni")
        for person in persons:
            osoba = list()
            try:
                prijmeni = person.find("nazev_osoby").text
                if prijmeni is not None:
                    osoba.append(prijmeni)
                else:
                    pass
            except AttributeError:
                pass
            try:
                jmeno = person.find("jmeno_osoby").text
                if jmeno is not None:
                    osoba.append(jmeno)
                else:
                    pass
            except AttributeError:
                pass
            try:
                datum_narozeni = person.find("datum_narozeni").text
                if datum_narozeni is not None:
                    osoba.append(datum_narozeni)
                else:
                    pass
            except AttributeError:
                pass
            try:
                vek = person.find("priznak_mladistvy_dospely").text
                if vek == "M":
                    osoba.append("Mladistvý")
                else:
                    pass
            except AttributeError:
                pass
            try:
                umrti = person.find("datum_umrti").text
                if umrti is not None:
                    osoba.append("Datum úmrtí: + " + umrti)
                else:
                    pass
            except AttributeError:
                pass
            osoba_full = ", ".join(osoba)
            osoby.append(osoba_full)
        vysledek = row.find("druh_vysledek").text
        obsah_list = (cizi_oganizace, popis, vysledek)
        osoby_str = "; ".join(osoby)
        obsah = "; ".join(obsah_list)
        datace = "-".join(datace_raw)
        csv_line = [
                "složka",
                "",
                """Všeobecné, spisy""",
                "karton",
                "",
                "karton",
                "",
                "Spisová značka",
                sp_zn,
                cj_cizi,
                cizi_spisova_znacka,
                obsah,
                datace,
                "dobrý",
                "nepublikovat",
                "osobní údaje - omezená přístupnost",
                osoby_str,
                ""
        ]
        with open(output_folder + "/proArch_isyz_" + output_file + ".csv",
                      "a",
                      newline="",
                      encoding="UTF-8") as csvfile:
            writer = csv.writer(csvfile, delimiter="|", quoting=csv.QUOTE_ALL)
            writer.writerow(csv_line)
            csvfile.close()


def elza_rejstrik_csv(xmlFile, file_name, output_folder):
    output_file = file_name.strip('.xml')
    with open(
            output_folder + "/elza_rejstrik_" + output_file + ".csv",
            "w",
            newline="",
            encoding="cp1250",
    ) as csvfile:
        writer = csv.writer(csvfile, delimiter=";", quoting=csv.QUOTE_ALL)
        writer.writerow(  # output_folder +
            [
                "Signatura původce",
                "Regest",
                "Role entit",
                "Úroveň",
                "Datace",
                "Typ úrovně popisu",
                "Způsob uložení",
                "Podmínky přístupu",
                "Rozhodnutí archivu"
                # "Úroveň popisu",
                # "1. úroveň série",
                # "Kategorie",
                # "Evidenční jednotka Typ",
                # "Evidenční jednotka Číslo",
                # "Ukládací jednotka Typ",
                # "Ukládací jednotka Číslo",
                # "Původní/jiné označení Typ1",
                # "Původní/jiné označení Hodnota1",
                # "Původní/jiné označení Typ2",
                # "Původní/jiné označení Hodnota2",
                # "Obsah",
                # "Datace vzniku",
                # "Možnost zveřejnění",
                # "Přístupnost",
            ]
        )
        writer.writerow(
            [
                "CE:ZP2015_OTHER_ID;ZP2015_OTHERID_DOCID",
                "CE:ZP2015_TITLE",
                "CE:ZP2015_ENTITY_ROLE_TEXT;ZP2015_ENTITY_ROLE_90",
                "UROVEN",
                "DATACE_TEXT",
                "UROVEN_TYP",
                "UKL_TYP",
                "CE:ZP2015_UNIT_ACCESS",
                ""
            ]
        )
        writer.writerow(
            ["", "Série", "", "1", "", "SERIE", "", "", "", ""]
        )
        csvfile.close()
        xml = ET.parse(xmlFile)
        for row in xml.findall("vec"):
            # Iterate over the XML file
            datace_raw = list()
            zprac = row.find("druh_stav_veci").text
            osoby = list()
            if "ODSKRTNUTA" == str(zprac):
                sp_zn_list = list()
                druh = row.find("druh_vec").text
                rocnik = row.find("rocnik").text
                bc_vec = row.find("bc_vec").text
                sp_zn_list.append(bc_vec)
                sp_zn_list.append(rocnik)
                # sp_zn_roc_list = (bc_vec, rocnik)
                # sp_zn_roc = "/".join(sp_zn_roc_list)
                obsah = list()
                predmet_rizeni = row.find("predmet_rizeni").text
                obsah.append(predmet_rizeni)
                raw_druh_vysledek = row.find("druh_vysledek").text
                druh_vysledek = list()
                if raw_druh_vysledek is not None:
                    druh_vysledek.append(raw_druh_vysledek.lower())
                else:
                    druh_vysledek.append('Nevyplněno')
                obsah.append('Vyřízení: ' + " ".join(druh_vysledek))
                # id_osoby_vyridil = row.find("id_osoby_vyridil").text
                senat_row = row.find("cislo_senatu").text
                if senat_row and int(senat_row) >= 0:
                    senat = senat_row
                    sp_zn_list.append(senat)
                else:
                    sp_zn_list = (druh, sp_zn_roc)
                sp_zn = "/".join(sp_zn_list)
                datum_doslo = row.find("datum_doslo").text
                datum_od = datum_doslo[0:4]
                datum_odskrtnuti = row.find("datum_odskrtnuti").text
                datum_do = datum_odskrtnuti[0:4]
                if datum_od == datum_do:
                    datace_raw.append(datum_od)
                else:
                    datace_raw.append(datum_od + "-" + datum_do)
                persons = row.findall("data_o_osobe_v_rizeni")
                for person in persons:
                    osoba = list()
                    role = person.find("druh_role_v_rizeni").text
                    if role is not None:
                        osoba.append(role)
                    else:
                        osoba.append('Role osoby nevyplněna')
                    prijmeni = person.find("nazev_osoby").text
                    if prijmeni is not None:
                        osoba.append(prijmeni)
                    else:
                        pass
                    jmeno = person.find("jmeno").text
                    if jmeno is not None:
                        osoba.append(jmeno)
                    else:
                        pass
                    try:
                        datum_narozeni = person.find("datum_narozeni").text
                        if datum_narozeni is not None:
                            osoba.append(datum_narozeni)
                        else:
                            pass
                    except AttributeError:
                        pass
                    try:
                        vek = person.find("priznak_mladistvy_dospely").text
                        if vek == "M":
                            osoba.append("Mladistvý")
                        else:
                            pass
                    except AttributeError:
                        pass
                    try:
                        umrti = person.find("datum_umrti").text
                        if umrti is not None:
                            osoba.append("Datum úmrtí: + " + umrti)
                        else:
                            pass
                    except AttributeError:
                        pass
                    osoba_full = ", ".join(osoba)
                    osoby.append(osoba_full)
            elif "MYLNÝ" in zprac:
                obsah = list()
                druh = row.find("druh_vec").text
                rocnik = row.find("rocnik").text
                bc_vec = row.find("bc_vec").text
                sp_zn_roc_list = (bc_vec, rocnik)
                sp_zn_roc = "/".join(sp_zn_roc_list)
                sp_zn_list = (druh, sp_zn_roc)
                sp_zn = " ".join(sp_zn_list)
                obsah.append("Mylný zápis")
                datace_raw.append("")
            elif "NEUKONČENO" in zprac:
                druh = row.find("druh_vec").text
                rocnik = row.find("rocnik").text
                bc_vec = row.find("bc_vec").text
                sp_zn_roc_list = (bc_vec, rocnik)
                sp_zn_roc = "/".join(sp_zn_roc_list)
                sp_zn_list = (druh, sp_zn_roc)
                sp_zn = " ".join(sp_zn_list)
                obsah.append("Neukončený spis, převedeno")
                datace_raw.append("")
                # print(rocnik, bc_vec, sp_zn, obsah)
            else:
                druh = row.find("druh_vec").text
                rocnik = row.find("rocnik").text
                bc_vec = row.find("bc_vec").text
                sp_zn_roc_list = (bc_vec, rocnik)
                sp_zn_roc = "/".join(sp_zn_roc_list)
                sp_zn_list = (druh, sp_zn_roc)
                sp_zn = " ".join(sp_zn_list)
                datace_raw.append("")
                obsah = "chybný vstup pro spisovou značku"
                # print(rocnik, bc_vec, sp_zn, obsah)
            obsah_raw = " | ".join(obsah)
            obsah_final = obsah_raw.replace(' | ', '; ')
            datace = "-".join(datace_raw)
            osoby_str = "; ".join(osoby)
            csv_line = [
                sp_zn,
                obsah_final,
                osoby_str,
                "2",
                datace,
                "SLOZKA-MEJ",
                "KAR",
                "Nelze zpřístupnit z důvodu ochrany osobních a citlivých údajů.",
                ""
            ]
            # print(csv_line)
            with open(
                    output_folder + "/elza_rejstrik_" + output_file + ".csv",
                    "a",
                    newline="",
                    encoding="cp1250",
            ) as csvfile:
                writer = csv.writer(csvfile, delimiter=";", quoting=csv.QUOTE_ALL)
                writer.writerow(csv_line)
                csvfile.close()


def elza_isyz_csv(xmlFile, file_name, output_folder):
    """Writes CSV file for Elza software"""
    output_file = file_name.strip('.xml')
    with open(
            output_folder + "/elza_isyz_" + output_file + ".csv",
            "w",
            newline="",
            encoding="cp1250",
    ) as csvfile:
        writer = csv.writer(csvfile, delimiter=";", quoting=csv.QUOTE_ALL)
        writer.writerow(  # output_folder +
            [
                "Signatura původce",
                "Regest",
                "Role entit",
                "Úroveň",
                "Datace",
                "Typ úrovně popisu",
                "Způsob uložení",
                "Podmínky přístupu",
                "Rozhodnutí archivu"
                # "Úroveň popisu",
                # "1. úroveň série",
                # "Kategorie",
                # "Evidenční jednotka Typ",
                # "Evidenční jednotka Číslo",
                # "Ukládací jednotka Typ",
                # "Ukládací jednotka Číslo",
                # "Původní/jiné označení Typ1",
                # "Původní/jiné označení Hodnota1",
                # "Původní/jiné označení Typ2",
                # "Původní/jiné označení Hodnota2",
                # "Obsah",
                # "Datace vzniku",
                # "Možnost zveřejnění",
                # "Přístupnost",
            ]
        )
        writer.writerow(
            [
                "CE:ZP2015_OTHER_ID;ZP2015_OTHERID_DOCID",
                "CE:ZP2015_TITLE",
                "CE:ZP2015_ENTITY_ROLE_TEXT;ZP2015_ENTITY_ROLE_90",
                "UROVEN",
                "DATACE_TEXT",
                "UROVEN_TYP",
                "UKL_TYP",
                "CE:ZP2015_UNIT_ACCESS",
                ""
            ]
        )
        writer.writerow(
            ["", "Série", "", "1", "", "SERIE", "", "", "", ""]
        )
        csvfile.close()
        xml = ET.parse(xmlFile)

        for row in xml.findall("vec"):
            # Iterate over the XML file
            osoby = list()
            sp_zn_list = list()
            cislo_rejstriku = row.find("cislo_rejstrik").text
            druh = row.find("druh_vec").text
            rocnik = row.find("rocnik").text
            bc_vec = row.find("bc_vec").text
            sp_zn_roc_list = (bc_vec, rocnik)
            sp_zn_roc = "/".join(sp_zn_roc_list)
            sp_zn_list.append(cislo_rejstriku)
            sp_zn_list.append(druh)
            sp_zn_list.append(sp_zn_roc)
            try:
                senat_row = row.find("cislo_senatu").text
                if senat_row and int(senat_row) > 0:
                    senat = senat_row
                    sp_zn_list.append(senat)
                else:
                    pass
            except AttributeError:
                pass
            sp_zn = " ".join(sp_zn_list)
            datum_od_raw = row.find("datum_doslo").text
            datum_od = datum_od_raw[0:4]
            datum_do_raw = row.find("datum_vyrizeni").text
            datum_do = datum_do_raw[0:4]
            if datum_od == datum_do:
                datace = str(datum_od)
            else:
                datace = str(datum_od + "-" + datum_do)
            # c_rejstrik = row.find("cislo_rejstrik").text
            # prezkoumani = row.find("priznak_an_prezkoumani_zakonn").text
            # skartace = row.find("vyloucen_ze_skartace").text
            try:
                cizi_organizace_raw = row.find("cizi_organizace").text
                if cizi_organizace_raw is None:
                    cizi_oganizace = ""
                else:
                    cizi_oganizace = cizi_organizace_raw
            except AttributeError:
                cizi_oganizace = ""
            # cizi_spisova_znacka_raw = row.find("cizi_spisova_znacka").text
            # if cizi_spisova_znacka_raw is None:
            #     cj_cizi = ""
            #     cizi_spisova_znacka = ""
            # else:
            #     cizi_spisova_znacka = cizi_spisova_znacka_raw
            #     # cj_cizi = "Spisová značka"
            #     cj_cizi = ""
            obsah_raw = row.find("popis_predmet_rizeni")
            if obsah_raw.text is not None:
                popis = str(obsah_raw)
            else:
                popis = "Nevyplněno"
            persons = row.findall("osoba_v_rizeni")
            for person in persons:
                osoba = list()
                try:
                    prijmeni = person.find("nazev_osoby").text
                    if prijmeni is not None:
                        osoba.append(prijmeni)
                    else:
                        pass
                except AttributeError:
                    pass
                try:
                    jmeno = person.find("jmeno_osoby").text
                    if jmeno is not None:
                        osoba.append(jmeno)
                    else:
                        pass
                except AttributeError:
                    pass
                try:
                    datum_narozeni = person.find("datum_narozeni").text
                    if datum_narozeni is not None:
                        osoba.append(datum_narozeni)
                    else:
                        pass
                except AttributeError:
                    pass
                try:
                    vek = person.find("priznak_mladistvy_dospely").text
                    if vek == "M":
                        osoba.append("Mladistvý")
                    else:
                        pass
                except AttributeError:
                    pass
                try:
                    umrti = person.find("datum_umrti").text
                    if umrti is not None:
                        osoba.append("Datum úmrtí: + " + umrti)
                    else:
                        pass
                except AttributeError:
                    pass
                osoba_full = ", ".join(osoba)
                osoby.append(osoba_full)
            vysledek = row.find("druh_vysledek").text
            obsah_list = (
                cizi_oganizace,
                popis,
                vysledek)
            osoby_str = "; ".join(osoby)
            obsah = "; ".join(obsah_list)
            csv_line = [
                sp_zn,
                obsah,
                osoby_str,
                "2",
                datace,
                "SLOZKA-MEJ",
                "KAR",
                "Nelze zpřístupnit z důvodu ochrany osobních a citlivých údajů."
                
                ""
            ]
            # print(csv_line)

            with open(
                    output_folder + "/elza_isyz_" + output_file + ".csv",
                    "a",
                    newline="",
                    encoding="cp1250",
            ) as csvfile:
                writer = csv.writer(csvfile, delimiter=";", quoting=csv.QUOTE_ALL)
                writer.writerow(csv_line)
                csvfile.close()


def file_process():
    export_folder = exe_folder()
    # print(export_folder)
    import_folder = str(export_folder).strip('/export')
    # print(import_folder)
    print('Národní Archiv\nisyz_soudy_appka spuštěna, vyhledávám XML data pro zpracování\n')
    rejstrik_files = 0
    isyz_files = 0
    xml_list = list()
    for dirpath, dirnames, filenames in os.walk(import_folder):
        for file_name in [f for f in filenames if f. endswith(".xml") or f. endswith(".XML")]:
            file_content = open(os.path.join(dirpath, file_name), encoding="utf8").read()
            file_path = os.path.join(dirpath, file_name)
            for line in file_content.splitlines():

                #Check for isyz
                if (line.find("isyz") or line.find("ZASTUPITELSTVÍ") or line.find("zastupitelství")
                            or line.find("Zastupitelství")) != -1 :
                    print('Rozpoznány XML data z ISYZ: ' + file_name)
                    xml_list.append(file_name)
                    isyz_files += 1
                    print('Zpracovávám vstupní data pro ProArch...')
                    proarch_isyz_csv(file_path, file_name, export_folder)
                    print('Zpracovávám vstupní data pro Elza...\n')
                    elza_isyz_csv(file_path, file_name, export_folder)
                    break
                elif (line.find("soud") or line.find("SOUD") or line.find("Soud")) != -1 :
                    print('Rozpoznány XML data z rejstříků soudů: ' + file_name)
                    xml_list.append(file_name)
                    rejstrik_files += 1
                    print('Zpracovávám vstupní data rejstříku pro ProArch...')
                    proarch_rejstrik_csv(file_path, file_name, export_folder)
                    print('Zpracovávám vstupní rejstříku data pro Elza...\n')
                    elza_rejstrik_csv(file_path, file_name, export_folder)
                    break
    with open("isyz_soudy_appka_zprava.html",
              "w",
            newline="\n",
            encoding="UTF-8") as report:
        report.write("<html><body><h1>Národní archiv, isyz_soudy_appka</h1><br><p>Zpracováno:<p><br>")
        report.close()
    with open("isyz_soudy_appka_zprava.html",
              "a",
            newline="\n",
            encoding="UTF-8") as report:
        for i in xml_list:
            report.write(i+'<br>')
        report.close()
    with open("isyz_soudy_appka_zprava.html",
              "a",
              newline="\n",
              encoding="UTF-8") as report:
        report.write('</p></body></html>')
        report.close()
    print('\nZpracováno celkem ' + str(rejstrik_files) + ' XML z rejstříků, ' + str(isyz_files) +
          ' XML z ISYZu. Report o průběhu je vytvořen, ukončuji zpracování')


if __name__ == "__main__":
    # export_folder = exe_folder()
    file_process()

