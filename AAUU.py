from contextlib import nullcontext
import pandas as pandas
from tkinter import Tk
from tkinter.filedialog import askdirectory
import glob
import os

# numbers indicate x-values, the first row (y = 0) is skipped by pandas, so (0,0) is a true value and not a label.
opmerking_sales_ = 7
tenant_ = 6
technissche_contactpersoon_ = 4
school_ = 2

bedrijfsnaam_ = 1
Primaire_domeinaam_ = 2
empty = "nan"

# SETTINGS:
exclude_array = ["OK", "nvt", "NVT", "A3"]


def main():
    mode = int(input("1. All\n2. Selective\n3. Autopilot" + "\n"))
    print("Selecting...")
    Tk().withdraw()  # use to hide tkinter window
    Dir = askdirectory(title="where are the files?") + "/"  # shows dialog box and return the path.

    aau_excel, customers_csv, autopilot = files(Dir)

    LENGTE_AAU, LENGTE_PARTNER, LENGTE_AUTOPILOT = len(aau_excel), len(customers_csv), len(autopilot)

    print("reading...")
    if mode == 1:
        pos_array = all_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv)
    elif mode == 2:
        pos_array = pos_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv)
    elif mode == 3:
        pos_array = func_autopilot(LENGTE_AUTOPILOT, LENGTE_PARTNER, autopilot, customers_csv)
    else:
        print("you have to select between teh 3 modes, closing...")
        return

    writer(Dir, pos_array)

def writer(Dir, pos_array):

    print("writing...")
    Dir = Dir + "Results AAUU.txt"
    results = open(Dir, "w", encoding="utf-8")
    results.write("In lijst:\n")
    for line in pos_array:
        results.write(f"{line}\n")

    os.startfile(Dir)
    print("done")


def files(Dir):
    try:
        autopilot_file_name = glob.glob(Dir + "*autopilot*.xlsx")
        autopilot_file_name.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        autopilot = autopilot_file_name[0]
        autopilot = pandas.read_excel(autopilot)
    except:
        autopilot = []
        print('no "*autopilot*.xlsx" detected')

    try:
        AAU_file_name = glob.glob(Dir + "*AAU*.xlsx")
        AAU_file_name.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        AAU_file_first = AAU_file_name[0]  # convert from array to a string
        aau_excel = pandas.read_excel(AAU_file_first)
    except:
        aau_excel = []
        print('no "*AAU*.xlsx" detected')

    try:
        customers_file_name = glob.glob(Dir + "*customers*.csv")
        customers_file_name.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        customers_file_first = customers_file_name[0]  # convert from array to a string
        customers_csv = pandas.read_csv(customers_file_first)
        customers_csv.to_excel(f"{Dir}customers.xlsx", index=False, header=True)
    except:
        try:
            customers_file_name = glob.glob(Dir + "*customers*.xlsx")
            customers_file_name.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            customers_file_first = customers_file_name[0]  # convert from array to a string
            customers_csv = customers_file_first
            customers_csv = pandas.read_excel(customers_file_first)
        except PermissionError:
            print("Close the required files that are open and try again.")
            return input("press enter to close")

    return aau_excel, customers_csv, autopilot


def all_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv):
    pos_array = []

    for y0 in range(LENGTE_PARTNER):
        csv_domein = str(customers_csv.iat[y0, Primaire_domeinaam_])

        for y1 in range(LENGTE_AAU):
            school_naam = str(aau_excel.iat[y1, school_])
            tenant_domein = str(aau_excel.iat[y1, tenant_])
            mail = str(aau_excel.iat[y1, technissche_contactpersoon_])
            opmerking_sales = str(aau_excel.iat[y1, opmerking_sales_])

            if (tenant_domein == empty and csv_domein in mail) or (csv_domein in tenant_domein):
                pos_array.append(f"{school_naam} ### {csv_domein} ### {opmerking_sales}")
                break

    return pos_array


def pos_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv):
    pos_array = []

    for y0 in range(LENGTE_PARTNER):
        csv_domein = str(customers_csv.iat[y0, Primaire_domeinaam_])

        for y1 in range(LENGTE_AAU):
            school_naam = str(aau_excel.iat[y1, school_])
            tenant_domein = str(aau_excel.iat[y1, tenant_])
            mail = str(aau_excel.iat[y1, technissche_contactpersoon_])
            opmerking_sales = str(aau_excel.iat[y1, opmerking_sales_])

            skip_switch_exclude_string = False
            for exclude_string in exclude_array:
                if exclude_string in opmerking_sales:
                    skip_switch_exclude_string = True
                    break

            if (skip_switch_exclude_string == False) and ((tenant_domein == empty and csv_domein in mail) or (csv_domein in tenant_domein)):
                pos_array.append(f"{school_naam} ### {csv_domein} ### {opmerking_sales}")
                break

    return pos_array


def func_autopilot(LENGTE_AUTOPILOT, LENGTE_PARTNER, autopilot, customers_csv):
    mail_ = 2
    school_ = 1
    pos_array = []

    for y0 in range(LENGTE_PARTNER):
        csv_domein = str(customers_csv.iat[y0, Primaire_domeinaam_])

        for y1 in range(LENGTE_AUTOPILOT):
            school_naam = str(autopilot.iat[y1, school_])
            mail_domein = str(autopilot.iat[y1, mail_])

            if csv_domein in mail_domein:
                pos_array.append(f"{school_naam} ### {csv_domein}")
                break

    return pos_array


main()
