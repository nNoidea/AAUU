import pandas as pd
from tkinter import *
from tkinter.filedialog import askdirectory

# numbers indicate x-values, the first row (y = 0) is skipped by pandas, so (0,0) is a true value and not a label.
opmerking_sales = 7
tenant = 6
technissche_contactpersoon = 4
school = 2

Bedrijfsnaam = 1
Primaire_domeinaam = 2


def main():
    Tk().withdraw()  # use to hide tkinter window
    Dir = askdirectory(title="Select Folder")  # shows dialog box and return the path.
    aau_excel = pd.read_excel(f"{Dir}/AAU.xlsx")
    customers_csv = pd.read_excel(f"{Dir}/customers.xlsx")

    LENGTE_AAU = len(aau_excel)
    LENGTE_PARTNER = len(customers_csv)


    print("...")
    pos_array = pos_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv)
    neg_array = neg_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv)
    print(pos_array)


def pos_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv):
    pos_array = []

    for y in range(LENGTE_PARTNER):
        csv_domein = customers_csv.iat[y, Primaire_domeinaam]

        for y1 in range(LENGTE_AAU):
            tenant_domein = aau_excel.iat[y1, tenant]
            mail = aau_excel.iat[y1, technissche_contactpersoon]

            if (tenant_domein != "" and str(csv_domein) in str(tenant_domein)) or str(csv_domein) in str(mail):
                pos_array.append(csv_domein)
                break

    return pos_array


def neg_comparer(LENGTE_AAU, LENGTE_PARTNER, aau_excel, customers_csv):
    neg_array = []


    return neg_array


main()
