import datetime
import sys

import pandas as pd
from dateutil.relativedelta import relativedelta
import pyodbc
import win32com.client as win32
from xbbg import blp

import paths_manager


def send_email_with_output_file(addresses, attachment):
    """
    Sends an email to an input address, with and input attachment (path to the file)
    Parameters:
        type (str): Type of positions in the report (excel file)
        addresses (list): Addresses the email will be sent to.
        attachment (str): Path to the file that will be attached to the email
    """

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(addresses)
    mail.Subject = f'Próximos dividendos a pagar para las acciones del fondo ZBLX'
    mail.Body = f':)'
    mail.Attachments.Add(attachment)
    mail.Send()

def main():
    paths_info = {
        'MONITOR.BDPRODUCTOS': ['Base de Datos', '5 - Bases de Datos/PRODUCTOS.accdb', 'FILE'],
    }

    paths = paths_manager.get_paths(paths_info)
    db_file = paths["MONITOR.BDPRODUCTOS"]

    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'fr'DBQ={db_file};')
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin-1')
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT DISTINCT Fecha FROM POSITIONS_ZBLX
        """
    )
    last_date = max([x[0] for x in cursor.fetchall()])

    today = datetime.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    start_date = today - relativedelta(years=1)
    oldest_payable_date = today - relativedelta(days=30)

    cursor.execute(
        """
        SELECT p.ZestID
        FROM T_DIRECTORIO d
        INNER JOIN POSITIONS_ZBLX p ON d.ZestID = p.ZestID
        WHERE p.Fecha = ?
        AND d.Tipo = ?
        ORDER BY p.ZestID;
        """,
        (last_date, "Equity")
    )

    active_stocks = cursor.fetchall()
    for note in active_stocks:
        print(note)

    subys = list(set([x[0] + " Equity" for x in active_stocks]))

    divs_by_suby = pd.DataFrame()
    first_df_was_found = False
    for i in range(len(subys)):
        prices = blp.bds(subys[0], 'DVD_Hist_All', DVD_Start_Dt=start_date.strftime('%Y%m%d')).reset_index()
        print(f"Subyacente {subys[0]} #######################################################")
        if prices:
            if not first_df_was_found:
                divs_by_suby = prices
                first_df_was_found = True
            else:
                divs_by_suby = pd.concat([divs_by_suby, prices])
            print(prices.to_string())

    # I asume the column payable_date is of type Datetime. If that's not the case, adjust accordingly
    divs_by_suby = divs_by_suby[divs_by_suby["payable_date"] > oldest_payable_date]

    filename = "divs_by_suby_" + today.strftime("%Y-%m-%d") + ".xlsx"
    divs_by_suby.to_excel(filename, index=False)

    addresses = [
        "kevinbarzola@zest.pe",
    ]
    send_email_with_output_file(addresses, filename)


if __name__ == "__main__":
    # This code will only be executed
    # if the script is run as the main program
    if datetime.datetime.today().weekday() == 0:
        main()