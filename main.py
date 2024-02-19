import datetime
import pyodbc
from xbbg import blp

import paths_manager


paths_info = {
    'MONITOR.BDPRODUCTOS': ['Base de Datos', '5 - Bases de Datos/PRODUCTOS.accdb', 'FILE'],
}

paths = paths_manager.get_paths(paths_info)
db_file = paths["MONITOR.BDPRODUCTOS"]

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'fr'DBQ={db_file};')
conn.setdecoding(pyodbc.SQL_WCHAR, encoding='latin-1')
cursor = conn.cursor()

today = datetime.datetime(2024, 2, 9).replace(hour=0, minute=0, second=0, microsecond=0)
today_minus_15 = today - datetime.timedelta(days=15)

cursor.execute(
    """
    SELECT a.ZestID, s.Subyacente
    FROM T_AUTOCALL a
    INNER JOIN REL_NOTASUBYAC s ON a.ZestID = s.ZestID
    WHERE a.FinalValueDate >= ?
    AND a.ISSUEDATE <= ?
    AND (a.FECHAAUTOCALL IS NULL OR a.FECHAAUTOCALL >= ?)
    AND s.FechaSalida IS NULL
    ORDER BY a.ZestID;
    """,
    (today, today, today)
)

active_notes = cursor.fetchall()
for note in active_notes:
    print(note)

subys = list(set([x[1] + " Equity" for x in active_notes]))
"""
for suby in subys:
    # Tantear el limite superior del intervalo con timedelta
    prices = blp.bds(suby, 'DVD_Hist_All', DVD_Start_Dt=today_minus_15.strftime('%Y%m%d'), DVD_End_Dt=today.strftime('%Y%m%d'))
    print(f"Subyacente {suby} #######################################################")
    print(prices.to_string())
"""
