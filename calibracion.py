from getpass import getuser
import pyodbc
from subprocess import call
import pandas as pd,csv
from concurrent import futures
import csv
import queue
import os

_conn_params = {
    "server": 'PUGINSQLP01',
    "database": 'BCP_GDH_PA_STAGE',
    "trusted_connection": "Yes",
    "driver": "{SQL Server}",
}

def crear_csv(df,file_name):
    file_name = file_name
    df.to_csv(file_name,index=False,sep='|',encoding='UTF-16',header=False,quotechar='`',quoting=csv.QUOTE_NONNUMERIC)

def select(q,t_params=()):
    cnxn = pyodbc.connect(**_conn_params)
    cnxn.autocommit = True
    c = cnxn.cursor()
    c.execute(q,t_params)

    columns = [column[0] for column in c.description]
    results = []
    for row in c.fetchall():
        results.append(dict(zip(columns, row)))

    cnxn.close()

    return results