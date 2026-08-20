import urllib
import urllib.parse
from sqlalchemy import create_engine, text
import os

# Ainda nao


Matricula = os.getenv("USERNAME") or os.getenv("USER") or "desconhecido" # substituir futuramente, ja uso isso em outros projetos, melhor deixar como uma funçao que retorna isso em um arquivo so de funçoes

def _pick_sqlserver_odbc_driver() -> str:
    forced = (os.getenv("x") or "").strip()
    if forced:
        if forced.isdigit():
            return f"ODBC Driver {forced} for SQL Server"
        else:
            return "Erro!!!!"
def crair_motor_sqlserver(server, database, user, password):
    driver_name = _pick_sqlserver_odbc_driver
    driver_q = urllib.parse.quote_plus(driver_name)

    password = urllib.parse.quote_plus(password)
    extra = ""
    conn_str = (
        f"mssql+pyodbc://{user}:{password}@{server}/{database}"
        f"?driver={driver_q}"
    )
    engine = create_engine(conn_str, connect_args={"fast_executemany": True})
    return engine, driver_name

def registrar_bug(): # pegar depois o codigo de provisao
    pass