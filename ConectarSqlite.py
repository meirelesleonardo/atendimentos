# coding=utf-8
#from sshtunnel import SSHTunnelForwarder
#import openpyxl
#from PyQt5.QtCore import  QDate, QDateTime
#from openpyxl.styles import Alignment
import sqlite3


# senha = 'suport2019'
# host = '127.0.0.1'
# user = 'app'
bd = r'\\10.7.51.11\ddt$\GSU\01-Suport\Agenda.tab'




# Conexão
def conectarBd (sql):
	# Tunnel
    # server = SSHTunnelForwarder('10.7.51.61', ssh_username="app", ssh_password="Gmrio12$", remote_bind_address=('127.0.0.1', 3306))
    #
    # server.start()
    # port=server.local_bind_port

    try:
        conectar = sqlite3.connect(bd)#(r'\\10.7.51.11\ddt$\GSU\01-Suport\Agenda.tab') #
        cursor = conectar.cursor()
        if 'SELECT' in sql:
            cursor.execute(sql)
            resultado = cursor.fetchall()
            conectar.close()
            return resultado
        else:
            cursor.execute(sql)
            conectar.commit()
        conectar.close()
        #server.stop()
    except:
        pass
        #server.stop()

def insertDados(dic,tabela):
    del dic['id']
    campos = tuple(dic.keys())
    valores = tuple(dic.values())
    sql = ("INSERT INTO "+tabela+str(campos)+" VALUES "+str(valores))
    conectarBd(sql)
    return sql
def getId(tabela):
    sql = ("SELECT MAX(id) FROM "+str(tabela))
    id = conectarBd(sql)
    if id[0][0] == None:
        id = 0
    else:
        id = int(id[0][0])
    return id


def getColumns (tabela):

    try:
        conectar = sqlite3.connect(bd)#(r'\\10.7.51.11\ddt$\GSU\01-Suport\Agenda.tab') #
        cursor = conectar.cursor()
        columns = []
        cursor.execute("PRAGMA table_info('"+tabela+"');")
        resultado = cursor.fetchall()
        conectar.close()
        for c in resultado:
            columns.append(c[1])
        return columns
    except:
        pass
        #server.stop()
#if __name__ == "__main__":
    #print(getColumns("bloqueio_impressoras"))
