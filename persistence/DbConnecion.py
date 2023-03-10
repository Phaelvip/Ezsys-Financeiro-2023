import mysql.connector
from mysql.connector import Error
from  Leitura import Leitura

class DbConnection:
    
    def __init__(self):
        self.leitura = Leitura()
        try:
            #self.__conn = mysql.connector.connect(host='eziparj.ddns.net', port='7129', database='ezsys', user='root', password='q9w8e7#MTB')
            self.__conn = mysql.connector.connect(host=self.leitura.host,port=self.leitura.port, database=self.leitura.dbname, user=self.leitura.username, password=self.leitura.password)
        except Error as erro:
            print(erro)
    
    def getConn(self):
        return self.__conn
    
    def getDbName(self):
        return self.leitura.dbname
            
    
    def closeConn(self, cursor):
        try:
            if(self.__conn.is_connected()):
                self.__conn.close()
                cursor.close()
        except Error as erro:
            print(erro)
    
    pass
