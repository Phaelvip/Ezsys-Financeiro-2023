import mysql.connector
from mysql.connector import Error
from persistence.DbConnecion import DbConnection

class dadosRepository:



    def Atualizadados(self):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        cursor.execute("drop table ezsys.saida;")
        cursor.execute("drop table ezsys.saicusto;")
        cursor.execute("drop table ezsys.vendas;")
        cursor.execute("drop table ezsys.entrada;")
        cursor.execute("create table ezsys.saida select * from ezipa.saida;")
        cursor.execute("create table ezsys.saicusto select * from ezipa.saicusto;")
        cursor.execute("create table ezsys.vendas select * from ezipa.vendas;")
        cursor.execute("create table ezsys.entrada select * from ezipa.entrada;")

        ver = cursor.fetchone()
        print(ver)
        cursor.execute("drop table `ezsys`.ccustos2;")
        db.closeConn(cursor)    
        return ver
      except Error as e :
        print(e)
        return usuario
    

