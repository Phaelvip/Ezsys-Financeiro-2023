import mysql.connector
from mysql.connector import Error
from persistence.DbConnecion import DbConnection

class SaidaRepository:
    
    def __init__(self):
        pass
    
    def buscarValorFinalDaSaidaPorcompetencia(self, periodoCompetenciainicial, periodoCompetenciaFinal, desp): 
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM ``'+db.getDbName()+'``.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database='`ez-sys`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (periodoCompetenciainicial, periodoCompetenciaFinal, desp)
            #cursor.execute('SELECT nome, senha FROM ``ez-sys``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT round(SUM(valor),2) AS VALOR_FINAL FROM `'+db.getDbName()+'`.saida where compete >= %s and compete <= %s and desp= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        
    def buscarValorFinalDasaidaPorPeriodo(self, dtInicial, dtFinal, Compete): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM ``'+db.getDbName()+'``.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database='`ez-sys`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, Compete)
            #cursor.execute('SELECT nome, senha FROM ``ez-sys``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM `'+db.getDbName()+'`.saida where data >= %s and data <= %s and desp= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
    def buscarValorFinalDasaidaPorPeriodo1(self,  dtInicial, dtFinal, despesa1, despesa2): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  ``'+db.getDbName()+'``.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' ``'+db.getDbName()+'``',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, despesa1, despesa2)
            #cursor.execute('SELECT nome, senha FROM ` ``'+db.getDbName()+'```.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM `'+db.getDbName()+'`.saida where data >= %s and data <= %s and (desp= %s or desp= %s);', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        pass
    def buscarValorFinalDasDespesasPorCompetencia(self,periodoCompetenciainicial, periodoCompetenciaFinal, desp): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  ``'+db.getDbName()+'``.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' ``'+db.getDbName()+'``',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (periodoCompetenciainicial, periodoCompetenciaFinal, desp)
            #cursor.execute('SELECT nome, senha FROM ` ``'+db.getDbName()+'```.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) as VALOR_FINAL FROM `'+db.getDbName()+'`.saida where compete >= %s and compete <= %s and desp= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        
    def buscarValorFinalDasaidaPorCompetencia1(self, periodoCompetenciainicial, periodoCompetenciaFinal, despesa1, despesa2): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  ``'+db.getDbName()+'``.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' ``'+db.getDbName()+'``',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (periodoCompetenciainicial, periodoCompetenciaFinal, despesa1, despesa2)
            #cursor.execute('SELECT nome, senha FROM ` ``'+db.getDbName()+'```.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) as VALOR_FINAL FROM `'+db.getDbName()+'`.saida where compete >= %s and compete <= %s  and (desp=%s OR desp=%s);', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        
        
    
    pass

