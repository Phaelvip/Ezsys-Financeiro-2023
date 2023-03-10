from mailbox import Mailbox
from optparse import Values
from string import octdigits
import mysql.connector
from mysql.connector import Error
from persistence.DbConnecion import DbConnection

class VendaRepository:
    
    def __init__(self):
        pass
    
    def consultarInformacoesNoBancoDeDados(self, cliente):       
            linhas = None
            con = mysql.connector.connect(host='127.0.0.1', user='root', database='eztestes',password='q9w8e7#MTB')
            con = Cursor.Conn()
            consulta_sql= None
            if(len(cliente)>0):
                consulta_sql= 'SELECT cliente, codigo, data, valorvenda, ID FROM `eztestes`.vendas where cliente="{}"'.format(cliente)
                print(consulta_sql)
            else:
                consulta_sql= "SELECT cliente, codigo, data, valorvenda, ID FROM `eztestes`.vendas"
                
                
            Cursor= con.cursor(dictionary = True)
            Cursor.execute(consulta_sql)
            linhas = Cursor.fetchall()
            print("Número total de registros retornados", Cursor.rowcount)
            print("\nMostrando os valores")
       
            Cursor.close()
            print("Concluido")
            return linhas    
            
    def inserirInformacoesNoBancoDeDados(self, cliente, codigo, data, valorvenda):
        try:
            # con = mysql.connector.connect(host='127.0.0.1', database='eztestes', user='root', password='q9w8e7#MTB')
            con= Cursor.Connection()
            # Inserir_vendas = """INSERT INTO vendas (cliente,codigo, data, valorvenda)
            #                     Values                  
            #                     ('?','?','?','?')
            # """
            
            queryString = """INSERT INTO VENDAS (cliente, codigo, data, valorvenda) values (%s,%s,%s,%s)"""
            values = (cliente, codigo, data, valorvenda)
            Cursor = con.cursor()
            Cursor.execute(queryString, values)
            con.commit()
            print(Cursor.rowcount,'Registros inseridos na tabela')
            Cursor.close()
        except Error as erro:
            print('Falha ao inserir dados no Mysql: {}'.format(erro))
        finally:
            con.closeConn(con, Cursor)
            print("conexão finalizada") 
        
    def buscarValorFinalDaVendaComNotaFiscal(self, dtInicial, dtFinal):
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCALNUM >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0   AND NFISCALNUM  >=1  AND NATOPERDS="VENDA" AND DATA >= %s AND DATA <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario

    
    
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCALNUM >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS="servico" AND DATA >= %s AND DATA <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario

    def buscarValorFinalDoServicoComNotaFiscal(self, dtInicial, dtFinal): 
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCALNUM >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where NFISCALNUM >=1 AND NATOPERDS="servico" AND DATA >= %s AND DATA <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario

    def buscarValorFinalDoServicoComBoleta(self, dtInicial, dtFinal): 
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCALNUM >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta >=1 AND NATOPERDS="VENDA" AND DATA >= %s AND DATA <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario
    
    def buscarValorFinalDaDevolucaoDaVenda(self, dtInicial, dtFinal): 
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where  NATOPERDS="DEVOLUCAO DE VENDA" AND DATA >= %s AND DATA <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario


    def buscaValorFinalDasVendasCanceladas(self, dtInicial, dtFinal): 
        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) AS VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas WHERE cancelada ="S" AND NATOPERDS="VENDA" AND DATA >= %s AND DATA <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario
  
    def buscaValorFinalDasDespesasPorCompetencia(self, periodoCompetenciainicial, periodoCompetenciaFinal, dtFinal): 

        #SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (periodoCompetenciainicial, periodoCompetenciaFinal, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valordespesa),2) as VALOR_FINAL FROM  `'+db.getDbName()+'`.entrada where compete >= %s and compete <= %s And tipodespesa = %s', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario
  
    def buscarValorFinalDasDespesasPorPeriodo(self, dtInicial, dtFinal, Compete,): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, Compete)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valordespesa),2) AS VALOR_FINAL FROM `'+db.getDbName()+'`.ENTRADA where data >= %s and  data <= %s and tipodespesa = %s', valores) 
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
    
    def buscarValorFinalDasDespesasPortipo(self, dtInicial, dtFinal, despesa1): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, despesa1)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT ROUND(SUM(VALOR),2) AS VALOR_FINAL FROM `'+db.getDbName()+'`.entrada WHERE data >= %s and data <= %s and TIPO = %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        pass

    def buscarValorFinalDasDespesasPortipoecompetencia(self, periodoCompetenciainicial, periodoCompetenciaFinal,  despesa1): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (periodoCompetenciainicial, periodoCompetenciaFinal, despesa1)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT ROUND(SUM(VALOR),2) AS VALOR_FINAL FROM `'+db.getDbName()+'`.ENTRADA WHERE compete >= %s and  compete <= %s and TIPO = %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        pass
    def buscarValorFinalDosJuros(self, dtInicial, dtFinal): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT ROUND(SUM(juros),2) AS VALOR_FINAL FROM`'+db.getDbName()+'`.entrada WHERE data >= %s and data <= %s;', valores)
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        pass
    

    def CadastrodeDadosdaWoodpaper(self, data, valor): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (data, valor)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            #cursor.execute('truncate table `'+db.getDbName()+'`.woodpaper')
            cursor.execute('insert into `ezsys`.woodpaper (data, valor) values (%s,%s)', valores)
            con.commit()
            db.closeConn(cursor)    
            return True
        except Error as e :
            print(e)
            return False    
        
    def BuscarDadosdaWoodpaper(self, dtinicial, dtFinal): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtinicial, dtFinal)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT round(valor,2) as VALOR_FINAL FROM `ezsys`.woodpaper where data >=%s and data <= %s;', valores)
            ver = cursor.fetchone()
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    

    
    def valorDeDepreciacao(self, valor): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        
        try:
            con = db.getConn()
            cursor = con.cursor()
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)  
            cursor.execute("delete from `ezsys`.depreciacao") 
            cursor.execute("truncate table  `ezsys`.depreciacao")             
            insertString = "insert into  `ezsys`.depreciacao (valor) values ( {} )".format(float(valor))
            cursor.execute(insertString)
            #values = (tuple(valor))
           
            con.commit()
            print(cursor.rowcount,'Registros inseridos na tabela')
            db.closeConn(cursor)    
            return True
        except Error as e :
            print(e)
            return False    
    
    def BuscarDadosDepreciacao(self): 
       # SELECT  round(SUM(valor),2) AS VALOR_FINAL as VALOR_FINAL FROM  `'+db.getDbName()+'`.vendas where boleta = 0 AND NFISCAL >=1  AND NATOPERDS='VENDA' AND DATA >= '2022-07-06' AND DATA <= '2022-07-07';
         #con = mysql.connector.connect(host='127.0.0.1', user='root', database=' `'+db.getDbName()+'`',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT  round(SUM(valor),2) as VALOR_FINAL FROM  `ezsys`.depreciacao;')
            ver = cursor.fetchone()
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario    
        
        

    
#caminho = '/caminho/até/o/seu/arquivo.xlsx'
#arquivo_excel = load_workbook(caminho) 
    pass


    
    
   
#caminho = '/caminho/até/o/seu/arquivo.xlsx'
#arquivo_excel = load_workbook(caminho) 