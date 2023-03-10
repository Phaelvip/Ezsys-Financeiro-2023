import mysql.connector
from mysql.connector import Error
from persistence.DbConnecion import DbConnection

class ScriptsRepository:



    def BuscaValordecusto(self, dtInicial, dtFinal, despesa, competencia):
        db = DbConnection()
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            cursor.execute("CREATE TEMPORARY TABLE IF NOT EXISTS `"+db.getDbName()+"`.ccustos2 (valor varchar(10) , porcento varchar(10) , total varchar(10) );")
            cursor.execute("insert into   `"+db.getDbName()+"`.ccustos2 (valor, porcento) SELECT e.valor, p.porcento  FROM   `"+db.getDbName()+"`.saida AS e  INNER JOIN   `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo  where e.data >= %s and  e.data <= %s and e.desp = %s and p.custo = %s; ", (dtInicial, dtFinal, despesa, competencia))
            cursor.execute("update   `"+db.getDbName()+"`.ccustos2 set total = (valor/100*porcento); ")
            cursor.execute("select round(sum(total),2) as  resultado from   `"+db.getDbName()+"`.ccustos2; ")
            ver = cursor.fetchone()
            print(ver)
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario
      
        pass  
    
    def BuscarValorPorCustoCompete(self, periododeCompetenciaInicial ,periododeCompetenciafinal, competencia, custo):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        valores = (periododeCompetenciaInicial ,periododeCompetenciafinal,  competencia, custo)
        cursor.execute("CREATE TEMPORARY TABLE IF NOT EXISTS `"+db.getDbName()+"`.ccustos2 (valor varchar(10) , porcento varchar(10) , total varchar(10) );")
        cursor.execute("insert into   `"+db.getDbName()+"`.ccustos2 (valor, porcento) SELECT e.valor, p.porcento  FROM   `"+db.getDbName()+"`.saida AS e  INNER JOIN   `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo  where e.compete >= %s and e.compete <= %s and e.desp = %s and p.custo = %s; ", (periododeCompetenciaInicial ,periododeCompetenciafinal, competencia, custo))
        cursor.execute("update  `"+db.getDbName()+"`.ccustos2 set total = (valor/100*porcento); ")
        cursor.execute("select round(sum(total),2) as resultado from  `"+db.getDbName()+"`.ccustos2; ")
        
        ver = cursor.fetchone()
        print(ver)
        db.closeConn(cursor)    
        return ver
      except Error as e :
        print(e)
        return usuario
      
    def TruncateValorCusto(self):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        cursor.execute("drop table IF EXISTS `"+db.getDbName()+"`.ccustos2;")
        con.commit()    
        db.closeConn(cursor)    
        return True
      except Error as e :
         print(e)
         return False         
              
    def BuscarValorPorCustoCompeteSemCusto(self, periododeCompetenciaInicial ,periododeCompetenciafinal, custo):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        valores = (periododeCompetenciaInicial ,periododeCompetenciafinal,  custo)
        cursor.execute("CREATE TEMPORARY TABLE IF NOT EXISTS `"+db.getDbName()+"`.ccustos2 (valor varchar(10) , porcento varchar(10) , total varchar(10) );")
        cursor.execute("insert into `"+db.getDbName()+"`.ccustos2 (valor, porcento) SELECT e.valor, p.porcento FROM `"+db.getDbName()+"`.saida AS e INNER JOIN `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo where e.compete >= %s and e.compete <= %s  and p.custo = %s;", (periododeCompetenciaInicial, periododeCompetenciafinal, custo))
        cursor.execute("update   `"+db.getDbName()+"`.ccustos2 set total = (valor/100*porcento); ")
        cursor.execute("select round(sum(total),2) as resultado from  `"+db.getDbName()+"`.ccustos2; ")
        
        ver = cursor.fetchone()
        print(ver)
        db.closeConn(cursor)    
        return ver
      except Error as e :
        print(e)
        return usuario
           
    def BuscarValorPorCustoData(self, dtInicial, dtFinal, custo):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        valores = (dtInicial, dtFinal, custo)
        cursor.execute("CREATE TEMPORARY TABLE IF NOT EXISTS `"+db.getDbName()+"`.ccustos2 (valor varchar(10) , porcento varchar(10) , total varchar(10) );")
        cursor.execute("insert into   `"+db.getDbName()+"`.ccustos2 (valor, porcento) SELECT e.valor, p.porcento  FROM   `"+db.getDbName()+"`.saida AS e  INNER JOIN   `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo  where e.data >= %s and  e.data <= %s and p.custo = %s", (dtInicial, dtFinal, custo))
        cursor.execute("update   `"+db.getDbName()+"`.ccustos2 set total = (valor/100*porcento); ")
        cursor.execute("select round(sum(total),2) as resultado from  `"+db.getDbName()+"`.ccustos2; ")
        
        ver = cursor.fetchone()
        print(ver)
        db.closeConn(cursor)    
        return ver
      except Error as e :
        print(e)
        return usuario
      
    def truncateAdm(self):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
      
        cursor.execute("truncate table ezsys.analiticoAdm;")
        
        con.commit()    
        db.closeConn(cursor)    
        return True
      except Error as e :
         print(e)
         return False         
      
    def analiticoAdm(self, dtInicial, dtFinal,custo):
      db = DbConnection()
      usuario = None  
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        valores = (dtInicial, dtFinal)
        cursor.execute("insert into ezsys.analiticoAdm (DESP, DATA, CREDOR, DISCRIMINA, CODIGO, valor, porcento) SELECT e.desp, e.data, e.credor, e.discrimina, e.codigo,e.valor, p.porcento FROM  `"+db.getDbName()+"`.saida AS e INNER JOIN  `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo where e.data >= %s and  e.data <= %s and p.custo = 1 order by e.desp, e.data;",valores)
        cursor.execute("update  ezsys.analiticoAdm set total = round((valor/100*porcento),2) ;")
        cursor.execute("update ezsys.analiticoAdm AS cc3 INNER JOIN ezsys.class AS cl ON cl.codigo = cc3.desp SET cc3.nomedesp = cl.nome;")


        con.commit()    
        db.closeConn(cursor)    
        return True
      except Error as e :
         print(e)
         return False         




    
    def DropTableValorAnalitico(self):
      db = DbConnection()
      usuario = None
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        cursor.execute("drop table IF EXISTS `"+db.getDbName()+"`.ccustos3;")
        con.commit()    
        db.closeConn(cursor)    
        return True
      except Error as e :
         print(e)
         return False     
             
    def BuscardespesaAnalitico(self, dtInicial, dtFinal, despesa): 
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, despesa)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute("CREATE TEMPORARY TABLE IF NOT EXISTS `"+db.getDbName()+"`.ccustos3 (valor varchar(10) , porcento varchar(10) , total varchar(10) );")
            cursor.execute("insert into   `"+db.getDbName()+"`.ccustos3 (valor, porcento) SELECT e.valor, p.porcento  FROM   `"+db.getDbName()+"`.saida AS e  INNER JOIN   `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo  where e.data >= %s and  e.data <= %s and e.desp = %s and p.custo = 1; ", (dtInicial, dtFinal, despesa))
            cursor.execute("update   `"+db.getDbName()+"`.ccustos3 set total = (valor/100*porcento); ")
            cursor.execute("select round(sum(total),2) as  total from   `"+db.getDbName()+"`.ccustos3; ")
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario        
          
    def BuscardespesaAnaliticodetalhado(self, dtInicial, dtFinal, despesa): 

        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, despesa)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)         
            cursor.execute('SELECT data, credor, discrimina, total FROM ezsys.analiticoAdm where data >= %s and data <= %s and desp=%s;', valores)
            
            ver = cursor.fetchall()
            #print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario                   
      
    def analiticoAdmCompetencia(self, dtInicial, dtFinal,custo):
      db = DbConnection()
      usuario = None  
      try:
        con = db.getConn()
        cursor = con.cursor(dictionary = True)
        valores = (dtInicial, dtFinal)
        cursor.execute("insert into ezsys.analiticoAdm (DESP, DATA, CREDOR, DISCRIMINA, CODIGO, valor, porcento, compete) SELECT e.desp, e.data, e.credor, e.discrimina, e.codigo,e.valor, p.porcento, e.compete FROM  `"+db.getDbName()+"`.saida AS e INNER JOIN  `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo where e.compete >= %s and e.compete <= %s  and p.custo = 1 order by e.desp;",valores)
        cursor.execute("update  ezsys.analiticoAdm set total = round((valor/100*porcento),2) ;")
        cursor.execute("update ezsys.analiticoAdm AS cc3 INNER JOIN ezsys.class AS cl ON cl.codigo = cc3.desp SET cc3.nomedesp = cl.nome;")


        con.commit()    
        db.closeConn(cursor)    
        return True
      except Error as e :
         print(e)
         return False          
    
    def BuscardespesaAnaliticoCompetencia(self, dtInicial, dtFinal, despesa): 
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, despesa)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute("CREATE TEMPORARY TABLE IF NOT EXISTS `"+db.getDbName()+"`.ccustos3 (valor varchar(10) , porcento varchar(10) , total varchar(10) );")
            cursor.execute("insert into   `"+db.getDbName()+"`.ccustos3 (valor, porcento) SELECT e.valor, p.porcento  FROM   `"+db.getDbName()+"`.saida AS e  INNER JOIN   `"+db.getDbName()+"`.saicusto AS p ON e.codigo = p.codigo  where e.compete >= %s and e.compete <= %s and e.desp = %s and p.custo = 1; ", (dtInicial, dtFinal, despesa))
            cursor.execute("update   `"+db.getDbName()+"`.ccustos3 set total = (valor/100*porcento); ")
            cursor.execute("select round(sum(total),2) as  total from   `"+db.getDbName()+"`.ccustos3; ")
            ver = cursor.fetchone()
            print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario         

    def BuscardespesaAnaliticodetalhadoCompetencia(self, dtInicial, dtFinal, despesa): 

        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (dtInicial, dtFinal, despesa)
            #cursor.execute('SELECT nome, senha FROM ` `'+db.getDbName()+'``.senha WHERE nome = %s AND senha = %s', valores)         
            cursor.execute('SELECT data, credor, discrimina, total FROM ezsys.analiticoAdm where compete >= %s and compete <= %s and desp=%s;', valores)
            
            ver = cursor.fetchall()
            #print(ver)
            
            
            db.closeConn(cursor)    
            return ver
        except Error as e :
            print(e)
            return usuario   
      
pass




        