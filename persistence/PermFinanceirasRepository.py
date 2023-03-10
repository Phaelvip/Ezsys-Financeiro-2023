from mysql.connector import Error
from persistence.DbConnecion import DbConnection
from entity.Usuariopc import Usuario
from entity.PermissaoFinanceiro import  PermFinanceiro


class PermissoesFinanceirasRepository:

    def gerarPermissionamentoResultFin(self,boolresultfin): 
       db = DbConnection()
        
       usuario = None
       try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (boolresultfin)
            #cursor.execute('SELECT nome, senha FROM ``ez-sys``.senha WHERE nome = %s AND senha = %s', valores)
            updateString = 'UPDATE `ezsys`.`perm_financeiro` SET  perm_gestfin_resultfin = {} where id= "%s";'.format(valores)
            
            cursor.execute(updateString)      
            print('ROLOU O UPDATE PARA ' + str(boolresultfin))
            
            con.commit()
            db.closeConn(cursor)    
            
       except Error as e :
            print(e)
            
    def PermissionamentoResultFin(self): 
       db = DbConnection()
        
       usuario = None
       try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            #valores = (cidade)
            #cursor.execute('SELECT nome, senha FROM ``ez-sys``.senha WHERE nome = %s AND senha = %s', valores)
            cursor.execute('SELECT perm_gestfin_resultfin FROM `ezsys`.perm_financeiro;')      
            ver = cursor.fetchall()
            
            #print(ver)
            
            
            db.closeConn(cursor)    
            return ver
       except Error as e :
            print(e)
            return usuario     
            
            
            
             

