from persistence.DbConnecion import DbConnection
from entity.Usuariopc import Usuario
from entity.PermissaoFinanceiro import  PermFinanceiro
from mysql.connector import Error

class LoginRepository:
    #Construtor... Não esquece carai
    def __init__(self):
        pass
    
    def EntrarComLogin(self, nome, senha):
        #con = mysql.connector.connect(host='127.0.0.1', user='root', database='ezsys',password='q9w8e7#MTB')
        db = DbConnection()
        
        usuario = None
        try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (nome, senha)
            cursor.execute('SELECT usuario, senha, usuarioid FROM ezsys.usuarios WHERE usuario = %s AND senha = %s', valores)
            #cursor.execute('SELECT nome, senha FROM ez-sys.senha WHERE nome = %s AND senha = %s', valores)
            ver = cursor.fetchone()
            print(ver)
            if(ver != None):
                usuario = Usuario(ver['usuarioid'],ver['usuario'], ver['senha'], None)

            db.closeConn(cursor)    
            return usuario
        except Error as e :
            print(e)
            return usuario
        
        
    def permResultFin(self,usuarioid): 
       db = DbConnection()
        
       try:
            con = db.getConn()
            cursor = con.cursor(dictionary = True)
            valores = (usuarioid)
            usuario='select * from ezsys.perm_financeiro where fk_usuarioid = "{}"'.format(valores)
            cursor.execute(usuario)
            ver = cursor.fetchone()

            if(ver != None):
             permissao = PermFinanceiro(ver["fin_id"], ver["perm_cadfin"], ver["perm_cadfin_contMonetaria"], ver["perm_cadfin_planReceitas"],
                                        ver["perm_cadfin_PlanPagt"], ver["perm_cadfin_grupdespesas"], ver["perm_cadfin_centdecustos"], ver["perm_cadfin_opfinanceiras"],
                                        ver["perm_cadfin_codfiscais"], ver["perm_cadfin_condpagamentos"], ver["perm_cadfin_benpatrimonios"],
                                        ver["perm_cadfin_situtributarias"], ver["perm_cadfin_classfiscais"], ver["perm_cadfin_formspagamento"],
                                        ver["perm_contpagar"], ver["perm_contreceber"], ver["perm_saimonetarias"], ver["perm_entmonetarias"], 
                                        ver["perm_cheques"], ver["perm_gestfin"], ver["perm_gestfin_extcontas"], ver["perm_gestfin_fluxocaixa"],
                                        ver["perm_gestfin_receitaseentradas"], ver["perm_gestfin_despesas"], ver["perm_gestfin_despcentrodecusto"],
                                        ver["perm_gestfin_posmonetfin"], ver["perm_gestfin_resultfin"])
            print(ver)
            db.closeConn(cursor)    
            return permissao
       except Error as e :
            print(e)
            return None
              
   
        
    pass

    
    
   
