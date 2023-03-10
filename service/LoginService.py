
from persistence.LoginRepository import LoginRepository
from persistence.PermFinanceirasRepository import PermissoesFinanceirasRepository
from tkinter import messagebox

class LoginService:
    def __init__(self):
        #Dao quer dizer Data Access Object
        self.__loginDao = LoginRepository()
        self.__permfinDao = PermissoesFinanceirasRepository()
        pass
    
    def realizarLoginComUsuarioESenha(self, usuario, senha):
        resultset = self.__loginDao.EntrarComLogin(usuario, senha)
        if resultset != None: 
            return resultset
        else:
            return False

    def verificarPermissionamento(self,usuarioid):

        resultado= self.__loginDao.permResultFin(usuarioid)
        return  resultado
    
        #if resultado == True:
        #    return resultado
        #else:
        #    return False
pass


