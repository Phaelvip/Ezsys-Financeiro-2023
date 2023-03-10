class Usuario:
    
    def __init__(self, usuarioid, nome, senha, permissoes):
        self.__usuarioid = usuarioid
        self.__nome = nome
        self.__senha = senha
        self.__permissoes = permissoes
        
    def getusuarioid(self):
        return self.__usuarioid
    
    def getNome(self):
        return self.__nome
    
    def getSenha(self):
        return self.__senha
    
    def getPermissoes(self):
        return self.__permissoes
    
    def setPermissoes(self, permissoes):
        self.__permissoes = permissoes
        
    def __str__(self):
        return "usuarioid:"+ self.__usuarioid, "Nome: " + self.__nome + ", Senha: " + self.__senha