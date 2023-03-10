from persistence.PermFinanceirasRepository import PermissoesFinanceirasRepository


class PermFinanceiroService: 
     def __init__(self):
        self.__PermFinDao = PermissoesFinanceirasRepository()
     
     
      
     def PermissionamentoResultFin(self):
       retorno = []
       resultados = self.__PermFinDao.PermissionamentoResultFin()
       for resultado in resultados:
             retorno.append(resultado['perm_gestfin_resultfin'])
      
       return retorno[0]        
 
     def gerarPermissionamentoResultFin(self, boolresultfin):
        self.__PermFinDao.gerarPermissionamentoResultFin(boolresultfin)
 