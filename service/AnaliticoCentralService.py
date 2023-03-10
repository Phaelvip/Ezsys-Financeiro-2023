from entity.dto.RelatorioAnalitico import RelatorioAnalitico
from persistence.VendaRepository import VendaRepository
from persistence.SaidaRepository import SaidaRepository
from persistence.ScriptsRepository import ScriptsRepository
import locale
from datetime import *
from  datetime import datetime
from mysql.connector import Error

class AnaliticoCentralService:

    def __init__(self):
        self.__scriptDao = VendaRepository()
        self.__scriptDao = ScriptsRepository()        
        self.__colunasAnaliticoGeral=[]
        locale.setlocale(locale.LC_ALL,'pt_BR')
        locale.setlocale(locale.LC_MONETARY,'pt_BR.UTF-8')
        self.__quantMeses = 0
        
    def Apagaranaliticoadm(self):
      linhas = []
      self.__scriptDao.truncateAdm()#trncate  
      return linhas
                
        
    def GerarRelatorioDetalhadoAnaliticoCentral(self, ano, mesFinal, mesInicial):
        
        linhas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        for mes in range(mesInicial, mesesInteiro,1):
            self.__quantMeses+=1
            dtInicial = "{}-{}-01".format(ano, mes)
            dtFinal = "{}-{}-31".format(ano, mes)                   
            analiticoAdm=(self.__scriptDao.analiticoAdm(dtInicial, dtFinal,'51'))#linha1   
            
            
    def GerarRelatorioAnaliticoCentral(self, ano, mesFinal, mesInicial, diainicial, diafinal):
        
        colunas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        #TEM QUE CONVERTER DE TEXTO PRA INTEIRO
        for mes in range(mesInicial, mesesInteiro,1):
            self.__quantMeses+=1
            #Template
            if(mes < 10):
                periododeCompetenciaInicial = "0{}/{}".format(mes,ano)
                periododeCompetenciafinal = "0{}/{}".format(mes,ano)
            else:
                periododeCompetenciaInicial = "{}/{}".format(mes,ano)
                periododeCompetenciafinal = "{}/{}".format(mes,ano)
            dtInicial = "{}-{}-01".format(ano, mes)
            dtFinal = "{}-{}-31".format(ano, mes)         
            
                  
