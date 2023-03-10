class Relatorio:
    
    def __init__(self, dtInicial, dtFinal, Compete,periododeCompetencia):
        self.__dtInicial = dtInicial
        self.__dtFinal = dtFinal
        self.__Compete = Compete
        self.__periododeCompetencia = periododeCompetencia
        self.__caminho
    
    def getperiodoEntry(self):
        return self.__dtInicial
    
    def getateEntry(self):
        return  self.__dtFinal
    
    def  getcompetenciaEntry(self):
        return  self.__Compete
    
    def  getPeriodoEntry(self):
        return  self.__periododeCompetencia
    
    def __str__(self):
        return "dtInicial" + self.__dtInicial + ", dtFinal: " + self.__dtFinal,