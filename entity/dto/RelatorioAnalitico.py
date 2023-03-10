class RelatorioAnalitico:
    
    def __init__(self):
        self.__total=0
        self.__nomeMes=""       
        self.__Desp=""
        self.__nomeDesp = ""
        self.__data = ""
        self.__credor = ""
        self.__discrimina = ""   
        self.__valor=""             
        self.__ALUGUEL = ""
        self.__IPTU = ""
        self.__CEDAE = ""
        self.__LIGHT = ""  
        self.__CEG=""      
        self.__TELEFONIA=""      
        self.__PNP="" 
        self.__CONTABILIDADE=""      
        self.__SIGRAF=""  
        self.__INFORMTATICAECONTRATOS=""     
        self.__MATERIALESCRITORIO=""       
        self.__MATERIALLIMPEZA=""       
        self.__CARTUCHOSPARAIMPRESSORAS=""       
        self.__ONIBUSCONTINUO=""           
        self.__CORREIOS=""       
        self.__DESPESASBANCARIAS=""       
        self.__OUTROSADMINISTRATIVOS=""       
        self.__KOMBI=""       
        self.__MOTOBOY=""
        self.__OUTROSFRETES=""
        self.__TRANSPORTADORAS=""                                                                                                          
        self.__SEDEXREMESSAVENDA=""
        self.__PEDAGIOFRETEVENDA=""       
        self.__SIMPLES=""           
        self.__OUTROSIMPOSTOS=""     
        self.__SALARIOS=""     
        self.__RECISOES=""   
        self.__VALETRANSPORTE=""  
        self.__VALEALIMENTACAO=""
        self.__INSS=""
        self.__FGTS="" #40
        self.__AJUDADECUSTO=""#41 
        self.__FERIAS="" #42
        self.__HORAEXTRA="" #43
        self.__TESTES="" #44
        self.__CESTABASICA=""#45   
        self.__LANCHES="" #46
        self.__PESQUISADESENVOLVIMENTO="" #46
        self.__FARMACIA="" #47
        self.__TREINAMENTO="" #48
        self.__BISCATE="" #49
        self.__OUTROS="" #50
        self.__EMPRESTIMOSINTERNOS="" #71
        self.__RESULTADOS="" #72 
        self.__PROLABOREDIRETORIA="" #73
        self.__PROLABOREGERENCIA="" #74
        self.__DESPCADASTRAISCLIENTES="" #75
        self.__EXAMES="" #77
        self.__CIPA="" #78  
        self.__SERVSEGURANCA="" #79
        self.__MINFORMATICA="" #86
        self.__PLANODESAUDE="" #91
        self.__FERIASTRABALHADAS="" #92
        self.__EPIUNIFORMES="" #93
        self.__FESTASCONFRATERNIZACOES="" #95
        self.__MPREDIAL="" #97
        self.__MELETRICA=""#98
        self.__MARCONDICIONADO=""#99 
        self.__MBEBEDOURO=""#100 
        self.__MMOVEIS=""#101 
        self.__MOUTROSMANUT=""#102 
        self.__MHIDRAULICA=""#103
        self.__BMAQUINASEEQUIP=""#104
        self.__OBRASEMELHORIAS=""#105
        self.__BMOVEIS=""#106
        self.__BINFORMATICAEHARD=""#107
        self.__BFERRAMENTAS=""#109
        self.__AJUDACOMBUSTIVEL=""#110  
        self.__DESPREPRESENTACAO=""#111
        self.__PROPAGANDA="" #112
        self.__BRINDES=""#113
        self.__BINFORMATICASOFT=""#120 
        self.__IMAQUINASEQUIPS=""#134  
        self.__ICOMPUTADORESHARD=""#135  
        self.__IOUTROSINVESTIMENTOS=""#140  
        self.__PREMIOSEGRATIFICACOES=""#141  
        self.__DEVOLUCAOVENDA=""#144  
        self.__IRRF=""#148  
        self.__IOF=""#149
        self.__JUROSEMPRESTIMOS=""#151
        self.__LICENCIAMENTOS_E_CERTIDOES=""#163
        #self.__TOTAL=""#151         
        self.__RelatorioAnalitico=""
                                                                                           
                                                                                                                                      
        
                                                        
                                                                                                                                                                                                                                             
                       
        
    def build(self):
        return self
    
    def gerarTotal(self):
        valor = ()
        self.__total=valor    
    
    
    def RelatórioAnaliticoDetalhado(self, valor):
        self.__RelatorioAnalitico = self.validaValor(valor)
        return self     
    def getRelatorioAnalitico(self):
        return self.__RelatorioAnalitico      

    #def LinhasIndice(self, valor):
        self.LinhasIndice = valor
        return self    
    
    def nomeMes(self, valor):
        self.__nomeMes = valor
        return self
    
    def Desp(self, valor):
        self.__Desp= valor
        return self    
    
    def nomeDesp(self, valor):
        self.__nomeDesp = valor
        return self
    
    def data(self, valor):
        self.__data = valor
        return self    
    
    def credor(self, valor):
        self.__credor = valor
        return self  
    
    def discrimina(self, valor):
        self.__discrimina = valor
        return self   
    
    def valor(self, valor):
        self.__valor = valor
        return self           
    
    def Aluguel(self, valor):
        self.__ALUGUEL = self.validaValor(valor)
        return self      
    
    def iptu(self, valor):
        self.__IPTU = self.validaValor(valor)
        return self
    
    def Cedae(self, valor):
        self.__CEDAE = self.validaValor(valor)
        return self         

    def Light(self, valor):
        self.__LIGHT = self.validaValor(valor)
        return self   

    def Ceg(self, valor):
        self.__CEG = self.validaValor(valor)
        return self 

    def Telefonia(self, valor):
        self.__TELEFONIA = self.validaValor(valor)
        return self     
    
    def Pnp(self, valor):
        self.__PNP = self.validaValor(valor)
        return self  
    
    def Contabilidade(self, valor):
        self.__CONTABILIDADE = self.validaValor(valor)
        return self 
    
    def Sigraf(self, valor):
        self.__SIGRAF = self.validaValor(valor)
        return self         
    
    def Informatica(self, valor):
        self.__INFORMTATICAECONTRATOS = self.validaValor(valor)
        return self        
    
    def MaterialdeEscritorio(self, valor):
        self.__MATERIALESCRITORIO = self.validaValor(valor)
        return self       
    
    def MaterialLimpeza(self, valor):
        self.__MATERIALLIMPEZA = self.validaValor(valor)
        return self  
    
    def CartuchosparaImpressoras(self, valor):
        self.__CARTUCHOSPARAIMPRESSORAS = self.validaValor(valor)
        return self     
    
    def OnibusContinuo(self, valor):
        self.__ONIBUSCONTINUO=self.validaValor(valor)
        return self 
    def Correios(self, valor):
        self.__CORREIOS=self.validaValor(valor)
        return self
    
    def DespesasBancarias(self, valor):
        self.__DESPESASBANCARIAS=self.validaValor(valor)
        return self
    
    def OutrosAdministrativos(self, valor):
        self.__OUTROSADMINISTRATIVOS=self.validaValor(valor)
        return self
    
    def Kombi(self, valor):
        self.__KOMBI=self.validaValor(valor)
        return self
    
    def Motoboy(self, valor):
        self.__MOTOBOY = self.validaValor(valor)
        return self 
    
    def OutrosFretes(self, valor):
        self.__OUTROSFRETES = self.validaValor(valor)
        return self 
    
    def Transportadoras(self, valor):
        self.__TRANSPORTADORAS = self.validaValor(valor)
        return self     
    
    def SedexRemessaVenda(self, valor):
        self.__SEDEXREMESSAVENDA = self.validaValor(valor)
        return self       
    
    def PedagioFretevenda(self, valor):
        self.__PEDAGIOFRETEVENDA = self.validaValor(valor)
        return self           
    
    def Simples(self, valor):
        self.__SIMPLES = self.validaValor(valor)
        return self        
    
    
    def OutrosImpostos(self, valor):
        self.__OUTROSIMPOSTOS= self.validaValor(valor)
        return self     
    
    def Salarios(self, valor):
        self.__SALARIOS = self.validaValor(valor)
        return self     
    
    def Recisoes(self, valor):
        self.__RECISOES = self.validaValor(valor)
        return self         
    
    def ValeTranspote(self, valor):
        self.__VALETRANSPORTE = self.validaValor(valor)
        return self   
    
    def ValeAlimentacao(self, valor):
        self.__VALEALIMENTACAO = self.validaValor(valor)
        return self       
    
    def Inss(self, valor):
        self.__INSS = self.validaValor(valor)
        return self        
    
    def Fgts(self, valor):
        self.__FGTS = self.validaValor(valor)
        return self
    
    def AjudaDeCusto(self, valor):
        self.__AJUDADECUSTO = self.validaValor(valor)
        return self        
   
    def Ferias(self, valor):
        self.__FERIAS = self.validaValor(valor)
        return self           
    
    def HoraExtra(self, valor):
        self.__HORAEXTRA = self.validaValor(valor)
        return self               
                              
    def Testes(self, valor):
        self.__TESTES = self.validaValor(valor)
        return self  
    
    def CestaBasica(self, valor):
        self.__CESTABASICA = self.validaValor(valor)
        return self       
    
    def Lanches(self, valor):
        self.__LANCHES = self.validaValor(valor)
        return self   
    def PesquisaeDesenvolvimento(self, valor):
        self.__PESQUISADESENVOLVIMENTO = self.validaValor(valor)
        return self       
    
    def Farmacia(self, valor):
        self.__FARMACIA = self.validaValor(valor)
        return self    
    
    def Treinamento(self, valor):
        self.__TREINAMENTO = self.validaValor(valor)
        return self       
           
    def Biscate(self, valor):
        self.__BISCATE = self.validaValor(valor)
        return self   
        
    def Outros(self, valor):
        self.__OUTROS = self.validaValor(valor)
        return self  
    
    def EmprestimosInternos(self, valor):
        self.__EMPRESTIMOSINTERNOS = self.validaValor(valor)
        return self

    def Resultados(self, valor):
        self.__RESULTADOS = self.validaValor(valor)
        return self                   

    def ProLaboreDiretoria(self, valor):
        self.__PROLABOREDIRETORIA = self.validaValor(valor)
        return self   
    
    def ProLaboreGerencia(self, valor):
        self.__PROLABOREGERENCIA = self.validaValor(valor)
        return self       

    def DespesasCadastraisClientes(self, valor):
        self.__DESPCADASTRAISCLIENTES = self.validaValor(valor)
        return self
    
    def Exames(self, valor):
        self.__EXAMES = self.validaValor(valor)
        return self  
    
    def Cipa(self, valor):
        self.__CIPA = self.validaValor(valor)
        return self     
    
    def ServSeguranca(self, valor):
        self.__SERVSEGURANCA = self.validaValor(valor)
        return self         
    
    def MInformatica(self, valor):
        self.__MINFORMATICA = self.validaValor(valor)
        return self     
    
    def PlanoDeSaude(self, valor):
        self.__PLANODESAUDE = self.validaValor(valor)
        return self    
    
    def FeriasTrabalhadas(self, valor):
        self.__FERIASTRABALHADAS = self.validaValor(valor)
        return self       
    
    def EpiUniformes(self, valor):
        self.__EPIUNIFORMES = self.validaValor(valor)
        return self           
          
    def FestaseConfraternizacoes(self, valor):
        self.__FESTASCONFRATERNIZACOES = self.validaValor(valor)
        return self        
    
    def MPredial(self, valor):
        self.__MPREDIAL = self.validaValor(valor)
        return self     
    
    def MEletrica(self, valor):
        self.__MELETRICA = self.validaValor(valor)
        return self  
    
    def MArCondicionado(self, valor):
        self.__MARCONDICIONADO = self.validaValor(valor)
        return self
    
    def MBebedouro(self, valor):
        self.__MBEBEDOURO = self.validaValor(valor)
        return self      
    
    def MMoveis(self, valor):
        self.__MMOVEIS = self.validaValor(valor)
        return self                               
    
    def MOutrosManutencao(self, valor):
        self.__MOUTROSMANUT = self.validaValor(valor)
        return self      

    def MHidraulica(self, valor):
        self.__MHIDRAULICA = self.validaValor(valor)
        return self  
    
    def BMaquinasEquip(self, valor):
        self.__BMAQUINASEEQUIP = self.validaValor(valor)
        return self  
    
    def ObrasEMelhorias(self, valor):
        self.__OBRASEMELHORIAS = self.validaValor(valor)
        return self  
    
    def BMoveis(self, valor):
        self.__BMOVEIS = self.validaValor(valor)
        return self  
    
    def BInformaticaEHardware(self, valor):
        self.__BINFORMATICAEHARD = self.validaValor(valor)
        return self  
    
    def BFerramentas(self, valor):
        self.__BFERRAMENTAS = self.validaValor(valor)
        return self  
    
    def AjudaCombustivel(self, valor):
        self.__AJUDACOMBUSTIVEL = self.validaValor(valor)
        return self  
    
    def DespRepresentacao(self, valor):
        self.__DESPREPRESENTACAO = self.validaValor(valor)
        return self  
    
    def Propaganda(self, valor):
        self.__PROPAGANDA = self.validaValor(valor)
        return self  
    
    def Brindes(self, valor):
        self.__BRINDES = self.validaValor(valor)
        return self  
    
    def BInformaticaSoft(self, valor):
        self.__BINFORMATICASOFT = self.validaValor(valor)
        return self 
    
    def IMaquinasEquips(self, valor):
        self.__IMAQUINASEQUIPS = self.validaValor(valor)
        return self     
    
    def IComputadoresHardware(self, valor):
        self.__ICOMPUTADORESHARD = self.validaValor(valor)
        return self  
    
    def IOutrosInvestimentos(self, valor):
        self.__IOUTROSINVESTIMENTOS = self.validaValor(valor)
        return self              
    
    def PremioseGratificacoes(self, valor):
        self.__PREMIOSEGRATIFICACOES = self.validaValor(valor)
        return self      
    
    def DevolucaoVenda(self, valor):
        self.__DEVOLUCAOVENDA = self.validaValor(valor)
        return self  

    def Irrf(self, valor):
        self.__IRRF = self.validaValor(valor)
        return self
    
    def Iof(self, valor):
        self.__IOF = self.validaValor(valor)
        return self    
    
    def JurosSEmprestimos(self, valor):
        self.__JUROSEMPRESTIMOS = self.validaValor(valor)
        return self    
    
    def LicenciamentoseCertidoes(self, valor):
        self.__LICENCIAMENTOS_E_CERTIDOES = self.validaValor(valor)
        return self                    
    
        
    
   
    
    
  
    
     
                                
    def getNomeMes(self):
        return self.__nomeMes
    
    def getDesp(self):
        return self.__Desp    

    def getNomeDesp(self):
        return self.__nomeDesp
    
    def getdata(self):
        return self.__data  
    
    def getcredor(self):
        return self.__credor  
    
    def getdiscrimina(self):
        return self.__discrimina     
    
    def getvalor(self):
        return  self.__valor       
    
    def getAluguel(self):
        return self.__ALUGUEL

    def getIptu(self):
        return self.__IPTU
    
    def getCedae(self):
        return self.__CEDAE                  

    def getLight(self):
        return self.__LIGHT  
              
    def getCeg(self):
        return self.__CEG         

    def getTelefonia(self):
        return self.__TELEFONIA       
    
    def getPnp(self):
        return self.__PNP       
    
    def getContabilidade(self):
        return self.__CONTABILIDADE         
    
    def getSigraf(self):
        return self.__SIGRAF        
    
    def getInformatica(self):
        return self.__INFORMTATICAECONTRATOS    
    
    def getMaterialdeEscritorio(self):
        return self.__MATERIALESCRITORIO    

    def getMaterialLimpeza(self):
        return self.__MATERIALLIMPEZA  
    
    def getCartuchosparaImpressoras(self):
        return self.__CARTUCHOSPARAIMPRESSORAS
    
    def getOnibusContinuo(self):
        return self.__ONIBUSCONTINUO
    
    def getCorreios(self):
        return self.__CORREIOS
    
    def getDespesasBancarias(self):
        return self.__DESPESASBANCARIAS
    
    def getOutrosAdminitrativos(self):
        return self.__OUTROSADMINISTRATIVOS
    
    def getKombi(self):
        return self.__KOMBI     
    
    def getMotoboy(self):
        return self.__MOTOBOY         
    
    def getOutrosFretes(self):
        return self.__OUTROSFRETES   
    
    def getTransportadoras(self):
        return self.__TRANSPORTADORAS       
    
    def getSedexRemessaVenda(self):
        return self.__SEDEXREMESSAVENDA   
    
    def getPedagioFreteVenda(self):
        return self.__PEDAGIOFRETEVENDA       
    
    def getSimples(self):
        return self.__SIMPLES   
    
    def getOutrosImpostos(self):
        return self.__OUTROSIMPOSTOS               
    
    def getSalarios(self):
        return self.__SALARIOS               
    
    def getRecisoes(self):
        return self.__RECISOES         
    
    def getValeTransporte(self):
        return self.__VALETRANSPORTE       

    def getValeAlimentacao(self):
        return self.__VALEALIMENTACAO
    
    def getiNSS(self):
        return self.__INSS
    
    def getFGTS(self):
        return self.__FGTS
    
    def getAjudaDeCusto(self):
        return self.__AJUDADECUSTO

    def getFerias(self):
        return self.__FERIAS
    
    def getHoraExtra(self):
        return self.__HORAEXTRA
    
    def getTestes(self):
        return self.__TESTES
    
    def getCestaBasica(self):
        return self.__CESTABASICA
    
    def getLanches(self):
        return self.__LANCHES
    
    def getPesquisaeDesenvolvimento(self):
        return self.__PESQUISADESENVOLVIMENTO    

    def getFarmacia(self):
        return self.__FARMACIA
    
    def getTreinamento(self):
        return self.__TREINAMENTO
    
    def getBiscate(self):
        return self.__BISCATE
    
    def getOutros(self):
        return self.__OUTROS
    
    def getEmprestimosInternos(self):
        return self.__EMPRESTIMOSINTERNOS
    
    def getResultados(self):
        return self.__RESULTADOS
    
    def getProLaboreDiretoria(self):
        return self.__PROLABOREDIRETORIA
    
    def getProLaboreGerencia(self):
        return self.__PROLABOREGERENCIA

    def getDespesasCadastraisClientes(self):
        return self.__DESPCADASTRAISCLIENTES
    
    def getExames(self):
        return self.__EXAMES
    
    def getCipa(self):
        return self.__CIPA            
    
    def getServSeguranca(self):
        return self.__SERVSEGURANCA        
    
    def getMInformatica(self):
        return self.__MINFORMATICA   
    
    def getPlanoDeSaude(self):
        return self.__PLANODESAUDE    
    
    def getFeriasTrabalhadas(self):
        return self.__FERIASTRABALHADAS       
    
    def getEpiUniformes(self):
        return self.__EPIUNIFORMES           
          
    def getFestaseConfraternizacoes(self):
        return self.__FESTASCONFRATERNIZACOES        
    
    def getMPredial(self):
        return self.__MPREDIAL                    
        
    
    def getMEletrica(self):
        return self.__MELETRICA
    
    def getMArCondicionado(self):
        return self.__MARCONDICIONADO  
    
    def getMBebedouro(self):
        return self.__MBEBEDOURO      
    
    def getMMoveis(self):
        return self.__MMOVEIS                                
    
    def getMOutrosManutencao(self):
        return self.__MOUTROSMANUT 
              

    def getMHidraulica(self):
        return self.__MHIDRAULICA 
    
    def getBMaquinasEquip(self):
        return self.__BMAQUINASEEQUIP   
    
    def getObrasEMelhorias(self):
        return self.__OBRASEMELHORIAS 
    
    def getBMoveis(self):
        return self.__BMOVEIS
    
    def getBInformaticaEHardware(self):
        return self.__BINFORMATICAEHARD 
    
    def getBFerramentas(self):
        return  self.__BFERRAMENTAS       
    
    def getAjudaCombustivel(self):
        return self.__AJUDACOMBUSTIVEL 
    
    def getDespRepresentacao(self):
        return self.__DESPREPRESENTACAO 
    
    def getPropaganda(self):
        return self.__PROPAGANDA 

    
    def getBrindes(self):
        return self.__BRINDES
  
    def getBInformaticaSoft(self):
        return self.__BINFORMATICASOFT 
        
    
    def getIMaquinasEquips(self):
        return self.__IMAQUINASEQUIPS
        
    
    def getIComputadoresHardware(self):
        return self.__ICOMPUTADORESHARD 
  
    
    def getIOutrosInvestimentos(self):
        return self.__IOUTROSINVESTIMENTOS 
                      
    
    def getPremioseGratificacoes(self):
        return self.__PREMIOSEGRATIFICACOES 
              
    
    def getDevolucaoVenda(self):
        return self.__DEVOLUCAOVENDA 
        

    def getIrrf(self):
        return self.__IRRF 

    
    def getIof(self):
        return self.__IOF 
    
    
    def getJurosSEmprestimos(self):
        return self.__JUROSEMPRESTIMOS 
    
    def getLicenciamentoseCertidoes(self):
        return self.__LICENCIAMENTOS_E_CERTIDOES      

  
                   
        
    
    
    
                                
        
    def validaValor(self, valor):
        if (valor is not None and valor!=''and valor!=dict):
            return round(float(valor),2)
        else:
            return round(float(0),2)            