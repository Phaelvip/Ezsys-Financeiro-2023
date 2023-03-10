class RelatorioVenda:
    
    def __init__(self):
        self.__total=0
        self.__nomeMes=""
        self.__VENDA=""
        self.__VENDA_C_N_VENDAS = ""
        self.__VENDA_C_N_SERVICOS = ""
        self.__VENDAS_S_N = ""
        self.__TOTAL_VENDAS  = ""
        self.__SIMPLES = ""
        self.__TOTAL_VENDAS_LIQ = ""
        self.__VENDAS_DEVOLVIDAS = ""
        self.__VENDAS_CANCELADAS = ""
        self.__CORTESIAS = ""
        self.__DESC_ERRO_OPERACIONAL = ""
        self.__DESC_ERRO_COMERCIAL = ""
        self.__COMISSÕES_INTERNAS = ""
        self.__COMISSÕES_EXTERNAS = ""
        self.__TOTAL_GERAL=""
        self.__Despesas=""

        self.__PRO_LABORE_GERENCIA = ""
        
        self.__LAMINAS = ""
        self.__MADEIRAS = ""
        self.__VAZADORES = ""
        self.__BORRACHAS = ""
        self.__PERTINAX = ""
        self.__MATERIAL_P_DESTAQUE = ""
        self.__MAQUINAS_PARA_REVENDA = ""
        
        self.__OUTROS = ""
        self.__SERVIÇOS_EXTERNOS = ""
        self.__DESPESAS_COM_ERROS = ""
        
        self.__LAMINAS_CLIENTES = ""
        self.__MADEIRAS_CLIENTES = ""
        self.__VAZADORES_CLIENTES = ""
        self.__BORRACHAS_CLIENTES = ""
        self.__PERTINAX_CLIENTES = ""
        self.__MATERIAL_P_DESTAQUE_CLIENTES = ""        
        self.__CLICHES = ""        
       
       
       
        self.__DESENHO = ""
        self.__SALARIOS_DESENHO = ""
        self.__H_E = ""
        self.__TESTES = ""
        self.__BISCATES = ""
        self.__FÉRIAS_TRABALHADAS  = ""        
        self.__USO_E_CONSUMO_DESENHO = ""
        self.__MANUTENÇAo=""
        
        self.__PLOTTER=""
        self.__USO_E_CONSUMO_PLOTTER=""
        self.__MANUTENÇÃO_PLOTTER=""


        self.__LASER=""
        self.__SALARIOS_LASER = ""
        self.__H_E_LASER = ""
        self.__TESTES_LASER = ""
        self.__BISCATES_LASER = ""
        self.__FÉRIAS_TRABALHADAS_LASER  = ""            
        self.__USO_E_CONSUMO_LASER=""
        self.__MANuTENÇÃOLASER=""
        self.__GAS_P_LASER=""
        
        self.__MONTAGEM=""
        self.__SALARIOS_MONTAGEM = ""
        self.__H_E_MONTAGEM = ""
        self.__TESTES_MONTAGEM = ""
        self.__BISCATES_MONTAGEM = ""
        self.__FÉRIAS_TRABALHADAS_MONTAGEM  = ""              
        self.__UsO_E_CONSUMO_MONTAGEM=""
        self.__MANUTENÇÃO_MONTAGEM="" 
               
        self.__BORRACHA=""        
        self.__SALARIOS_BORRACHA = ""
        self.__H_E_BORRACHA = ""
        self.__TESTES_BORRACHA = ""
        self.__BISCATES_BORRACHA = ""
        self.__FÉRIAS_TRABALHADAS_BORRACHA  = ""             
        self.__USO_E_CONSUMO_BORRACHA  =""
        self.__MANUTENÇÃO_BORRACHA =""
        
        self.__TOTAL_GERAL_SETORES=""
        self.__TOTAL_DE_RESULTADOS_1=""
        self.__DEPRECIAÇÃO=""
        self.__TOTAL_DE_RESULTADOS_2=""
        self.__DESPESAS_ADMINISTRATIVAS=""
        self.__DESPESAS_CENTRAIS="" 
        self.__SALARIOS_ADM=""   
        self.__SALARIO_COMERCIAL=""
        self.__SALÁRIOS_TRANSPORTE=""
        self.__TOTAL_GERAL_ADMINISTRAIVO=""
        self.__DESPESAS_FINANCEIRAS=""
        self.__JUROS=""
        self.__JUROS_SOBRE_APLICACOES=""
        self.__JUROS_SOBRE_VENDAS=""        
        
        self.__TOTAL_RESULTADO_EZIPA=""
        self.__WOODPAPER=""
        self.__TOTAL_RESULTADO_EZIPA_E_WOODPAPER=""

        
    def build(self):
        return self

    #def LinhasIndice(self, valor):
        self.LinhasIndice = valor
        return self    
    
    def nomeMes(self, valor):
        self.__nomeMes = valor
        return self
    
    def Vendas(self, valor):
        self.__VENDA = valor
        return self
    
    def vendasCNVendas(self, valor):
        self.__VENDA_C_N_VENDAS = self.validaValor(valor)
        return self
    
    def vendasCNServicos(self, valor):
        self.__VENDA_C_N_SERVICOS = self.validaValor(valor)
        return self
    
    def vendas_S_N(self, valor):
        self.__VENDAS_S_N = self.validaValor(valor)
        return self
    
    def Simples(self, valor):
        self.__SIMPLES = self.validaValor(valor)
        return self    
    
    
    def Vendas_Devolvidas(self, valor):
        self.__VENDAS_DEVOLVIDAS = self.validaValor(valor)
        return self         
   
    def Vendas_Canceladas(self, valor):
        self.__VENDAS_CANCELADAS = self.validaValor(valor)
        return self        
   
    def Cortesias(self, valor):
        self.__CORTESIAS = self.validaValor(valor)
        return self          
  
    def DescErroOperacional(self, valor):
        self.__DESC_ERRO_OPERACIONAL = self.validaValor(valor)
        return self          
  
    def DescErroComercial(self, valor):
        self.__DESC_ERRO_COMERCIAL = self.validaValor(valor)
        return self   
  
    def ComissoesInternas(self, valor):
        self.__COMISSÕES_INTERNAS = self.validaValor(valor)
        return self   
   
    def ComissoesExternas(self, valor):
        self.__COMISSÕES_EXTERNAS = self.validaValor(valor)
        return self      

    def ProLaboreGerencia(self, valor):
        self.__PRO_LABORE_GERENCIA = self.validaValor(valor)
        return self     

    def getProLaboreGerencia(self):
        return self.__PRO_LABORE_GERENCIA      
#_____________________________________________________________________ESTOQUE_________________________________________________________________________
    
    def Laminas(self, valor):
        self.__LAMINAS = self.validaValor(valor)
        return self 
   
    def Madeiras(self, valor):
        self.__MADEIRAS = self.validaValor(valor)
        return self  
   
    def Vazadores(self, valor):
        self.__VAZADORES = self.validaValor(valor)
        return self  
   
    def Pertinax(self, valor):
        self.__PERTINAX = self.validaValor(valor)
        return self  
  
    def Borrachas(self, valor):
        self.__BORRACHAS = self.validaValor(valor)
        return self  
 
    def MaterialPDestaque(self, valor):
        self.__MATERIAL_P_DESTAQUE = self.validaValor(valor)
        return self                            
  
    def MaquinasParaRevenda(self, valor):
        self.__MAQUINAS_PARA_REVENDA = self.validaValor(valor)
        return self   
      
    def Outros(self, valor):
        self.__OUTROS = self.validaValor(valor)
        return self         
  
    def ServicosExternos(self, valor):
        self.__SERVIÇOS_EXTERNOS = self.validaValor(valor)
        return self    
     
    def DespesacomErros(self, valor):
        self.__DESPESAS_COM_ERROS = self.validaValor(valor)
        return self

 
    def getLaminas(self):
        return self.__LAMINAS
    
    def getMadeiras(self):
        return self.__MADEIRAS
    
    def getVazadores(self):
        return self.__VAZADORES
    
    def getBorrachas(self):
        return self.__BORRACHAS
    
    def getPertinax(self):
        return self.__PERTINAX
    
    def getMaterialPDestaque(self):
        return self.__MATERIAL_P_DESTAQUE
    
    def getMaquinasParaRevenda(self):
        return self.__MAQUINAS_PARA_REVENDA 
    
    def getOutros(self):
        return self.__OUTROS
    
    def getServicosExternos(self):
        return self.__SERVIÇOS_EXTERNOS
        
    def getDespesacomErros(self):
        return self.__DESPESAS_COM_ERROS    

    
    
   #___________________________________________________CLIENTES___________________________________________________
    def LaminasClientes(self, valor):
        self.__LAMINAS_CLIENTES = self.validaValor(valor)
        return self 
    
    def MadeirasClientes(self, valor):
        self.__MADEIRAS_CLIENTES = self.validaValor(valor)
        return self  

    def VazadoresClientes(self, valor):
        self.__VAZADORES_CLIENTES = self.validaValor(valor)
        return self  

    def PertinaxClientes(self, valor):
        self.__PERTINAX_CLIENTES = self.validaValor(valor)
        return self  

    def BorrachasClientes(self, valor):
        self.__BORRACHAS_CLIENTES = self.validaValor(valor)
        return self  

    def MaterialPDestaqueClientes(self, valor):
        self.__MATERIAL_P_DESTAQUE_CLIENTES = self.validaValor(valor)
        return self       
    
    def Cliches(self, valor):
        self.__CLICHES = self.validaValor(valor)
        return self          
    

    def getLaminasClientes(self):
        return self.__LAMINAS_CLIENTES 
    
    def getMadeirasClientes(self):
        return self.__MADEIRAS_CLIENTES
    
    def getVazadoresClientes(self):
        return self.__VAZADORES_CLIENTES
    
    def getBorrachasClientes(self):
        return self.__BORRACHAS_CLIENTES
    
    def getPertinaxClientes(self):
        return self.__PERTINAX_CLIENTES
    
    def getMaterialPDestaqueClientes(self):
        return self.__MATERIAL_P_DESTAQUE_CLIENTES        

    def getCliches(self):
        return self.__CLICHES       
    
   #___________________________________________________DESENHO____________________________________________________
    def Desenho(self,valor):
        self.__DESENHO = valor
        return self
    
    def SalariosDesenho(self, valor):
        self.__SALARIOS_DESENHO = self.validaValor(valor)
        return self 
       
    def H_e(self, valor):
        self.__H_E = self.validaValor(valor)
        return self     
    
    def testes(self, valor):
        self.__TESTES = self.validaValor(valor)
        return self 
    
    def biscates(self, valor):
        self.__BISCATES = self.validaValor(valor)
        return self 
    
    def FeriasTrabalhadas(self, valor):
        self.__FÉRIAS_TRABALHADAS = self.validaValor(valor)
        return self      
       
    def usoeConsumoDesenho(self, valor):
        self.__USO_E_CONSUMO_DESENHO = self.validaValor(valor)
        return self 
    
    def manutencao(self, valor):
        self.__MANUTENÇAo = self.validaValor(valor)
        return self   

    
    def getDesenho(self):
        return self.__DESENHO 
    
    def getSalariosDesenho(self):
        return self.__SALARIOS_DESENHO
    
    def getH_E(self):
        return self.__H_E   
         
    def getTestes(self):
        return self.__TESTES    
    
    def getBiscates(self):
        return self.__BISCATES    
    
    def getFeriasTrabalhadas(self):
        return self.__FÉRIAS_TRABALHADAS  
    
    def getusoeConsumoDesenho(self):
        return self.__USO_E_CONSUMO_DESENHO
    
    def getmanutencao(self):
        return self.__MANUTENÇAo   
    
    #_____________________________________________________________________PLOTTER________________________________________________________________ 
    def Plotter(self,valor):
        self.__PLOTTER = valor
        return self   
   
    def UsoeConsumoPlotter(self, valor):
        self.__USO_E_CONSUMO_PLOTTER = self.validaValor(valor)
        return self 
 
    def ManutencaoPlotter(self, valor):
        self.__MANUTENÇÃO_PLOTTER  = self.validaValor(valor)
        return self         


    def getPlotter(self):
        return self.__PLOTTER    
    
    def getUsoeConsumoPlotter(self):
        return self.__USO_E_CONSUMO_PLOTTER 
    
    def getManutencaoPlotter(self):
        return self.__MANUTENÇÃO_PLOTTER 
 #________________________________________________________________________Laser___________________________________________________________________  
    def Laser(self,valor):
        self.__LASER = valor
        return self  
    
    def SalariosLaser(self, valor):
        self.__SALARIOS_LASER = self.validaValor(valor)
        return self 
       
    def H_eLaser(self, valor):
        self.__H_E_LASER = self.validaValor(valor)
        return self     
    
    def testesLaser(self, valor):
        self.__TESTES_LASER = self.validaValor(valor)
        return self 
    
    def biscatesLaser(self, valor):
        self.__BISCATES_LASER = self.validaValor(valor)
        return self 
    
    def FeriasTrabalhadasLaser(self, valor):
        self.__FÉRIAS_TRABALHADAS_LASER = self.validaValor(valor)
        return self      
    
    def UsoeConsumoLaser(self, valor):
        self.__USO_E_CONSUMO_LASER = self.validaValor(valor)
        return self 
   
    def ManUtencaoLaser(self, valor):
        self.__MANuTENÇÃOLASER = self.validaValor(valor)
        return self 
 
    def GaspLaser(self,valor):
        self.__GAS_P_LASER = self.validaValor(valor)
        return self  


    def getLaser(self):
        return self.__LASER              
      
    def getSalariosLaser(self):
        return self.__SALARIOS_LASER
    
    def getH_ELaser(self):
        return self.__H_E_LASER   
         
    def getTestesLaser(self):
        return self.__TESTES_LASER    
    
    def getBiscatesLaser(self):
        return self.__BISCATES_LASER    
    
    def getFeriasTrabalhadasLaser(self):
        return self.__FÉRIAS_TRABALHADAS_LASER  

    def getUsoeConsumoLaser(self):
        return self.__USO_E_CONSUMO_LASER
    
    def getManUtencaoLaser(self):
        return self.__MANuTENÇÃOLASER       
    
    def getGaspLaser(self):
        return self.__GAS_P_LASER      
    #_______________________________________________________MONTAGEM_____________________________________________________________________________
    
    def Montagem(self,valor):
        self.__MONTAGEM = valor
        return self   
    
    def SalariosMontagem(self, valor):
        self.__SALARIOS_MONTAGEM = self.validaValor(valor)
        return self 
       
    def H_eMontagem(self, valor):
        self.__H_E_MONTAGEM = self.validaValor(valor)
        return self     
    
    def testesMontagem(self, valor):
        self.__TESTES_MONTAGEM = self.validaValor(valor)
        return self 
    
    def biscatesMontagem(self, valor):
        self.__BISCATES_MONTAGEM = self.validaValor(valor)
        return self 
    
    def FeriasTrabalhadasMontagem(self, valor):
        self.__FÉRIAS_TRABALHADAS_MONTAGEM = self.validaValor(valor)
        return self            
   
    def usoEconsumoMontagem(self, valor):
        self.__UsO_E_CONSUMO_MONTAGEM = self.validaValor(valor)
        return self 
   
    def ManutencaoMontagem(self, valor):
        self.__MANUTENÇÃO_MONTAGEM = self.validaValor(valor)
        return self  
    
    
    def getMontagem(self):
        return self.__MONTAGEM          
    
    def getSalariosMontagem(self):
        return self.__SALARIOS_MONTAGEM
    
    def getH_eMontagem(self):
        return self.__H_E_MONTAGEM   
         
    def getTestesMontagem(self):
        return self.__TESTES_MONTAGEM    
    
    def getbiscatesMontagem(self):
        return self.__BISCATES_MONTAGEM
    
    def getFeriasTrabalhadasMontagem(self):
        return self.__FÉRIAS_TRABALHADAS_MONTAGEM
    
    def getUsoeConsumoMontagem(self):
        return self.__UsO_E_CONSUMO_MONTAGEM
    
    def getManuTencaoMontagem(self):
        return self.__MANUTENÇÃO_MONTAGEM    

#________________________________________________________________Borracha_____________________________________________________________       
    def Borracha(self,valor):
        self.__MONTAGEM_AUTOMATICA = valor
        return self     
   
    def SalariosBorracha(self, valor):
        self.__SALARIOS_BORRACHA = self.validaValor(valor)
        return self 
       
    def H_eBorracha(self, valor):
        self.__H_E_BORRACHA = self.validaValor(valor)
        return self     
    
    def testesBorracha(self, valor):
        self.__TESTES_BORRACHA = self.validaValor(valor)
        return self 
    
    def biscatesBorracha(self, valor):
        self.__BISCATES_BORRACHA = self.validaValor(valor)
        return self 
    
    def FeriasTrabalhadasBorracha(self, valor):
        self.__FÉRIAS_TRABALHADAS_BORRACHA = self.validaValor(valor)
        return self        
    
    def UsoeConsumoBorracha(self, valor):
        self.__USO_E_CONSUMO_BORRACHA = self.validaValor(valor)
        return self 
   
    def ManutencaoBorracha(self, valor):
        self.__MANUTENÇÃO_BORRACHA = self.validaValor(valor)
        return self    


    def getBorracha(self):
        return self.__BORRACHA              
    
    def getSalariosBorracha(self):
        return self.__SALARIOS_BORRACHA
    
    def getH_eBorracha(self):
        return self.__H_E_BORRACHA   
         
    def getTestesBorracha(self):
        return self.__TESTES_BORRACHA    
    
    def getbiscatesBorracha(self):
        return self.__BISCATES_BORRACHA
    
    def getFeriasTrabalhadasBorracha(self):
        return self.__FÉRIAS_TRABALHADAS_BORRACHA

    def getUsoeConsumoBorracha(self):
        return self.__USO_E_CONSUMO_BORRACHA
    
    def getManutencaoBorracha(self):
        return self.__MANUTENÇÃO_BORRACHA        
           
    #def getLinhasIndice(self):
        return self.LinhasIndice
    def DespesasAdminsitrativas(self, valor):
        self.__DESPESAS_ADMINISTRATIVAS = self.validaValor(valor)
        return self 
    
    def DespesasCentrais(self, valor):
        self.__DESPESAS_CENTRAIS = self.validaValor(valor)
        return self     
  
    def SalariosAdm(self, valor):
        self.__SALARIOS_ADM = self.validaValor(valor)
        return self          
   
    def SalariosComercial(self, valor):
        self.__SALARIO_COMERCIAL = self.validaValor(valor)
        return self      
  
    def SalariosTransporte(self, valor):
        self.__SALÁRIOS_TRANSPORTE = self.validaValor(valor)
        return self  

    def DespesasFinanceiras(self, valor):
        self.__DESPESAS_FINANCEIRAS = self.validaValor(valor)
        return self    
   
    def juros(self, valor):
        self.__JUROS = self.validaValor(valor)
        return self    
    
    def JurosSobreAplicacoces(self, valor):
        self.__JUROS_SOBRE_APLICACOES = self.validaValor(valor)
        return self       
    
    def JurosSobreVendas(self, valor):
        self.__JUROS_SOBRE_VENDAS = self.validaValor(valor)
        return self              
            
            
    def getNomeMes(self):
        return self.__nomeMes
    
    def getVendas(self):
        return self.__VENDA 

    def getvendasCNVendas(self):
        return self.__VENDA_C_N_VENDAS    
    

    
    def getVendasCNServicos(self):
        return self.__VENDA_C_N_SERVICOS
    
    def getVendas_S_N(self):
        return self.__VENDAS_S_N
    
    def TotalVendas(self):
        valor = float(self.__VENDA_C_N_VENDAS) + float(self.__VENDA_C_N_SERVICOS)+ float(self.__VENDAS_S_N)
        self.__TOTAL_VENDAS=self.validaValor(valor)
        return self
    
    def getTotalVendas(self):
        return self.__TOTAL_VENDAS
    
    def gerarTotal(self):
        valor = (self.validaValor(self.__VENDA_C_N_VENDAS))+ (self.validaValor(self.__VENDA_C_N_SERVICOS))+ (self.validaValor(self.__VENDAS_S_N))
        self.__total=valor
        
    def gerarTotalVendas(self):
        valor = float(self.__VENDA_C_N_VENDAS) + float(self.__VENDA_C_N_SERVICOS)+ float(self.__VENDAS_S_N)
        self.__total=(valor)
        return self

    
    def getSimples(self):
        return self.__SIMPLES
    
    def TotalLiq(self):
        valor = float(self.__TOTAL_VENDAS)- float(self.validaValor(self.__SIMPLES))
        self.__TOTAL_VENDAS_LIQ=self.validaValor(valor)
        return self
   
    def getTotalLiq(self):
        return self.__TOTAL_VENDAS_LIQ     
    
    def getVendas_Devolvidas(self):
        return self.__VENDAS_DEVOLVIDAS 
    
    def getVendas_Canceladas(self):
        return self.__VENDAS_CANCELADAS
    
    def getCortesias(self):
        return self.__CORTESIAS

    def getDescErroOperacional(self):
        return self.__DESC_ERRO_OPERACIONAL    
    
    def getDescErroComercial(self):
        return self.__DESC_ERRO_COMERCIAL  
    
    def getComissoesInternas(self):
        return self.__COMISSÕES_INTERNAS   
    
    def getComissoesExternas(self):
        return self.__COMISSÕES_EXTERNAS
    
    def totalVendasGeral(self):
        valor = float(self.__TOTAL_VENDAS_LIQ)- float(self.validaValor(self.__VENDAS_DEVOLVIDAS))- float(self.validaValor(self.__VENDAS_CANCELADAS))- float(self.validaValor(self.__CORTESIAS))- float(self.validaValor(self.__DESC_ERRO_COMERCIAL))- float(self.validaValor(self.__DESC_ERRO_OPERACIONAL))-float(self.validaValor(self.__COMISSÕES_INTERNAS))-float(self.validaValor(self.__COMISSÕES_EXTERNAS))
        self.__TOTAL_GERAL=self.validaValor(valor)
        return self

    def getTotalVendasGeral(self):
        return self.__TOTAL_GERAL
    
    def totalGeralServiços(self):
        valor = float(self.__SALARIOS_DESENHO)+float(self.__H_E)+float(self.__TESTES)+float(self.__BISCATES)+float(self.__FÉRIAS_TRABALHADAS)+float(self.__USO_E_CONSUMO_DESENHO)+float(self.__MANUTENÇAo)+float(self.__PRO_LABORE_GERENCIA)+float(self.__LAMINAS)+float(self.__MADEIRAS)+float(self.__VAZADORES)+float(self.__BORRACHAS)+float(self.__PERTINAX)+float(self.__MATERIAL_P_DESTAQUE)+float(self.__MAQUINAS_PARA_REVENDA)+float(self.__OUTROS)+float(self.__SERVIÇOS_EXTERNOS)+float(self.__DESPESAS_COM_ERROS)+float(self.__LAMINAS_CLIENTES)+float(self.__MADEIRAS_CLIENTES)+float(self.__VAZADORES_CLIENTES)+float(self.__BORRACHAS_CLIENTES)+float(self.__PERTINAX_CLIENTES)+float(self.__MATERIAL_P_DESTAQUE_CLIENTES)+float(self.__CLICHES)+float(self.__USO_E_CONSUMO_PLOTTER)+float(self.__MANUTENÇÃO_PLOTTER)+float(self.__SALARIOS_LASER)+float(self.__H_E_LASER)+float(self.__TESTES_LASER)+float(self.__BISCATES_LASER)+float(self.__FÉRIAS_TRABALHADAS_LASER)+float(self.__USO_E_CONSUMO_LASER)+float(self.__MANuTENÇÃOLASER)+float(self.__GAS_P_LASER)+float(self.__SALARIOS_MONTAGEM)+float(self.__H_E_MONTAGEM)+float(self.__TESTES_MONTAGEM)+float(self.__BISCATES_MONTAGEM)+float(self.__FÉRIAS_TRABALHADAS_MONTAGEM)+float(self.__UsO_E_CONSUMO_MONTAGEM)+float(self.__MANUTENÇÃO_MONTAGEM)+float(self.__SALARIOS_BORRACHA)+float(self.__H_E_BORRACHA)+float(self.__TESTES_BORRACHA)+float(self.__BISCATES_BORRACHA)+float(self.__FÉRIAS_TRABALHADAS_BORRACHA)+float(self.__USO_E_CONSUMO_BORRACHA)+float(self.__MANUTENÇÃO_BORRACHA)
        self.__TOTAL_GERAL_SETORES=self.validaValor(valor)
        return self
            

    def getTotalGeralServiços(self):
        return self.__TOTAL_GERAL_SETORES
    
    def totaldeResultados1(self):
        valor = float(self.__TOTAL_GERAL)-float(self.__TOTAL_GERAL_SETORES)
        self.__TOTAL_DE_RESULTADOS_1=self.validaValor(valor)
        return self
    
    def getTotaldeResultados1(self):
        return self.__TOTAL_DE_RESULTADOS_1

    
    def Depreciacao(self,valor):
        self.__DEPRECIAÇÃO=self.validaValor(valor)
        return self
    
    def getDepreciacao(self):
        return self.__DEPRECIAÇÃO   
    
    def gerartotalWoodpaper(self,valor):
        self.__WOODPAPER=self.validaValor(valor)
        return self    
    
    def gettotalWoodpaper(self):
        return self.__WOODPAPER    
    
    
    def totaldeResultados2(self):
        valor = float(self.__TOTAL_DE_RESULTADOS_1) + float(self.__DEPRECIAÇÃO)
        self.__TOTAL_DE_RESULTADOS_2=self.validaValor(valor)
        return self
    
    def getTotaldeResultados2(self):
        return self.__TOTAL_DE_RESULTADOS_2
    
    def getDespesasAdminsitrativas(self):
        return self.__DESPESAS_ADMINISTRATIVAS

    def getDespesasCentrais(self):
        return self.__DESPESAS_CENTRAIS  

   
    def getSalariosAdm(self):
        return self.__SALARIOS_ADM
   
    def getSalariosComercial(self):
        return self.__SALARIO_COMERCIAL   
  
    def getSalariosTransporte(self):
        return self.__SALÁRIOS_TRANSPORTE
   
    def totalGeralAdministrativo(self):
        valor = float(self.__DESPESAS_ADMINISTRATIVAS)+float(self.__DESPESAS_CENTRAIS)+ float(self.__SALARIOS_ADM)+ float(self.__SALARIO_COMERCIAL  )+ float(self.__SALÁRIOS_TRANSPORTE)
        self.__TOTAL_GERAL_ADMINISTRAIVO=self.validaValor(valor)
        return self
    
    def gettotalGeralAdministrativo(self):
        return self.__TOTAL_GERAL_ADMINISTRAIVO
    
            
 
   
    def getDespesasFinanceiras(self):
        return self.__DESPESAS_FINANCEIRAS   
    
    def getJuros(self):
        return self.__JUROS      
   
    def getJurosSobreAplicacoes(self):
        return self.__JUROS_SOBRE_APLICACOES
  
    def getJurosSobreVendas(self):
        return self.__JUROS_SOBRE_VENDAS               
    
    def totalResultadoEzipa(self):
        valor = float(self.__TOTAL_DE_RESULTADOS_2)-float(self.__TOTAL_GERAL_ADMINISTRAIVO)-float(self.__DESPESAS_FINANCEIRAS)+ float(self.__JUROS)+float(self.__JUROS_SOBRE_APLICACOES)+float(self.__JUROS_SOBRE_VENDAS)
        self.__TOTAL_RESULTADO_EZIPA=self.validaValor(valor)
        return self
    
    def gettotalResultadoEzipa(self):
        return self.__TOTAL_RESULTADO_EZIPA
    

    def gerartotalEzipaeWoodpaper(self):
        valor = float(self.__TOTAL_RESULTADO_EZIPA) + float(self.__WOODPAPER)        
        self.__TOTAL_RESULTADO_EZIPA_E_WOODPAPER=self.validaValor(valor)
        return self    
            
    def gettotalEzipaeWoodpaper(self):
        return self.__TOTAL_RESULTADO_EZIPA_E_WOODPAPER
      
    def validaValor(self, valor):
        if (valor is not None and valor!=''and valor!=dict):
            return round(float(valor),2)
        else:
            return round(float(0),2)
        
    
    pass