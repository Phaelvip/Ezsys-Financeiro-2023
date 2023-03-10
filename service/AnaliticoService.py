from entity.dto.RelatorioAnalitico import RelatorioAnalitico
from persistence.VendaRepository import VendaRepository
from persistence.SaidaRepository import SaidaRepository
from persistence.ScriptsRepository import ScriptsRepository
import locale
from datetime import *
from  datetime import datetime
from mysql.connector import Error

class AnaliticoService:

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

#ABA Competencia
    def GerarRelatorioDetalhadoCompetenciaAnalitico(self, ano, mesInicial,mesFinal):
        
        linhas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        for mes in range(mesInicial, mesesInteiro,1):
            self.__quantMeses+=1
            #Template
            if(mes < 10):
                periododeCompetenciaInicial = "0{}/{}".format(mes,ano)
                periododeCompetenciafinal = "0{}/{}".format(mes,ano)
            else:
                periododeCompetenciaInicial = "{}/{}".format(mes,ano)
                periododeCompetenciafinal = "{}/{}".format(mes,ano)                  
            analiticoAdm=(self.__scriptDao.analiticoAdmCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'1'))#linha1        
        return linhas
                        
    def GerarRelatorioAnaliticoCompetencia(self, ano,  mesInicial ,mesFinal):
 
        
        colunas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        for mes in range(mesInicial, mesesInteiro,1):
            self.__quantMeses+=1
            #Template
            if(mes < 10):
                periododeCompetenciaInicial = "0{}/{}".format(mes,ano)
                periododeCompetenciafinal = "0{}/{}".format(mes,ano)
            else:
                periododeCompetenciaInicial = "{}/{}".format(mes,ano)
                periododeCompetenciafinal = "{}/{}".format(mes,ano)              
            colunaDoRelatorioAnaliticoPorMes = RelatorioAnalitico() 
                 
                      
            Aluguel=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'3'))['total']#linha1
            Iptu=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'4'))['total']#linha2
            Cedae=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'5'))['total']#linha3
            light=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'6'))['total']#linha4            
            Ceg=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'7'))['total']#linha5  
            Telefonia=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'8'))['total']#linha6  
            PnP=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'9'))['total']#linha7 
            Contabilidade=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'11'))['total']#linha8                                                
            Sigraf=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'12'))['total']#linha9         
            Informatica=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'14 '))['total']#linha10    
            MaterialdeEscricorio=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'15 '))['total']#linha11   
            MaterialLimpeza=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'16'))['total']#linha12                                                                                                                
            CartuchosParaImpressoras=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'17'))['total']#linha13 
            OnibusContinuo=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'18'))['total']#linha14   
            Correios=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'19'))['total']#linha15   
            DespesasBancarias=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'20'))['total']#linha16   
            OutrosAdminitrativos=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'21'))['total']#linha17 
            Kombi=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'23'))['total']#linha18  
            Motoboy=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'24'))['total']#linha19  
            OutrosFretes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'25'))['total']#linha20  
            Transportadoras=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'26'))['total']#linha21  
            Sedex=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'27'))['total']#linha22  
            Pedagio=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'30'))['total']#linha23  
            Simples=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'31'))['total']#linha24  
            OutrosImposos=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'34'))['total']#linha25
            Salarios=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'35'))['total']#linha26            
            Recisoes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'36'))['total']#linha27 
            ValeTransporte=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'37'))['total']#linha28
            ValeAlimentacao=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'38'))['total']#linha29  
            Inss=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'39'))['total']#linha30        
            Fgts=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'40'))['total']#linha31
            AjudaDeCusto=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'41'))['total']#linha32  
            Ferias=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'42'))['total']#linha33  
            HoraExtra=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'43'))['total']#linha34  
            Testes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'44'))['total']#linha35
            CestaBasica=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'45'))['total']#linha36  
            Lanches=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'46'))['total']#linha37  
            PesquisaDesenvolvimento=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'68'))['total']#linha37              
            Farmacia=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'47'))['total']#linha38  
            Treinamento=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'48'))['total']#linha39  
            Biscate=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'49'))['total']#linha40     
            Outros=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'50'))['total']#linha41  
            EmprestimosInternos=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'71'))['total']#linha42  
            Resultados=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'72'))['total']#linha43  
            ProLaboreDiretoria=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'73'))['total']#linha44
            ProLaboreGerencia=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'74'))['total']#linha45  
            DespesasCadastraisClientes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'75'))['total']#linha46
            Exame=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'77'))['total']#linha47
            Cipa=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'78'))['total']#linha48
            ServSeguranca=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'79'))['total']##linha49
            Minformatica=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'86'))['total']#linha50
            PlanodeSaude=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'91'))['total']#linha51
            FeriasTrabalhadas=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'92'))['total']#linha52
            EpiUniformes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'93'))['total']#linha53
            FestasEConfraternizacoes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'95'))['total']#linha54
            Mpredial=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'97'))['total']#linha55                                                                                                                                                                                                                                                                                                                                                           
            Meletrica=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'98'))['total']#linha56
            Marcondicionado=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'99'))['total']#linha57
            Mbebedouro=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'100'))['total']#linha58
            Mmoveis=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'101'))['total']#linha59 
            Moutrosmanut=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'102'))['total']#linha60
            MHidrualica=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'103'))['total']#linha61
            Maquinasequip=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'104'))['total']#linha62 
            ObraseMelhorias=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'105'))['total']#linha63 
            Bmoveis=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'106'))['total']#linha64 
            BInformaaticaeHard=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'107'))['total']#linha65 
            BFerramentas=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'109'))['total']#linha66
            AjudaECombustivel=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'110'))['total']#linha67 
            DespRepresentacao=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'111'))['total']#linha68 
            Propaganda=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'112'))['total']#linha69 
            Brindes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'113'))['total']#linha70 
            Binformaticasoft=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'120'))['total']#linha71
            Imaquinasequip=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'134'))['total']#linha72
            IcomputadoresHard=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'135'))['total']#linha73
            IoutrosInvestimentos=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'140'))['total']#linha74
            PremiosEGratificacoes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'141'))['total']#linha75
            DevolucaoVenda=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'144'))['total']#linha76
            Irrf=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'148'))['total']#linha77 
            Iof=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'149'))['total']#linha78
            JurosSemprestimos=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'151'))['total']#linha79     
            LicenciamentoseCertidoes=(self.__scriptDao.BuscardespesaAnaliticoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'163'))['total']#linha79                 
            
                     
                                                                                                           
                                                                                                                                                                                                              
            colunaDoRelatorioAnaliticoPorMes.build().nomeMes(self.getNomeMesPorNumeroPeriodo(mes)).Aluguel(Aluguel).iptu(Iptu).Cedae(Cedae).Light(light).Ceg(Ceg).Telefonia(Telefonia)            
            colunaDoRelatorioAnaliticoPorMes.build().Pnp(PnP).Contabilidade(Contabilidade).Sigraf(Sigraf).Informatica(Informatica).MaterialdeEscritorio(MaterialdeEscricorio).MaterialLimpeza(MaterialLimpeza)
            colunaDoRelatorioAnaliticoPorMes.build().CartuchosparaImpressoras(CartuchosParaImpressoras).OnibusContinuo(OnibusContinuo).Correios(Correios).DespesasBancarias(DespesasBancarias).OutrosAdministrativos(OutrosAdminitrativos).Kombi(Kombi)
            colunaDoRelatorioAnaliticoPorMes.build().Motoboy(Motoboy).OutrosFretes(OutrosFretes).Transportadoras(Transportadoras).SedexRemessaVenda(Sedex).PedagioFretevenda(Pedagio).Simples(Simples)            
            colunaDoRelatorioAnaliticoPorMes.build().OutrosImpostos(OutrosImposos).Salarios(Salarios).Recisoes(Recisoes).ValeTranspote(ValeTransporte).ValeAlimentacao(ValeAlimentacao).Inss(Inss)  
            colunaDoRelatorioAnaliticoPorMes.build().Fgts(Fgts).AjudaDeCusto(AjudaDeCusto).Ferias(Ferias).HoraExtra(HoraExtra).Testes(Testes).CestaBasica(CestaBasica)  
            colunaDoRelatorioAnaliticoPorMes.build().Lanches(Lanches).PesquisaeDesenvolvimento(PesquisaDesenvolvimento).Farmacia(Farmacia).Treinamento(Treinamento).Biscate(Biscate).Outros(Outros).EmprestimosInternos(EmprestimosInternos)  
            colunaDoRelatorioAnaliticoPorMes.build().Resultados(Resultados).ProLaboreDiretoria(ProLaboreDiretoria).ProLaboreGerencia(ProLaboreGerencia).DespesasCadastraisClientes(DespesasCadastraisClientes).Exames(Exame).Cipa(Cipa)            
            colunaDoRelatorioAnaliticoPorMes.build().ServSeguranca(ServSeguranca).MInformatica(Minformatica).PlanoDeSaude(PlanodeSaude).FeriasTrabalhadas(FeriasTrabalhadas).EpiUniformes(EpiUniformes).FestaseConfraternizacoes(FestasEConfraternizacoes).MPredial(Mpredial)            
            colunaDoRelatorioAnaliticoPorMes.build().MEletrica(Meletrica).MArCondicionado(Marcondicionado).MBebedouro(Mbebedouro).MMoveis(Mmoveis).MOutrosManutencao(Moutrosmanut).MHidraulica(MHidrualica).BMaquinasEquip(Maquinasequip)            
            colunaDoRelatorioAnaliticoPorMes.build().ObrasEMelhorias(ObraseMelhorias).BMoveis(Bmoveis).BInformaticaEHardware(BInformaaticaeHard).BFerramentas(BFerramentas).AjudaCombustivel(AjudaECombustivel)
            colunaDoRelatorioAnaliticoPorMes.build().DespRepresentacao(DespRepresentacao).Propaganda(Propaganda).Brindes(Brindes).BInformaticaSoft(Binformaticasoft).IMaquinasEquips(Imaquinasequip).IComputadoresHardware(IcomputadoresHard)
            colunaDoRelatorioAnaliticoPorMes.build().IOutrosInvestimentos(IoutrosInvestimentos).PremioseGratificacoes(PremiosEGratificacoes).DevolucaoVenda(DevolucaoVenda).Irrf(Irrf).Iof(Iof).JurosSEmprestimos(JurosSemprestimos).LicenciamentoseCertidoes(LicenciamentoseCertidoes)
           
           
            colunaDoRelatorioAnaliticoPorMes.gerarTotal()            
            colunas.append(colunaDoRelatorioAnaliticoPorMes)
        self.__colunasAnaliticoGeral = colunas
        #self.GerarRelatorioAnaliticodetalhado(ano, mesFinal, mesInicial, diainicial, diafinal)
        return colunas
 
    def AnaliticoCompetenciaDetalhado(self, ano, mesInicial,mesFinal, desp):
        linhas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        #TEM QUE CONVERTER DE TEXTO PRA INTEIRO
        for mes in range(mesInicial, mesesInteiro, 1):
            self.__quantMeses+=1
            #Template
            if(mes < 10):
                periododeCompetenciaInicial = "0{}/{}".format(mes,ano)
                periododeCompetenciafinal = "0{}/{}".format(mes,ano)
            else:
                periododeCompetenciaInicial = "{}/{}".format(mes,ano)
                periododeCompetenciafinal = "{}/{}".format(mes,ano)
            colunaDoRelatorioAnaliticoPorMes = RelatorioAnalitico()     
            Detalhado=(self.__scriptDao.BuscardespesaAnaliticodetalhadoCompetencia(periododeCompetenciaInicial, periododeCompetenciafinal, desp))#linha1  
            for aluguel in Detalhado:
                colunas = []
                colunas.append(datetime.strftime(aluguel['data'],'%d-%m-%Y'))
                colunas.append(aluguel['credor'])
                colunas.append(aluguel['discrimina'])
                colunas.append(locale.currency(aluguel['total'], symbol=False, grouping=True))
                linhas.append(colunas)  
        return linhas      
                
        
   
   
   
#ABA Relatorio   
    def GerarRelatorioDetalhadoAnalitico(self, ano, mesFinal, mesInicial):
        
        linhas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        for mes in range(mesInicial, mesesInteiro,1):
            self.__quantMeses+=1
            dtInicial = "{}-{}-01".format(ano, mes)
            dtFinal = "{}-{}-31".format(ano, mes)                   
            analiticoAdm=(self.__scriptDao.analiticoAdm(dtInicial, dtFinal,'1'))#linha1        

        
        return linhas
                        
        pass        
        
    def GerarRelatorioAnalitico(self, ano, mesFinal, mesInicial, diainicial, diafinal):
        
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
            colunaDoRelatorioAnaliticoPorMes = RelatorioAnalitico() 
                 
                      
            Aluguel=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'3'))['total']#linha1
            Iptu=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'4'))['total']#linha2
            Cedae=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'5'))['total']#linha3
            light=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'6'))['total']#linha4            
            Ceg=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'7'))['total']#linha5  
            Telefonia=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'8'))['total']#linha6  
            PnP=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'9'))['total']#linha7 
            Contabilidade=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'11'))['total']#linha8                                                
            Sigraf=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'12'))['total']#linha9         
            Informatica=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'14 '))['total']#linha10    
            MaterialdeEscricorio=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'15 '))['total']#linha11   
            MaterialLimpeza=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'16'))['total']#linha12                                                                                                                
            CartuchosParaImpressoras=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'17'))['total']#linha13 
            OnibusContinuo=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'18'))['total']#linha14   
            Correios=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'19'))['total']#linha15   
            DespesasBancarias=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'20'))['total']#linha16   
            OutrosAdminitrativos=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'21'))['total']#linha17 
            Kombi=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'23'))['total']#linha18  
            Motoboy=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'24'))['total']#linha19  
            OutrosFretes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'25'))['total']#linha20  
            Transportadoras=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'26'))['total']#linha21  
            Sedex=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'27'))['total']#linha22  
            Pedagio=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'30'))['total']#linha23  
            Simples=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'31'))['total']#linha24  
            OutrosImposos=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'34'))['total']#linha25
            Salarios=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'35'))['total']#linha26            
            Recisoes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'36'))['total']#linha27 
            ValeTransporte=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'37'))['total']#linha28
            ValeAlimentacao=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'38'))['total']#linha29  
            Inss=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'39'))['total']#linha30        
            Fgts=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'40'))['total']#linha31
            AjudaDeCusto=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'41'))['total']#linha32  
            Ferias=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'42'))['total']#linha33  
            HoraExtra=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'43'))['total']#linha34  
            Testes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'44'))['total']#linha35
            CestaBasica=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'45'))['total']#linha36  
            Lanches=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'46'))['total']#linha37  
            PesquisaDesenvolvimento=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'68'))['total']#linha37              
            Farmacia=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'47'))['total']#linha38  
            Treinamento=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'48'))['total']#linha39  
            Biscate=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'49'))['total']#linha40     
            Outros=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'50'))['total']#linha41  
            EmprestimosInternos=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'71'))['total']#linha42  
            Resultados=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'72'))['total']#linha43  
            ProLaboreDiretoria=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'73'))['total']#linha44
            ProLaboreGerencia=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'74'))['total']#linha45  
            DespesasCadastraisClientes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'75'))['total']#linha46
            Exame=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'77'))['total']#linha47
            Cipa=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'78'))['total']#linha48
            ServSeguranca=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'79'))['total']##linha49
            Minformatica=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'86'))['total']#linha50
            PlanodeSaude=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'91'))['total']#linha51
            FeriasTrabalhadas=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'92'))['total']#linha52
            EpiUniformes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'93'))['total']#linha53
            FestasEConfraternizacoes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'95'))['total']#linha54
            Mpredial=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'97'))['total']#linha55                                                                                                                                                                                                                                                                                                                                                           
            Meletrica=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'98'))['total']#linha56
            Marcondicionado=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'99'))['total']#linha57
            Mbebedouro=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'100'))['total']#linha58
            Mmoveis=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'101'))['total']#linha59 
            Moutrosmanut=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'102'))['total']#linha60
            MHidrualica=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'103'))['total']#linha61
            Maquinasequip=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'104'))['total']#linha62 
            ObraseMelhorias=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'105'))['total']#linha63 
            Bmoveis=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'106'))['total']#linha64 
            BInformaaticaeHard=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'107'))['total']#linha65 
            BFerramentas=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'109'))['total']#linha66
            AjudaECombustivel=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'110'))['total']#linha67 
            DespRepresentacao=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'111'))['total']#linha68 
            Propaganda=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'112'))['total']#linha69 
            Brindes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'113'))['total']#linha70 
            Binformaticasoft=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'120'))['total']#linha71
            Imaquinasequip=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'134'))['total']#linha72
            IcomputadoresHard=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'135'))['total']#linha73
            IoutrosInvestimentos=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'140'))['total']#linha74
            PremiosEGratificacoes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'141'))['total']#linha75
            DevolucaoVenda=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'144'))['total']#linha76
            Irrf=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'148'))['total']#linha77 
            Iof=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'149'))['total']#linha78
            JurosSemprestimos=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'151'))['total']#linha79     
            LicenciamentoseCertidoes=(self.__scriptDao.BuscardespesaAnalitico(dtInicial, dtFinal,'163'))['total']#linha79                 
            
                     
                                                                                                           
                                                                                                                                                                                                              
            colunaDoRelatorioAnaliticoPorMes.build().nomeMes(self.getNomeMesPorNumeroPeriodo(mes)).Aluguel(Aluguel).iptu(Iptu).Cedae(Cedae).Light(light).Ceg(Ceg).Telefonia(Telefonia)            
            colunaDoRelatorioAnaliticoPorMes.build().Pnp(PnP).Contabilidade(Contabilidade).Sigraf(Sigraf).Informatica(Informatica).MaterialdeEscritorio(MaterialdeEscricorio).MaterialLimpeza(MaterialLimpeza)
            colunaDoRelatorioAnaliticoPorMes.build().CartuchosparaImpressoras(CartuchosParaImpressoras).OnibusContinuo(OnibusContinuo).Correios(Correios).DespesasBancarias(DespesasBancarias).OutrosAdministrativos(OutrosAdminitrativos).Kombi(Kombi)
            colunaDoRelatorioAnaliticoPorMes.build().Motoboy(Motoboy).OutrosFretes(OutrosFretes).Transportadoras(Transportadoras).SedexRemessaVenda(Sedex).PedagioFretevenda(Pedagio).Simples(Simples)            
            colunaDoRelatorioAnaliticoPorMes.build().OutrosImpostos(OutrosImposos).Salarios(Salarios).Recisoes(Recisoes).ValeTranspote(ValeTransporte).ValeAlimentacao(ValeAlimentacao).Inss(Inss)  
            colunaDoRelatorioAnaliticoPorMes.build().Fgts(Fgts).AjudaDeCusto(AjudaDeCusto).Ferias(Ferias).HoraExtra(HoraExtra).Testes(Testes).CestaBasica(CestaBasica)  
            colunaDoRelatorioAnaliticoPorMes.build().Lanches(Lanches).PesquisaeDesenvolvimento(PesquisaDesenvolvimento).Farmacia(Farmacia).Treinamento(Treinamento).Biscate(Biscate).Outros(Outros).EmprestimosInternos(EmprestimosInternos)  
            colunaDoRelatorioAnaliticoPorMes.build().Resultados(Resultados).ProLaboreDiretoria(ProLaboreDiretoria).ProLaboreGerencia(ProLaboreGerencia).DespesasCadastraisClientes(DespesasCadastraisClientes).Exames(Exame).Cipa(Cipa)            
            colunaDoRelatorioAnaliticoPorMes.build().ServSeguranca(ServSeguranca).MInformatica(Minformatica).PlanoDeSaude(PlanodeSaude).FeriasTrabalhadas(FeriasTrabalhadas).EpiUniformes(EpiUniformes).FestaseConfraternizacoes(FestasEConfraternizacoes).MPredial(Mpredial)            
            colunaDoRelatorioAnaliticoPorMes.build().MEletrica(Meletrica).MArCondicionado(Marcondicionado).MBebedouro(Mbebedouro).MMoveis(Mmoveis).MOutrosManutencao(Moutrosmanut).MHidraulica(MHidrualica).BMaquinasEquip(Maquinasequip)            
            colunaDoRelatorioAnaliticoPorMes.build().ObrasEMelhorias(ObraseMelhorias).BMoveis(Bmoveis).BInformaticaEHardware(BInformaaticaeHard).BFerramentas(BFerramentas).AjudaCombustivel(AjudaECombustivel)
            colunaDoRelatorioAnaliticoPorMes.build().DespRepresentacao(DespRepresentacao).Propaganda(Propaganda).Brindes(Brindes).BInformaticaSoft(Binformaticasoft).IMaquinasEquips(Imaquinasequip).IComputadoresHardware(IcomputadoresHard)
            colunaDoRelatorioAnaliticoPorMes.build().IOutrosInvestimentos(IoutrosInvestimentos).PremioseGratificacoes(PremiosEGratificacoes).DevolucaoVenda(DevolucaoVenda).Irrf(Irrf).Iof(Iof).JurosSEmprestimos(JurosSemprestimos).LicenciamentoseCertidoes(LicenciamentoseCertidoes)
           
           
            colunaDoRelatorioAnaliticoPorMes.gerarTotal()            
            colunas.append(colunaDoRelatorioAnaliticoPorMes)
        self.__colunasAnaliticoGeral = colunas
        #self.GerarRelatorioAnaliticodetalhado(ano, mesFinal, mesInicial, diainicial, diafinal)
        return colunas
#Linhas  do programa    
    def  getLinhaDesp(self):
        linha=[]
        linha.append("Desp:")  
        for coluna in self.__colunasAnaliticoGeral:
            linha.append(coluna.getDesp())
        return linha         
        
        
    def  getLinhaNomeDesp(self):
        linha=[]
        linha.append("Nome Da Despesa:")  
        for coluna in self.__colunasAnaliticoGeral:
            linha.append(coluna.getNomeDesp())
        return linha         
        
    def getLinhaAluguel(self):  
        linha=[]
        linha.append("Aluguel:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getAluguel()
            linha.append(locale.currency(int(coluna.getAluguel()),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha        
    
    
    def getLinhaIptu(self):
        linha=[]
        linha.append("IPTU:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getIptu()
            linha.append(locale.currency(coluna.getIptu(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha           
    
    
    def getLinhaCedae(self):
        linha=[]
        linha.append("CEDAE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getCedae()
            linha.append(locale.currency(coluna.getCedae(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        
       # if total == 0:
        #    return []
        
        return linha       
    
    
    
    def getLinhaLight(self):
        linha=[]
        linha.append("LIGHT:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getLight()
            linha.append(locale.currency(coluna.getLight(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    
    def getLinhaCeg(self):
        linha=[]
        linha.append("CEG:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getCeg()
            linha.append(locale.currency(coluna.getCeg(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    def getLinhaTelefonia(self):
        linha=[]
        linha.append("TELEFONIA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getTelefonia()
            linha.append(locale.currency(coluna.getTelefonia(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha               

    def getLinhaPnP(self):
        linha=[]
        linha.append("PNP:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getPnp()
            linha.append(locale.currency(coluna.getPnp(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha       
    
    def getLinhaContabilidade(self):
        linha=[]
        linha.append("CONTABILIDADE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getContabilidade()
            linha.append(locale.currency(coluna.getContabilidade(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha        

    def getLinhaSigraf(self):
        linha=[]
        linha.append("SIGRAF:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getSigraf()
            linha.append(locale.currency(coluna.getSigraf(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    
    def getLinhaInformatica(self):
        linha=[]
        linha.append("INFORMÁTICA/CONTRATOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getInformatica()
            linha.append(locale.currency(coluna.getInformatica(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    
    def getLinhaMaterialdeEscritorio(self):
        linha=[]
        linha.append("MAT.ESCRITORIO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMaterialdeEscritorio()
            linha.append(locale.currency(coluna.getMaterialdeEscritorio(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    
    def getLinhaMaterialLimpeza(self):
        linha=[]
        linha.append("MAT.LIMPEZA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMaterialLimpeza()
            linha.append(locale.currency(coluna.getMaterialLimpeza(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
    
    def getLinhaCartuchosParaImpressoras(self):
        linha=[]
        linha.append("CARTUCHOS P/ IMPRESSORAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getCartuchosparaImpressoras()
            linha.append(locale.currency(coluna.getCartuchosparaImpressoras(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha        
    
    def getLinhaOnibusContinuo(self):
        linha=[]
        linha.append("ÔNIBUS CONTINUO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getOnibusContinuo()
            linha.append(locale.currency(coluna.getOnibusContinuo(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha   
    
    def getLinhaCorreios(self):
        linha=[]
        linha.append("CORREIOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getCorreios()
            linha.append(locale.currency(coluna.getCorreios(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                 
    
    def getLinhaDespesasBancarias(self):
        linha=[]
        linha.append("DESPESAS BANCÁRIAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getDespesasBancarias()
            linha.append(locale.currency(coluna.getDespesasBancarias(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha             
    
    
    def getLinhaOutrosAdministrativos(self):
        linha=[]
        linha.append("OUTROS ADMINISTRATIVOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getOutrosAdminitrativos()
            linha.append(locale.currency(coluna.getOutrosAdminitrativos(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha             
    
    
    def getLinhaKombi(self):
        linha=[]
        linha.append("KOMBI:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getKombi()
            linha.append(locale.currency(coluna.getKombi(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                           
        
    def getLinhaMotoboy(self):
        linha=[]
        linha.append("MOTOBOY:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMotoboy()
            linha.append(locale.currency(coluna.getMotoboy(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
    
    def getLinhaOutrosFretes(self):
        linha=[]
        linha.append("OUTROS FRETES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getOutrosFretes()
            linha.append(locale.currency(coluna.getOutrosFretes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha               
          
    
    def getLinhaTranpostadoras(self):
        linha=[]
        linha.append("TRANSPORTADORAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getTransportadoras()
            linha.append(locale.currency(coluna.getTransportadoras(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha          
    
    def getLinhaSedexRemessaVenda(self):
        linha=[]
        linha.append("SEDEX REMESSA VENDA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getSedexRemessaVenda()
            linha.append(locale.currency(coluna.getSedexRemessaVenda(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha          
    
    def getLinhaPedagioFreteVenda(self):
        linha=[]
        linha.append("PEDAGIO FRETE VENDA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getPedagioFreteVenda()
            linha.append(locale.currency(coluna.getPedagioFreteVenda(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
    
    def getLinhaSimples(self):
        linha=[]
        linha.append("SIMPLES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getSimples()
            linha.append(locale.currency(coluna.getSimples(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                                      
        
        
    def getLinhaOutrosImpostos(self):
        linha=[]
        linha.append("OUTROS IMPOSTOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna. getOutrosImpostos()
            linha.append(locale.currency(coluna. getOutrosImpostos(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaSalarios(self):
        linha=[]
        linha.append("SALARIOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getSalarios()
            linha.append(locale.currency(coluna.getSalarios(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaRecisoes(self):
        linha=[]
        linha.append("RECISÕES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getRecisoes()
            linha.append(locale.currency(coluna.getRecisoes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                  
        
        
    def getLinhaValeTransporte(self):
        linha=[]
        linha.append("VALE TRANSPORTE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getValeTransporte()
            linha.append(locale.currency(coluna.getValeTransporte(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaValeAlimentacao(self):
        linha=[]
        linha.append("VALE ALIMENTAÇÃO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getValeAlimentacao()
            linha.append(locale.currency(coluna.getValeAlimentacao(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha             
    
    def getLinhaInss(self):
        linha=[]
        linha.append("INSS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getiNSS()
            linha.append(locale.currency(coluna.getiNSS(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    def getLinhaFgts(self):
        linha=[]
        linha.append("FGTS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getFGTS()
            linha.append(locale.currency(coluna.getFGTS(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha   
    def getLinhaAjudaDeCusto(self):
        linha=[]
        linha.append("AJUDA DE CUSTO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getAjudaDeCusto()
            linha.append(locale.currency(coluna.getAjudaDeCusto(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 

    def getLinhaFerias(self):
        linha=[]
        linha.append("FÉRIAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getFerias()
            linha.append(locale.currency(coluna.getFerias(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaHoraExtra(self):
        linha=[]
        linha.append("HORA EXTRA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getHoraExtra()
            linha.append(locale.currency(coluna.getHoraExtra(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaTestes(self):
        linha=[]
        linha.append("TESTES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getTestes()
            linha.append(locale.currency(coluna.getTestes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaCestaBasica(self):
        linha=[]
        linha.append("CESTA BASICA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getCestaBasica()
            linha.append(locale.currency(coluna.getCestaBasica(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaLanches(self):
        linha=[]
        linha.append("LANCHES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getLanches()
            linha.append(locale.currency(coluna.getLanches(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhapesquisaDesenvolvimento(self):
        linha=[]
        linha.append("PESQUISA/DESENVOLVIMENTO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getPesquisaeDesenvolvimento()
            linha.append(locale.currency(coluna.getPesquisaeDesenvolvimento(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))    
        return linha  
    
    def getLinhaFarmacia(self):
        linha=[]
        linha.append("FARMACIA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getFarmacia()
            linha.append(locale.currency(coluna.getFarmacia(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaTreinamento(self):
        linha=[]
        linha.append("TREINAMENTO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getTreinamento()
            linha.append(locale.currency(coluna.getTreinamento(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaBiscate(self):
        linha=[]
        linha.append("BISCATE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBiscate()
            linha.append(locale.currency(coluna.getBiscate(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaOutros(self):
        linha=[]
        linha.append("OUTROS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getOutros()
            linha.append(locale.currency(coluna.getOutros(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha             
                                    
    def getLinhaEmprestimosInternos(self):
        linha=[]
        linha.append("EMPRÉSTIMOS INTERNOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getEmprestimosInternos()
            linha.append(locale.currency(coluna.getEmprestimosInternos(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha         

    def getLinhaResultados(self):
        linha=[]
        linha.append("RESULTADOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getResultados()
            linha.append(locale.currency(coluna.getResultados(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
            
    def getLinhaProLaboreDiretoria(self):
        linha=[]
        linha.append("PROLABORE DIRETORIA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getProLaboreDiretoria()
            linha.append(locale.currency(coluna.getProLaboreDiretoria(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaProLaboreGerencia(self):
        linha=[]
        linha.append("PROLABORE GERENCIA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getProLaboreGerencia()
            linha.append(locale.currency(coluna.getProLaboreGerencia(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     

    def getLinhaDespesasCadastraisClientes(self):
        linha=[]
        linha.append("DESP. CADASTRAIS CLIENTES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getDespesasCadastraisClientes()
            linha.append(locale.currency(coluna.getDespesasCadastraisClientes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha         

    def getLinhaExames(self):
        linha=[]
        linha.append("EXAMES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getExames()
            linha.append(locale.currency(coluna.getExames(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaCipa(self):
        linha=[]
        linha.append("CIPA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getCipa()
            linha.append(locale.currency(coluna.getCipa(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha              
    
    def getLinhaServSeguranca(self):
        linha=[]
        linha.append("SERV SEGURANÇA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getServSeguranca()
            linha.append(locale.currency(coluna.getServSeguranca(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha               

    def getLinhaMinformatica(self):
        linha=[]
        linha.append("(M)INFORMÁTICA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMInformatica()
            linha.append(locale.currency(coluna.getMInformatica(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha        
    
    def getLinhaPlanoDeSaude(self):
        linha=[]
        linha.append("(PLANO DE SAÚDE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getPlanoDeSaude()
            linha.append(locale.currency(coluna.getPlanoDeSaude(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
    
    
    def getLinhaFeriasTrabalhadas(self):
        linha=[]
        linha.append("FÉRIAS TRABALHADAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getFeriasTrabalhadas()
            linha.append(locale.currency(coluna.getFeriasTrabalhadas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha        
    
    def getLinhaEpiUniformes(self):
        linha=[]
        linha.append("EPI/UNIFORMES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getEpiUniformes()
            linha.append(locale.currency(coluna.getEpiUniformes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha             
    
    
    def getLinhaFestasEConfraternizações(self):
        linha=[]
        linha.append("FESTAS E CONFRATERNIZAÇÕES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getFestaseConfraternizacoes()
            linha.append(locale.currency(coluna.getFestaseConfraternizacoes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                       

    def getLinhaMPredial(self):
        linha=[]
        linha.append("(M)PREDIAL/ALVENARIA E PINTURA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMPredial()
            linha.append(locale.currency(coluna.getMPredial(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     

    def getLinhaMEletrica(self):
        linha=[]
        linha.append("(M)ELÉTRICA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMEletrica()
            linha.append(locale.currency(coluna.getMEletrica(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaMArCondicionado(self):
        linha=[]
        linha.append("(M) AR CONDICIONADO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMArCondicionado()
            linha.append(locale.currency(coluna.getMArCondicionado(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaMBebedouro(self):
        linha=[]
        linha.append("(M)BEBEDOURO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMBebedouro()
            linha.append(locale.currency(coluna.getMBebedouro(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaMMoveis(self):
        linha=[]
        linha.append("(M)MÓVEIS E UTENSILIOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMMoveis()
            linha.append(locale.currency(coluna.getMMoveis(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaMOutrosManut(self):
        linha=[]
        linha.append("(M)OUTROS MANUT.ADM:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMOutrosManutencao()
            linha.append(locale.currency(coluna.getMOutrosManutencao(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaMHidraulica(self):
        linha=[]
        linha.append("(M)HIDRAULICA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getMHidraulica()
            linha.append(locale.currency(coluna.getMHidraulica(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaBMaquinaseEquip(self):
        linha=[]
        linha.append("(B)MAQUINAS E EQUIPAMENTOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBMaquinasEquip()
            linha.append(locale.currency(coluna.getBMaquinasEquip(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaObrasEMelhorias(self):
        linha=[]
        linha.append("OBRAS E MELHORIAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getObrasEMelhorias()
            linha.append(locale.currency(coluna.getObrasEMelhorias(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha  
    
    def getLinhaBMoveis(self):
        linha=[]
        linha.append("(B)MÓVEIS E UTENSILIOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBMoveis()
            linha.append(locale.currency(coluna.getBMoveis(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    def getLinhaBInformaticaEHardware(self):
        linha=[]
        linha.append("(B)INFORMATICA E HARDWARE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBInformaticaEHardware()
            linha.append(locale.currency(coluna.getBInformaticaEHardware(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha   
    
    def getLinhaBFerramentas(self):
        linha=[]
        linha.append("(B)FERRAMENTAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBFerramentas()
            linha.append(locale.currency(coluna.getBFerramentas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                       
                       
                       
    def getLinhaAjudaCombustivel(self):
        linha=[]
        linha.append("AJUDA E COMBUSTIVEL:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getAjudaCombustivel()
            linha.append(locale.currency(coluna.getAjudaCombustivel(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha   
    
    def getLinhaDespRepresentacao(self):
        linha=[]
        linha.append("DESP. REPRESENTAÇÃO:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getDespRepresentacao()
            linha.append(locale.currency(coluna.getDespRepresentacao(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha   
    
    def getLinhaPropaganda(self):
        linha=[]
        linha.append("PROPAGANDA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getPropaganda()
            linha.append(locale.currency(coluna.getPropaganda(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                                              
    
    def getLinhaBrindes(self):
        linha=[]
        linha.append("BRINDES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBrindes()
            linha.append(locale.currency(coluna.getBrindes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha   
    
    def getLinhaBInformaticaSoft(self):
        linha=[]
        linha.append("(B)INFORMÁTICA E SOFTWARE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getBInformaticaSoft()
            linha.append(locale.currency(coluna.getBInformaticaSoft(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
    
    def getLinhaIMaquinasEquips(self):
        linha=[]
        linha.append("(I)MAQUINAS E EQUIPAMENTOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getIMaquinasEquips()
            linha.append(locale.currency(coluna.getIMaquinasEquips(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha    
    
    def getLinhaIComputadoresHardware(self):
        linha=[]
        linha.append("(I)COMPUTADORES E HARDWARE:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getIComputadoresHardware()
            linha.append(locale.currency(coluna.getIComputadoresHardware(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha 
    
    def getLinhaIOutrosInvestimentos(self):
        linha=[]
        linha.append("(I)OUTROS INVESTIMENTOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getIOutrosInvestimentos()
            linha.append(locale.currency(coluna.getIOutrosInvestimentos(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha       
    
    def getLinhaPremioseGratificacoes(self):
        linha=[]
        linha.append("PREMIOS E GRATIFICAÇÕES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getPremioseGratificacoes()
            linha.append(locale.currency(coluna.getPremioseGratificacoes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha    
    
    def getLinhaDevolucaoVenda(self):
        linha=[]
        linha.append("DEVOLUÇÃO VENDA:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getDevolucaoVenda()
            linha.append(locale.currency(coluna.getDevolucaoVenda(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha     
    
    def getLinhaIrrf(self):
        linha=[]
        linha.append("IRRF-IMP. RENDA S/REN FIN:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getIrrf()
            linha.append(locale.currency(coluna.getIrrf(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha       
    
    
    def getLinhaIof(self):
        linha=[]
        linha.append("IOF S/OP.FINANCEIRAS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getIof()
            linha.append(locale.currency(coluna.getIof(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha        
    
    def getLinhaJurosSEmprestimos(self):
        linha=[]
        linha.append("JUROS / EMPRÉSTIMOS:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getJurosSEmprestimos()
            linha.append(locale.currency(coluna.getJurosSEmprestimos(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
    
    def getLinhaLicenciamentoseCertidoes(self):
        linha=[]
        linha.append("LICENCIAMENTOS E CERTIDÕES:")
        total =0
        for coluna in self.__colunasAnaliticoGeral:
            total+=coluna.getLicenciamentoseCertidoes()
            linha.append(locale.currency(coluna.getLicenciamentoseCertidoes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha                              
        
    def getNomeMesPorNumeroPeriodo(self, numMes):
        match numMes:
            case 1:
                return "JAN"
            case 2:
                return "FEV"
            case 3:
                return "MAR"
            case 4:
                return "ABR"
            case 5:
                return "MAI"
            case 6:
                return "JUN"
            case 7:
                return "JUL"
            case 8:
                return "AGO"
            case 9:
                return "SET"
            case 10:
                return "OUT"
            case 11:
                return "NOV"
            case 12:
                return "DEZ"        
                      
    def AnaliticoDetalhado(self, ano, mesFinal, mesInicial, desp):
        linhas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        #TEM QUE CONVERTER DE TEXTO PRA INTEIRO
        
        for mes in range(mesInicial, mesesInteiro,1):
            self.__quantMeses+=1
            #Template
            #if(mes < 10):
              #  periododeCompetenciaInicial = "0{}/{}".format(mes,ano)
             #   periododeCompetenciafinal = "0{}/{}".format(mes,ano)
            #else:
             #   periododeCompetenciaInicial = "{}/{}".format(mes,ano)
             #  periododeCompetenciafinal = "{}/{}".format(mes,ano)
            dtInicial = "{}-{}-01".format(ano, mes)
            dtFinal = "{}-{}-31".format(ano, mes)                   
            #if mesInicial==mesFinal:    
             #   dtInicial = "{}-{}-{}".format(ano, mes, diainicial)
              #  dtFinal = "{}-{}-{}".format(ano, mes, diafinal) 
            colunaDoRelatorioAnaliticoPorMes = RelatorioAnalitico()     
            Detalhado=(self.__scriptDao.BuscardespesaAnaliticodetalhado(dtInicial, dtFinal, desp))#linha1  
            for aluguel in Detalhado:
                colunas = []
                colunas.append(datetime.strftime(aluguel['data'],'%d-%m-%Y'))
                colunas.append(aluguel['credor'])
                colunas.append(aluguel['discrimina'])
                colunas.append(locale.currency(aluguel['total'], symbol=False, grouping=True))
                linhas.append(colunas)  
        return linhas      
           
    def ApagarValorAnalitico(self):
      linhas = []
      self.__scriptDao.DropTableValorAnalitico()#trncate  
      return linhas    
  
    def getNomeMesPorNumeroPeriodo(self, numMes):
        match numMes:
            case 1:
                return "JAN"
            case 2:
                return "FEV"
            case 3:
                return "MAR"
            case 4:
                return "ABR"
            case 5:
                return "MAI"
            case 6:
                return "JUN"
            case 7:
                return "JUL"
            case 8:
                return "AGO"
            case 9:
                return "SET"
            case 10:
                return "OUT"
            case 11:
                return "NOV"
            case 12:
                return "DEZ"          


pass  