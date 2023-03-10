from entity.dto.AnaliticoDTO import AnaliticoDTO
from entity.dto.DespesaAnaliticoDTO import DespesaAnaliticoDTO
from persistence.VendaRepository import VendaRepository
from persistence.SaidaRepository import SaidaRepository
from persistence.ScriptsRepository import ScriptsRepository
from datetime import *
from  datetime import datetime
from mysql.connector import Error
import pandas as xlsService
import pandas as xlswriter


import pandas as  pd
import datetime
from entity.dto.RelatorioVenda import RelatorioVenda
import locale


class VendaService :

    def __init__(self):
        self.__vendaDao = VendaRepository()
        self.__saidaDao = SaidaRepository()
        self.__custoDao = ScriptsRepository()
        self.__colunasRelatorioGeral=[]
        locale.setlocale(locale.LC_ALL,'pt_BR')
        locale.setlocale(locale.LC_MONETARY,'pt_BR.UTF-8')
        self.__quantMeses = 0
    
    def gerarTabelaPorData(self, dtInicial, dtFinal, despesa1, despesa2,  despesa3):

        linhas=[]


        linhasIndice=["VENDA C/N VENDAS","VENDA C/N SERVIÇOS","VENDAS S/N","TOTAL VENDAS","SIMPLES", "TOTAL VENDAS LIQ.","VENDAS DEVOLVIDAS","VENDAS CANCELADAS","CORTESIAS","DESC ERRO OPERACIONAL","DESC ERRO COMERCIAL","COMISSÕES INTERNAS",
        "COMISSÕES EXTERNAS","SALARIOS PRODUÇÃO","H.E.","TESTES","BISCATES","FÉRIAS TRABALHADAS ","PRO LABORE GERENCIA","LAMINAS","MADEIRAS","VAZADORES","BORRACHAS","PERTINAX","MATERIAL P / DESTAQUE"," MAQUINAS PARA REVENDA","USO E CONSUMO",
        "MANUTENÇÃO","OUTROS","SERVIÇOS EXTERNOS","DESPESAS COM ERROS"]


        
        linhas.append(self.__vendaDao.buscarValorFinalDaVendaComNotaFiscal(dtInicial, dtFinal)) #linha1
        linhas.append(self.__vendaDao.buscarValorFinalDoServicoComNotaFiscal(dtFinal, dtFinal)) #linha2
        linhas.append(self.__vendaDao.buscarValorFinalDoServicoComBoleta(dtInicial, dtFinal))#linha3
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '31'))#linha4
        linhas.append(self.__vendaDao.buscarValorFinalDaDevolucaoDaVenda(dtInicial, dtFinal))#linha5
        linhas.append(self.__vendaDao.buscaValorFinalDasVendasCanceladas(dtInicial, dtFinal))#linha6
        linhas.append(self.__vendaDao.buscarValorFinalDasDespesasPorPeriodo(dtInicial, dtFinal,'129'))#linha7
        linhas.append(self.__vendaDao.buscarValorFinalDasDespesasPorPeriodo(dtInicial, dtFinal,'130'))#linha8
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '28'))#linha9
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '28'))#linha10
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '29'))#linha11
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','45','2'))#linha12
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','43','1'))#linha13
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '44'))#linha14
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '49'))#linha15
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '92'))#linha16
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '74'))#linha17
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '51'))#linha18
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '52'))#linha19
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '54'))#linha20
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '55'))#linha21
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '59'))#linha22
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '62'))#linha23
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '142'))#linha24
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','2'))#linha25
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','1'))#linha26
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '65')) #linha27
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo1(dtInicial , dtFinal, '66','145'))#linha28
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '119')) #linha29
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','8'))#linha30
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','8'))#linha31
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','9'))#linha32
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','9'))#linha33
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','13'))#linha34
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','13'))#linha35
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','10'))#linha36
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','10'))#linha37
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','56','10'))#linha38
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','12'))#linha39
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','12'))#linha40  
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','63','11'))#linha41
        linhas.append(self.__custoDao.BuscaValordecusto('01/02/2021', '28/02/2021','64','11'))#linha42
        linhas.append(self.__custoDao.BuscarValorPorCustoData('01/02/2021' , '28/02/2021', '1'))#linha43
        linhas.append(self.__custoDao.BuscarValorPorCustoData('01/02/2021' , '28/02/2021', '46'))#linha44
        linhas.append(self.__custoDao.BuscarValorPorCustoData('01/02/2021' , '28/02/2021', '47'))#linha45
        linhas.append(self.__custoDao.BuscarValorPorCustoData('01/02/2021' , '28/02/2021', '48'))#linha46
        linhas.append(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '20')) #linha47
        linhas.append(self.__vendaDao.buscarValorFinalDasDespesasPortipo(dtInicial,dtFinal,'3', '5' ,'15')) #linha48
                              
        
        
        print(linhas)
        
        coluna1='Índices'
        coluna2= 'Janeiro'
        coluna3= 'Fevereiro'
        coluna4= 'Março'
        coluna5= 'Abril'
        coluna6= 'Maio'
        coluna7= 'Junho'
        coluna8= 'Julho'
        coluna9= 'Agosto'
        coluna10= 'Setembro'
        coluna11= 'Outubro'
        coluna12= 'Novembro'
        coluna13= 'Dezembro'
        
        #saida = ({coluna1:linhasIndice, coluna2:linhas, coluna3:linhas, coluna4:linhas, coluna5:linhas,
        #coluna6:linhas, coluna7:linhas, coluna8:linhas, coluna9:linhas, coluna10:linhas, coluna11:linhas, coluna12:linhas
        #,coluna13:linhas})
        return linhas 
        
    def gerarTabelaPorCompete(self, ano, mesInicial, mesFinal):
        
        colunas = []
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

            dtInicial = "{}-{}-01".format(ano, mes)
            dtFinal = "{}-{}-31".format(ano, mes)
            colunaDoRelatorioPorMes = RelatorioVenda()
            
            Vendas=()
            vendasCNVendas = self.__vendaDao.buscarValorFinalDaVendaComNotaFiscal(dtInicial, dtFinal)['VALOR_FINAL']#linha1
            vendasCNServicos = self.__vendaDao.buscarValorFinalDoServicoComNotaFiscal(dtInicial, dtFinal)['VALOR_FINAL'] #linha 2
            vendas_S_N= self.__vendaDao.buscarValorFinalDoServicoComBoleta(dtInicial, dtFinal)['VALOR_FINAL'] #linha3
            Totalvendas=()
            Simples=(self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial,periododeCompetenciafinal,'31'))['VALOR_FINAL']#linha 4 
            TotalvendasLiq=()
            Vendas_Devolvidas=(self.__vendaDao.buscaValorFinalDasDespesasPorCompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal,'144'))['VALOR_FINAL']#linha 5
            Vendas_Canceladas=(self.__vendaDao.buscaValorFinalDasVendasCanceladas(dtInicial, dtFinal))['VALOR_FINAL']#linha 6 
            Cortesias=(self.__vendaDao.buscaValorFinalDasDespesasPorCompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '129' ))['VALOR_FINAL']#LINHA7
            DescErroOperacional=(self.__vendaDao.buscaValorFinalDasDespesasPorCompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal,'130'))['VALOR_FINAL']#LINHA8
            DescErroComercial=self.__vendaDao.buscaValorFinalDasDespesasPorCompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal,'131')['VALOR_FINAL']#linha9
            ComissoesInternas=(self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial , periododeCompetenciafinal, '28'))['VALOR_FINAL']#linha10
            ComissoesExternas=(self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial , periododeCompetenciafinal, '29'))['VALOR_FINAL']#linha11
            TotalGeral=()
            Despesas=()
            ProLaboreGerencia=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '74','2'))['resultado']#linha17
            Laminas=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '51','47'))['resultado']#linha18
            Madeiras=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '52','47'))['resultado']#linha19
            Vazadores=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '54','47'))['resultado']#linha20
            Borrachas=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '55','47'))['resultado']#linha21
            Pertinax=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '59','47'))['resultado']#linha22
            MaterialPDestaque=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '62','47'))['resultado']#linha23
            MaquinasParaRevenda=self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '142')['VALOR_FINAL']#linha24
            Outros=(self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '65'))['VALOR_FINAL']#linha 27
            ServicosExternos=(self.__saidaDao.buscarValorFinalDasaidaPorCompetencia1(periododeCompetenciaInicial ,periododeCompetenciafinal, '66','145'))['VALOR_FINAL']#linha28
            DespesacomErros=(self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '119'))['VALOR_FINAL']#linha 29
           
            #Clientes
            LaminasClientes=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '51','50'))['resultado']#linha18
            MadeirasClientes=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '52','50'))['resultado']#linha19
            VazadoresClientes=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '54','50'))['resultado']#linha20
            BorrachasClientes=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '55','50'))['resultado']#linha21
            PertinaxClientes=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '59','50'))['resultado']#linha22
            MaterialPDestaqueClientes=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '62','50'))['resultado']#linha23    
            Cliches=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '58','50'))['resultado']#linha28                    
            
            Desenho=()
            
            SalariosDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'45','8'))['resultado']#linha12
            HeDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'43','8'))['resultado']#linha13
            testesDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '44','8'))['resultado']#linha14
            biscatesDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '49','8'))['resultado']#linha15
            FeriasTrabalhadasDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '92','8'))['resultado']#linha16            
            usoeConsumoDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial,periododeCompetenciafinal,'63','8'))['resultado']#linha30
            manutencaoDesenho=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'64','8'))['resultado']#linha30
           
            Plotter=()
            UsoeconsumoPlotter=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'63','9'))['resultado']#linha32
            MAnutencaoPlotter=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'64','9'))['resultado']#linha33
            
            Laser=()
            SalariosLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'45','10'))['resultado']#linha12
            HeLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'43','10'))['resultado']#linha13
            testesLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '44','10'))['resultado']#linha14
            biscatesLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '49','10'))['resultado']#linha15
            FeriasTrabalhadasLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '92','10'))['resultado']#linha16                
            UsoeConsumoLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'63','10'))['resultado']#linha36
            ManutencaoLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '64','10'))['resultado']#linha37
            GaspLaser=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'56','10'))['resultado']#linha38
            
            Montagem=()
            SalariosMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'45','48'))['resultado']#linha12
            HeMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'43','48'))['resultado']#linha13
            testesMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '44','48'))['resultado']#linha14
            biscatesMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '49','48'))['resultado']#linha15
            FeriasTrabalhadasMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '92','48'))['resultado']#linha16                
            UsoeConsumoMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'63','48'))['resultado']#linha39
            ManutencaoMontagem=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'64','48'))['resultado']#linha40
           
            Borracha=()
            SalariosBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'45','49'))['resultado']#linha12
            HeBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'43','49'))['resultado']#linha13
            testesBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '44','49'))['resultado']#linha14
            biscatesBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '49','49'))['resultado']#linha15
            FeriasTrabalhadasBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '92','49'))['resultado']#linha16                
            UsoeConsumoBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'63','49'))['resultado']#linha41
            ManutencaoBorracha=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal,'64','49'))['resultado']#linha42
            TotalVendasGeral=()
            TotalResultado1=()
            #self.__vendaDao.BuscarDadosDepreciacao() 
            Depreciação=(self.__vendaDao.BuscarDadosDepreciacao())['VALOR_FINAL']
            TotalResultado2=()
            DespesasAdministrativas=(self.__custoDao.BuscarValorPorCustoCompeteSemCusto(periododeCompetenciaInicial ,periododeCompetenciafinal, '1'))['resultado']#linha43
            DespesasCentral=(self.__custoDao.BuscarValorPorCustoCompeteSemCusto(periododeCompetenciaInicial ,periododeCompetenciafinal, '51'))['resultado']#linha43
            SalariosAdm=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '35','1'))['resultado']#linha44
            SalariosComercial=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '35','45'))['resultado']#linha45
            SalariosTransporte=(self.__custoDao.BuscarValorPorCustoCompete(periododeCompetenciaInicial ,periododeCompetenciafinal, '35','46'))['resultado']#linha46
            TotalAdm=()
            DespesasFinanceiras=(self.__saidaDao.buscarValorFinalDasDespesasPorCompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '20'))['VALOR_FINAL'] #linha47
            Juros=(self.__vendaDao.buscarValorFinalDasDespesasPortipoecompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '5'))['VALOR_FINAL'] #linha48
            JurosSobreAplicacoes=(self.__vendaDao.buscarValorFinalDasDespesasPortipoecompetencia(periododeCompetenciaInicial ,periododeCompetenciafinal, '15'))['VALOR_FINAL'] #linha49
            JurosSobreVendas=(self.__vendaDao.buscarValorFinalDosJuros(dtInicial,dtFinal))['VALOR_FINAL'] #linha50   
            TotalResultadoEzipa=()
            ResultadoWoodPaper=(self.__vendaDao.BuscarDadosdaWoodpaper(dtInicial ,dtFinal))['VALOR_FINAL']
            ResultadoGeral=()
              

            colunaDoRelatorioPorMes.build().nomeMes(self.getNomeMesPorNumeroPeriodo(mes)).vendasCNVendas(vendasCNVendas).vendasCNServicos(vendasCNServicos).vendas_S_N(vendas_S_N).Simples(Simples).Vendas_Devolvidas(Vendas_Devolvidas).Vendas_Canceladas(Vendas_Canceladas)
            colunaDoRelatorioPorMes.build().Cortesias(Cortesias).DescErroOperacional(DescErroOperacional).DescErroComercial(DescErroComercial).ComissoesInternas(ComissoesInternas).ComissoesExternas(ComissoesExternas)
            colunaDoRelatorioPorMes.build().ProLaboreGerencia(ProLaboreGerencia).Laminas(Laminas).Madeiras(Madeiras).Vazadores(Vazadores).Borrachas(Borrachas).Pertinax(Pertinax).MaterialPDestaque(MaterialPDestaque).MaquinasParaRevenda(MaquinasParaRevenda).Outros(Outros).ServicosExternos(ServicosExternos).DespesacomErros(DespesacomErros)
            colunaDoRelatorioPorMes.build().LaminasClientes(LaminasClientes).MadeirasClientes(MadeirasClientes).VazadoresClientes(VazadoresClientes).BorrachasClientes(BorrachasClientes).PertinaxClientes(PertinaxClientes).MaterialPDestaqueClientes(MaterialPDestaqueClientes).Cliches(Cliches)
            colunaDoRelatorioPorMes.build().Desenho(Desenho).SalariosDesenho(SalariosDesenho).H_e(HeDesenho).testes(testesDesenho).biscates(biscatesDesenho).FeriasTrabalhadas(FeriasTrabalhadasDesenho).usoeConsumoDesenho(usoeConsumoDesenho).manutencao(manutencaoDesenho)
            colunaDoRelatorioPorMes.build().Plotter(Plotter).UsoeConsumoPlotter(UsoeconsumoPlotter).ManutencaoPlotter(MAnutencaoPlotter)
            colunaDoRelatorioPorMes.build().Laser(Laser).SalariosLaser(SalariosLaser).H_eLaser(HeLaser).testesLaser(testesLaser).biscatesLaser(biscatesLaser).FeriasTrabalhadasLaser(FeriasTrabalhadasLaser).UsoeConsumoLaser(UsoeConsumoLaser).ManUtencaoLaser(ManutencaoLaser).GaspLaser(GaspLaser)
            colunaDoRelatorioPorMes.build().Montagem(Montagem).SalariosMontagem(SalariosMontagem).H_eMontagem(HeMontagem).testesMontagem(testesMontagem).biscatesMontagem(biscatesMontagem).FeriasTrabalhadasMontagem(FeriasTrabalhadasMontagem).usoEconsumoMontagem(UsoeConsumoMontagem).ManutencaoMontagem(ManutencaoMontagem)
            colunaDoRelatorioPorMes.build().Borracha(Borracha).SalariosBorracha(SalariosBorracha).H_eBorracha(HeBorracha).testesBorracha(testesBorracha).biscatesBorracha(biscatesBorracha).FeriasTrabalhadasBorracha(FeriasTrabalhadasBorracha).UsoeConsumoBorracha(UsoeConsumoBorracha).ManutencaoBorracha(ManutencaoBorracha)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
            colunaDoRelatorioPorMes.build().DespesasAdminsitrativas(DespesasAdministrativas).DespesasCentrais(DespesasCentral).SalariosAdm(SalariosAdm).SalariosComercial(SalariosComercial).SalariosTransporte(SalariosTransporte).DespesasFinanceiras(DespesasFinanceiras).juros(Juros).JurosSobreAplicacoces(JurosSobreAplicacoes).JurosSobreVendas(JurosSobreVendas)
            colunaDoRelatorioPorMes.build().TotalVendas()
            colunaDoRelatorioPorMes.build().TotalLiq()
            colunaDoRelatorioPorMes.build().totalVendasGeral()       
            colunaDoRelatorioPorMes.build().totalGeralServiços() 
            colunaDoRelatorioPorMes.build().totaldeResultados1()    
            colunaDoRelatorioPorMes.build().Depreciacao(Depreciação)
            colunaDoRelatorioPorMes.build().totaldeResultados2() 
            colunaDoRelatorioPorMes.build().totalGeralAdministrativo() 
            colunaDoRelatorioPorMes.build().totalResultadoEzipa()    
            colunaDoRelatorioPorMes.build().gerartotalWoodpaper(ResultadoWoodPaper)
            colunaDoRelatorioPorMes.build().gerartotalEzipaeWoodpaper()                         
            colunaDoRelatorioPorMes.gerarTotal()
            colunas.append(colunaDoRelatorioPorMes)
        self.__colunasRelatorioGeral = colunas
        return colunas

    def  getLinhaVendas(self):
        linha=[]
        linha.append("VENDAS")  
        for coluna in self.__colunasRelatorioGeral:
            linha.append(coluna.getVendas())
        return linha 
        

    pass
          
    def getNomeMesPorNumeroCompetencia(self, numMes):
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
            

        saida = xlsService.DataFrame({})
    
        saida.to_excel('c:/Projeto/Resultado Financeiro.xlsx', sheet_name='Relatório Financeiro Ezipa', index=True)
   
   

    
    
    def talvezGerarRelatorioGeralAnual(self, ano, mesFinal, mesInicial, diainicial, diafinal):
        
        colunas = []
        mesesInteiro=int(mesFinal)+1 
        self.__quantMeses = 0
        #CONVERTE DE TEXTO PRA INTEIRO
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
            if mesInicial==mesFinal:
                 
                dtInicial = "{}-{}-{}".format(ano, mes, diainicial)
                dtFinal = "{}-{}-{}".format(ano, mes, diafinal) 
            dataInicial = "{}-{}-01".format(ano, mes)
            dataFinal = "{}-{}-31".format(ano, mes)
            colunaDoRelatorioPorMes = RelatorioVenda()             

            vendasCNVendas = self.__vendaDao.buscarValorFinalDaVendaComNotaFiscal(dtInicial, dtFinal)['VALOR_FINAL'] #linha1
            vendasCNServicos = self.__vendaDao.buscarValorFinalDoServicoComNotaFiscal(dtFinal, dtFinal)['VALOR_FINAL'] #linha 2
            vendas_S_N= self.__vendaDao.buscarValorFinalDoServicoComBoleta(dtInicial, dtFinal)['VALOR_FINAL'] #linha3
            TotalVendas=()
            Simples=(self.__saidaDao.buscarValorFinalDaSaidaPorcompetencia(periododeCompetenciaInicial, periododeCompetenciafinal,'31'))['VALOR_FINAL']#linha4
            TotalVendasLiq=()
            Vendas_Devolvidas=(self.__vendaDao.buscarValorFinalDasDespesasPorPeriodo(dtInicial, dtFinal,'144'))['VALOR_FINAL']#linha5
            Vendas_Canceladas=(self.__vendaDao.buscaValorFinalDasVendasCanceladas(dtInicial, dtFinal))['VALOR_FINAL']#linha6
            Cortesias=(self.__vendaDao.buscarValorFinalDasDespesasPorPeriodo(dtInicial, dtFinal,'129'))['VALOR_FINAL']#linha7
            DescErroOperacional=(self.__vendaDao.buscarValorFinalDasDespesasPorPeriodo(dtInicial, dtFinal,'130'))['VALOR_FINAL']#linha8
            DescErroComercial=(self.__vendaDao.buscarValorFinalDasDespesasPorPeriodo(dtInicial, dtFinal,'131'))['VALOR_FINAL']#linha9
            ComissoesInternas=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '28'))['VALOR_FINAL']#linha10
            ComissoesExternas=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '29'))['VALOR_FINAL']#linha11
            TotalVendasGeral=()
            ProLaboreGerencia=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '74','2'))['resultado']#linha17 
            #Estoque
            Laminas=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '51','47'))['resultado']#linha18
            Madeiras=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '52','47'))['resultado']#linha19
            Vazadores=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '54','47'))['resultado']#linha20
            Borrachas=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '55','47'))['resultado']#linha21
            Pertinax=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '59','47'))['resultado']#linha22
            MaterialPDestaque=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '62','47'))['resultado']#linha23
            MaquinasParaRevenda=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '142'))['VALOR_FINAL']#linha24
            Outros=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '65'))['VALOR_FINAL']#linha27
            ServicosExternos=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo1(dtInicial , dtFinal, '66','145')['VALOR_FINAL'])#linha28           
            #Clientes 
            LaminasClientes=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '51','50'))['resultado']#linha18
            MadeirasClientes=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '52','50'))['resultado']#linha19
            VazadoresCLientes=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '54','50'))['resultado']#linha20
            BorrachasCLientes=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '55','50'))['resultado']#linha21
            PertinaxClientes=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '59','50'))['resultado']#linha22
            MaterialPDestaqueClientes=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '62','50'))['resultado']#linha23
            Cliches=(self.__custoDao.BuscaValordecusto(dtInicial , dtFinal, '58','50'))['resultado']#linha28.5                        
            DespesacomErros=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '119'))['VALOR_FINAL']#linha29            
            Desenho=()
            SalariosDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'45','8'))['resultado']#linha12
            HeDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'43','8'))['resultado']#linha13   
            testesDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '44','8'))['resultado']#linha14 
            biscatesDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '49','8'))['resultado']#linha15 
            FeriasTrabalhadasDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '92','8'))['resultado']#linha16            
            usoeConsumoDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'63','8'))['resultado']#linha30
            manutencaoDesenho=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'64','8'))['resultado']#linha31             
            Plotter=()        
            UsoeconsumoPlotter=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'63','9'))['resultado']#linha32
            MAnutencaoPlotter=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'64','9'))['resultado']#linha33  
            Laser=()
            SalariosLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'45','10'))['resultado']#linha12
            HeLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'43','10'))['resultado']#linha13   
            testesLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '44','10'))['resultado']#linha14 
            biscatesLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '49','10'))['resultado']#linha15 
            FeriasTrabalhadasLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '92','10'))['resultado']#linha16                      
            UsoeConsumoLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'63','10'))['resultado']#linha36
            ManutencaoLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'64','10'))['resultado']#linha37  
            GaspLaser=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'56','10'))['resultado']#linha38 
            Montagem=()
            SalariosMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'45','48'))['resultado']#linha48
            HeMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'43','48'))['resultado']#linha13   
            testesMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '44','48'))['resultado']#linha14 
            biscatesMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '49','48'))['resultado']#linha15 
            FeriasTrabalhadasMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '92','48'))['resultado']#linha16                             
            UsoeConsumoMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'63','48'))['resultado']#linha39
            ManutencaoMontagem=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'64','48'))['resultado']#linha40
            Borracha=()
            SalariosBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'45','49'))['resultado']#linha12
            HeBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'43','49'))['resultado']#linha13   
            testesBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '44','49'))['resultado']#linha14 
            biscatesBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '49','49'))['resultado']#linha15 
            FeriasTrabalhadasBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '92','49'))['resultado']#linha16                
            UsoeConsumoBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'63','49'))['resultado']#linha41
            ManutencaoBorracha=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal,'64','49'))['resultado']#linha42
            
            TotalVendasGeral=()
            TotalResultado1=()
            
            Depreciação=(self.__vendaDao.BuscarDadosDepreciacao())['VALOR_FINAL']
            TotalResultados2=()
            DespesasAdministrativas=(self.__custoDao.BuscarValorPorCustoData(dtInicial, dtFinal, '1'))['resultado']#linha43
            DespesasCentral=(self.__custoDao.BuscarValorPorCustoData(dtInicial, dtFinal, '51'))['resultado']#linha43
            SalariosAdm=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '35','1'))['resultado']#linha44
            SalariosComercial=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '35','45'))['resultado']#linha45
            SalariosTransporte=(self.__custoDao.BuscaValordecusto(dtInicial, dtFinal, '35','46'))['resultado']#linha46
            TotalAdm=()
            
            DespesasFinanceiras=(self.__saidaDao.buscarValorFinalDasaidaPorPeriodo(dtInicial , dtFinal, '20'))['VALOR_FINAL'] #linha47
            Juros=(self.__vendaDao.buscarValorFinalDasDespesasPortipo(dtInicial,dtFinal,'5'))['VALOR_FINAL'] #linha48    
            JurosSobreAplicacoes=(self.__vendaDao.buscarValorFinalDasDespesasPortipo(dtInicial,dtFinal,'15'))['VALOR_FINAL'] #linha49
            JurosSobreVendas=(self.__vendaDao.buscarValorFinalDosJuros(dtInicial,dtFinal))['VALOR_FINAL'] #linha50        
            TotalResultadoEzipa=()  
         
            ResultadoWoodPaper=(self.__vendaDao.BuscarDadosdaWoodpaper(dataInicial, dataFinal))['VALOR_FINAL']
            ResultadoGeral=()
              
    
            colunaDoRelatorioPorMes.build().nomeMes(self.getNomeMesPorNumeroPeriodo(mes)).vendasCNVendas(vendasCNVendas).vendasCNServicos(vendasCNServicos).vendas_S_N(vendas_S_N).Simples(Simples).Vendas_Devolvidas(Vendas_Devolvidas).Vendas_Canceladas(Vendas_Canceladas)
            colunaDoRelatorioPorMes.build().Cortesias(Cortesias).DescErroOperacional(DescErroOperacional).DescErroComercial(DescErroComercial).ComissoesInternas(ComissoesInternas).ComissoesExternas(ComissoesExternas)
            colunaDoRelatorioPorMes.build().ProLaboreGerencia(ProLaboreGerencia).Laminas(Laminas).Madeiras(Madeiras).Vazadores(Vazadores).Borrachas(Borrachas).Pertinax(Pertinax).MaterialPDestaque(MaterialPDestaque).MaquinasParaRevenda(MaquinasParaRevenda).Outros(Outros).ServicosExternos(ServicosExternos).DespesacomErros(DespesacomErros)
            colunaDoRelatorioPorMes.build().LaminasClientes(LaminasClientes).MadeirasClientes(MadeirasClientes).VazadoresClientes(VazadoresCLientes).BorrachasClientes(BorrachasCLientes).PertinaxClientes(PertinaxClientes).MaterialPDestaqueClientes(MaterialPDestaqueClientes).Cliches(Cliches)
            colunaDoRelatorioPorMes.build().Desenho(Desenho).SalariosDesenho(SalariosDesenho).H_e(HeDesenho).testes(testesDesenho).biscates(biscatesDesenho).FeriasTrabalhadas(FeriasTrabalhadasDesenho).usoeConsumoDesenho(usoeConsumoDesenho).manutencao(manutencaoDesenho)
            colunaDoRelatorioPorMes.build().Plotter(Plotter).UsoeConsumoPlotter(UsoeconsumoPlotter).ManutencaoPlotter(MAnutencaoPlotter)
            colunaDoRelatorioPorMes.build().Laser(Laser).SalariosLaser(SalariosLaser).H_eLaser(HeLaser).testesLaser(testesLaser).biscatesLaser(biscatesLaser).FeriasTrabalhadasLaser(FeriasTrabalhadasLaser).UsoeConsumoLaser(UsoeConsumoLaser).ManUtencaoLaser(ManutencaoLaser).GaspLaser(GaspLaser)
            colunaDoRelatorioPorMes.build().Montagem(Montagem).SalariosMontagem(SalariosMontagem).H_eMontagem(HeMontagem).testesMontagem(testesMontagem).biscatesMontagem(biscatesMontagem).FeriasTrabalhadasMontagem(FeriasTrabalhadasMontagem).usoEconsumoMontagem(UsoeConsumoMontagem).ManutencaoMontagem(ManutencaoMontagem)
            colunaDoRelatorioPorMes.build().Borracha(Borracha).SalariosBorracha(SalariosBorracha).H_eBorracha(HeBorracha).testesBorracha(testesBorracha).biscatesBorracha(biscatesBorracha).FeriasTrabalhadasBorracha(FeriasTrabalhadasBorracha).UsoeConsumoBorracha(UsoeConsumoBorracha).ManutencaoBorracha(ManutencaoBorracha)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
            colunaDoRelatorioPorMes.build().DespesasAdminsitrativas(DespesasAdministrativas).DespesasCentrais(DespesasCentral).SalariosAdm(SalariosAdm).SalariosComercial(SalariosComercial).SalariosTransporte(SalariosTransporte).DespesasFinanceiras(DespesasFinanceiras).juros(Juros).JurosSobreAplicacoces(JurosSobreAplicacoes).JurosSobreVendas(JurosSobreVendas)
            colunaDoRelatorioPorMes.build().TotalVendas()
            colunaDoRelatorioPorMes.build().TotalLiq()
            colunaDoRelatorioPorMes.build().totalVendasGeral()       
            colunaDoRelatorioPorMes.build().totalGeralServiços() 
            colunaDoRelatorioPorMes.build().totaldeResultados1()    
            colunaDoRelatorioPorMes.build().Depreciacao(Depreciação)
            colunaDoRelatorioPorMes.build().totaldeResultados2() 
            colunaDoRelatorioPorMes.build().totalGeralAdministrativo() 
            colunaDoRelatorioPorMes.build().totalResultadoEzipa()    
            colunaDoRelatorioPorMes.build().gerartotalWoodpaper(ResultadoWoodPaper)
            colunaDoRelatorioPorMes.build().gerartotalEzipaeWoodpaper()                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           
            #.LinhasIndice(self.getLinhasIndice())                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
            #DEPOIS DE PEGAR TODOS OS VALORES, CRIAR UM METODO PARA SOMAR TODOS OS VALORES E COLOCAR NUM ATRIBUTO
            #DA CLASSE
            colunaDoRelatorioPorMes.gerarTotal()
            colunas.append(colunaDoRelatorioPorMes)
        self.__colunasRelatorioGeral = colunas
        return colunas


        


    
    def getLinhaVendasCNVendas(self):
        linha=[]
        linha.append("VENDA C/N VENDAS:")
        total =0
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getvendasCNVendas()
            linha.append(locale.currency(coluna.getvendasCNVendas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha
    
    def getLinhaVendasCNServicos(self):
        linha=[]
        linha.append("VENDA C/N SERVIÇOS:")
        total =0
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getVendasCNServicos()
            linha.append(locale.currency(coluna.getVendasCNServicos(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha
    
    def getLinhaVendas_S_N(self):
        linha=[]    
        linha.append("VENDAS S/N:") 
        total =0       
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getVendas_S_N()
            linha.append(locale.currency(coluna.getVendas_S_N(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))   
        return linha
    
    def getLinhaTotal(self):
        linha=[]
        linha.append("TOTAL VENDAS")          
        for coluna in self.__colunasRelatorioGeral:
            linha.append(coluna.Totalvendas(),symbol=False, grouping=True)
        return linha
        
    def getLinhaSimples(self):
        linha=[]
        linha.append("SIMPLES:")   
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSimples()    
            linha.append(locale.currency(coluna.getSimples(),symbol=False, grouping=True))           
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha
   
    def  getLinhaTotalLiq(self):
        linha=[]
        linha.append("TOTAL VENDAS LIQ:")  
        total =0                 
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTotalLiq()           
            linha.append(locale.currency(coluna.getTotalLiq(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha

    def getLinhaTotalDeVendas(self):
        linha=[]
        linha.append("TOTAL VENDAS:")
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTotalVendas()           
            linha.append(locale.currency(coluna.getTotalVendas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha
    
    def  getLinhaVendas_Devolvidas(self):
        linha=[]
        linha.append("VENDAS DEVOLVIDAS:") 
        total =0            
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getVendas_Devolvidas() 
            linha.append(locale.currency(coluna.getVendas_Devolvidas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha
    
    def  getLinhaVendas_Canceladas(self):
        linha=[]
        linha.append("VENDAS CANCELADAS:")   
        total =0                
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getVendas_Canceladas()           
            linha.append(locale.currency(coluna.getVendas_Canceladas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))           
        return linha
    
    def  getLinhaCortesias(self):
        linha=[]
        linha.append("CORTESIAS:")    
        total =0         
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getCortesias()              
            linha.append(locale.currency(coluna.getCortesias(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha
    
    def  getLinhaDescErroOperacional(self):
        linha=[]
        linha.append("DESC ERRO OPERACIONAL:")     
        total =0       
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDescErroOperacional()              
            linha.append(locale.currency(coluna.getDescErroOperacional(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha
    
    def  getLinhaDescErroComercial(self):
        linha=[]
        linha.append("DESC ERRO COMERCIAL:")    
        total =0                
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDescErroComercial()              
            linha.append(locale.currency(coluna.getDescErroComercial(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True)) 
        return linha    
   
    def  getLinhaComissoesInternas(self):
        linha=[]
        linha.append("COMISSÕES INTERNAS:") 
        total =0                  
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getComissoesInternas()                
            linha.append(locale.currency(coluna.getComissoesInternas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))          
        return linha       
    
    def  getLinhaComissoesExternas(self):
        linha=[]
        linha.append("COMISSÕES EXTERNAS:")  
        total =0                   
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getComissoesExternas()    
            linha.append(locale.currency(coluna.getComissoesExternas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha 
    
    def getLinhaTotalVendasGeral(self):
        linha=[]
        linha.append("TOTAL GERAL:")      
        total =0                             
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTotalVendasGeral() 
            linha.append(locale.currency(coluna.getTotalVendasGeral(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))        
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))           
        return linha 
   
    def  getLinhaProLaboreGerencia(self):
        linha=[]
        linha.append("PRO LABORE GERENCIA:")   
        total =0                
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getProLaboreGerencia()   
            linha.append(locale.currency(coluna.getProLaboreGerencia(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                 
        return linha  
    
    def  getLinhaLaminas(self):
        linha=[]
        linha.append("LAMINAS:")  
        total =0                  
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getLaminas()               
            linha.append(locale.currency(coluna.getLaminas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha    
    
    def  getLinhaMadeiras(self):
        linha=[]
        linha.append("MADEIRAS:") 
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getMadeiras()             
            linha.append(locale.currency(coluna.getMadeiras(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                 
        return linha     
   
    def  getLinhaVazadores(self):
        linha=[]
        linha.append("VAZADORES:") 
        total =0               
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getVazadores()               
            linha.append(locale.currency(coluna.getVazadores(),symbol=False, grouping=True))    
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                    
        return linha    
   
    def  getLinhaBorrachas(self):
        linha=[]
        linha.append("BORRACHAS:") 
        total =0            
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getBorrachas()                  
            linha.append(locale.currency(coluna.getBorrachas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha 
    
    def  getLinhaPertinax(self):
        linha=[]
        linha.append("PERTINAX:")   
        total =0                 
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getPertinax()            
            linha.append(locale.currency(coluna.getPertinax(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))      
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))              
        return linha   
  
    def  getLinhaMaterialPDestaque(self):
        linha=[]
        linha.append("MATERIAL P / DESTAQUE:")  
        total =0                        
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getMaterialPDestaque()             
            linha.append(locale.currency(coluna.getMaterialPDestaque(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha     
  
    def  getLinhaMaquinasParaRevenda(self):
        linha=[]
        linha.append("MAQUINAS PARA REVENDA:")  
        total =0             
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getMaquinasParaRevenda()            
            linha.append(locale.currency(coluna.getMaquinasParaRevenda(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))        
        return linha                  
   
   
    def  getLinhaLaminasClientes(self):
        linha=[]
        linha.append("LAMINAS:")  
        total =0                  
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getLaminasClientes()               
            linha.append(locale.currency(coluna.getLaminasClientes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha    
    
    def  getLinhaMadeirasClientes(self):
        linha=[]
        linha.append("MADEIRAS:") 
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getMadeirasClientes()             
            linha.append(locale.currency(coluna.getMadeirasClientes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                 
        return linha     
   
    def  getLinhaVazadoresClientes(self):
        linha=[]
        linha.append("VAZADORES:") 
        total =0               
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getVazadoresClientes()               
            linha.append(locale.currency(coluna.getVazadoresClientes(),symbol=False, grouping=True))    
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                    
        return linha    
   
    def  getLinhaBorrachasClientes(self):
        linha=[]
        linha.append("BORRACHAS:") 
        total =0            
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getBorrachasClientes()                  
            linha.append(locale.currency(coluna.getBorrachasClientes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha 
    
    def  getLinhaPertinaxClientes(self):
        linha=[]
        linha.append("PERTINAX:")   
        total =0                 
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getPertinaxClientes()            
            linha.append(locale.currency(coluna.getPertinaxClientes(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))      
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))              
        return linha   
  
    def  getLinhaMaterialPDestaqueClientes(self):
        linha=[]
        linha.append("MATERIAL P / DESTAQUE:")  
        total =0                        
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getMaterialPDestaqueClientes()             
            linha.append(locale.currency(coluna.getMaterialPDestaqueClientes(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha     

    
    def  getLinhaOutros(self):
        linha=[]
        linha.append("OUTROS:")  
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getOutros()               
            linha.append(locale.currency(coluna.getOutros(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha     
    
    def  getLinhaServicosExternos(self):
        linha=[]
        linha.append("SERVIÇOS EXTERNOS:")  
        total =0   
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getServicosExternos()
            linha.append(locale.currency(coluna.getServicosExternos(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha    
  
    def  getLinhaCliches(self):
        linha=[]
        linha.append("CLICHES:")  
        total =0   
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getCliches()
            linha.append(locale.currency(coluna.getCliches(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha    
    
    def  getLinhaDespesacomErros(self):
        linha=[]
        linha.append("DESPESAS COM ERROS:")  
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDespesacomErros()           
            linha.append(locale.currency(coluna.getDespesacomErros(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))      
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha    
              
    def  getLinhaDesenho(self):
        linha=[]
        linha.append("DESENHO:")  
        for coluna in self.__colunasRelatorioGeral:
           linha.append(coluna.getDesenho()) 
        return linha     

    def  getLinhaSalariosDesenho(self):
        linha=[]
        linha.append("SALARIOS:")   
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosDesenho()                
            linha.append(locale.currency(coluna.getSalariosDesenho(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))        
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))           
        return linha 
   
    def  getLinhaHE(self):
        linha=[]
        linha.append("H.E.:")
        total =0              
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getH_E()               
            linha.append(locale.currency(coluna.getH_E(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha     
   
    def  getLinhatestes(self):
        linha=[]
        linha.append("TESTES:")  
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTestes()              
            linha.append(locale.currency(coluna.getTestes(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha       
   
    def  getLinhabiscates(self):
        linha=[]
        linha.append("BISCATES:") 
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getBiscates()              
            linha.append(locale.currency(coluna.getBiscates(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                  
        return linha   
   
    def  getLinhaFeriasTrabalhadas(self):
        linha=[]
        linha.append("FÉRIAS TRABALHADAS:") 
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getFeriasTrabalhadas()             
            linha.append(locale.currency(coluna.getFeriasTrabalhadas(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                   
        return linha  
    
    def  getLinhausoeConsumoDesenho(self):
        linha=[]
        linha.append("USO E CONSUMO:")  
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getusoeConsumoDesenho()                    
            linha.append(locale.currency(coluna.getusoeConsumoDesenho(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                 
        return linha    
  
    def  getLinhamanutencao(self):
        linha=[]
        linha.append("MANUTENÇÃO:")  
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getmanutencao()             
            linha.append(locale.currency(coluna.getmanutencao(),symbol=False, grouping=True) )
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))               
        return linha    

   
    def  getLinhaPlotter(self):
        linha=[]
        linha.append("PLOTTER:")  
        for coluna in self.__colunasRelatorioGeral:                 
           linha.append(coluna.getPlotter())            
        return linha    
    
    def  getLinhaUsoeConsumoPlotter(self):
        linha=[]
        linha.append("USO E CONSUMO:")  
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getUsoeConsumoPlotter()            
            linha.append(locale.currency(coluna.getUsoeConsumoPlotter(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                  
        return linha  
   
    def  getLinhaManutencaoPlotter(self):
        linha=[]
        linha.append("MANUTENÇÃO:")  
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getManutencaoPlotter()                
            linha.append(locale.currency(coluna.getManutencaoPlotter(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha  
   
   
    
    def  getLinhaLaser(self): 
        linha=[]           
        linha.append("LASER:")  
        for coluna in self.__colunasRelatorioGeral:
            linha.append(coluna.getLaser())
        return linha 
    
    def  getLinhaSalariosLaser(self):
        linha=[]
        linha.append("SALARIOS:")   
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosLaser()                
            linha.append(locale.currency(coluna.getSalariosLaser(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))        
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))           
        return linha 
   
    def  getLinhaHELaser(self):
        linha=[]
        linha.append("H.E.:")
        total =0              
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getH_ELaser()               
            linha.append(locale.currency(coluna.getH_ELaser(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha     
   
    def  getLinhatestesLaser(self):
        linha=[]
        linha.append("TESTES:")  
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTestesLaser()              
            linha.append(locale.currency(coluna.getTestesLaser(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha       
   
    def  getLinhabiscatesLaser(self):
        linha=[]
        linha.append("BISCATES:") 
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getBiscatesLaser()              
            linha.append(locale.currency(coluna.getBiscatesLaser(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                  
        return linha   
   
    def  getLinhaFeriasTrabalhadasLaser(self):
        linha=[]
        linha.append("FÉRIAS TRABALHADAS:") 
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getFeriasTrabalhadasLaser()             
            linha.append(locale.currency(coluna.getFeriasTrabalhadasLaser(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                   
        return linha  
    
    def  getLinhaUsoeConsumoLaser(self):
        linha=[]
        linha.append("USO E CONSUMO:") 
        total =0                
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getUsoeConsumoLaser()           
            linha.append(locale.currency(coluna.getUsoeConsumoLaser(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))      
        return linha    
      
    def  getLinhaManutencaoLaser(self):
        linha=[]
        linha.append("MANUTENÇÃO:")  
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getManUtencaoLaser()             
            linha.append(locale.currency(coluna.getManUtencaoLaser(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha      
   
    def  getLinhaGaspLaser(self):
        linha=[]
        linha.append("GAS P/ LASER:")  
        total =0               
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getGaspLaser()           
            linha.append(locale.currency(coluna.getGaspLaser(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))   
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))  
        return linha      
  
 
   
    def  getLinhaMontagem(self): 
        linha=[]           
        linha.append("MONTAGEM:")  
        for coluna in self.__colunasRelatorioGeral:
            linha.append(coluna.getMontagem())
        return linha     
   
    def  getLinhaSalariosMontagem(self):
        linha=[]
        linha.append("SALARIOS:")   
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosMontagem()                
            linha.append(locale.currency(coluna.getSalariosMontagem(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))        
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))           
        return linha 
   
    def  getLinhaHEMontagem(self):
        linha=[]
        linha.append("H.E.:")
        total =0              
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getH_eMontagem()               
            linha.append(locale.currency(coluna.getH_eMontagem(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha     
   
    def  getLinhatestesMontagem(self):
        linha=[]
        linha.append("TESTES:")  
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTestesMontagem()              
            linha.append(locale.currency(coluna.getTestesMontagem(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha       
   
    def  getLinhabiscatesMontagem(self):
        linha=[]
        linha.append("BISCATES:") 
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getbiscatesMontagem()              
            linha.append(locale.currency(coluna.getbiscatesMontagem(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                  
        return linha   
   
    def  getLinhaFeriasTrabalhadasMontagem(self):
        linha=[]
        linha.append("FÉRIAS TRABALHADAS:") 
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getFeriasTrabalhadasMontagem()             
            linha.append(locale.currency(coluna.getFeriasTrabalhadasMontagem(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                   
        return linha     
   
    def  getLinhausoEconsumoMontagem(self):
        linha=[]
        linha.append("USO E CONSUMO:")  
        total =0                 
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getUsoeConsumoMontagem()            
            linha.append(locale.currency(coluna.getUsoeConsumoMontagem(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))      
        return linha      
    
    def  getLinhaManuTencaoMontagem(self):
        linha=[]
        linha.append("MANUTENÇÃO:")  
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getManuTencaoMontagem()                
            linha.append(locale.currency(coluna.getManuTencaoMontagem(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha    
    
  
      
    def  getLinhaBorracha(self): 
        linha=[]           
        linha.append("BORRACHA:")  
        for coluna in self.__colunasRelatorioGeral:
            linha.append(coluna.getBorracha())
        return linha    
     
    def  getLinhaSalariosBorracha(self):
        linha=[]
        linha.append("SALARIOS:")   
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosBorracha()                
            linha.append(locale.currency(coluna.getSalariosBorracha(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))        
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))           
        return linha 
   
    def  getLinhaHEBorracha(self):
        linha=[]
        linha.append("H.E.:")
        total =0              
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getH_eBorracha()               
            linha.append(locale.currency(coluna.getH_eBorracha(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))            
        return linha     
   
    def  getLinhatestesBorracha(self):
        linha=[]
        linha.append("TESTES:")  
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTestesBorracha()              
            linha.append(locale.currency(coluna.getTestesBorracha(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                
        return linha       
   
    def  getLinhabiscatesBorracha(self):
        linha=[]
        linha.append("BISCATES:") 
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getbiscatesBorracha()              
            linha.append(locale.currency(coluna.getbiscatesBorracha(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                  
        return linha   
   
    def  getLinhaFeriasTrabalhadasBorracha(self):
        linha=[]
        linha.append("FÉRIAS TRABALHADAS:") 
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getFeriasTrabalhadasBorracha()             
            linha.append(locale.currency(coluna.getFeriasTrabalhadasBorracha(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))        
        return linha   
    def  getLinhaUsoeConsumoBorracha(self):
        linha=[]
        linha.append("USO E CONSUMO:") 
        total =0                 
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getUsoeConsumoBorracha()            
            linha.append(locale.currency(coluna.getUsoeConsumoBorracha(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha      
   
    def  getLinhaManutencaoBorracha(self):
        linha=[]
        linha.append("MANUTENÇÃO:")  
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getManutencaoBorracha()                   
            linha.append(locale.currency(coluna.getManutencaoBorracha(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha    
   
    
    
    def getLinhatotalGeralServiços(self):
        linha=[]
        linha.append("TOTAL GERAL:")  
        total =0                  
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTotalGeralServiços()             
            linha.append(locale.currency(coluna.getTotalGeralServiços(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                    
        return linha
    
    def getLinhatotaldeResultados1(self):
        linha=[]
        linha.append("TOTAL DE RESULTADOS 1:")  
        total =0               
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTotaldeResultados1()                   
            linha.append(locale.currency(coluna.getTotaldeResultados1(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                    
        return linha
   
    def getLinhaDepreciacao(self):
        linha=[]
        linha.append("DEPRECIAÇÃO:") 
        total =0                      
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDepreciacao()             
            linha.append(locale.currency(coluna.getDepreciacao(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                    
        return linha
    
    def getLinhatotaldeResultados2(self):
        linha=[]
        linha.append("TOTAL DE RESULTADOS 2:")   
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getTotaldeResultados2()              
            linha.append(locale.currency(coluna.getTotaldeResultados2(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))               
        return linha
        
    def  getLinhaDespesasAdministrativas(self):
        linha=[]
        linha.append("DESPESAS ADMINISTRATIVAS:")
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDespesasAdminsitrativas()      
            linha.append(locale.currency(coluna.getDespesasAdminsitrativas(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha       
    
    def  getLinhaDespesasCentrais(self):
        linha=[]
        linha.append("DESPESAS CENTRAL:")
        total =0                     
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDespesasCentrais()      
            linha.append(locale.currency(coluna.getDespesasCentrais(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha    
   
    def  getLinhaSalariosAdm(self):
        linha=[]
        linha.append("SALARIOS ADM:")  
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosAdm()       
            linha.append(locale.currency(coluna.getSalariosAdm(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))      
        return linha    
    
    def  getLinhaSalariosComercial(self):
        linha=[]
        linha.append("SALARIOS ATENDIMENTO:")  
        total =0         
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosComercial()           
            linha.append(locale.currency(coluna.getSalariosComercial(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True)) 
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))      
        return linha    
  
    def  getLinhaSalariosTransporte(self):
        linha=[]
        linha.append("SALARIOS EXPEDIÇÃO:")
        total =0                 
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getSalariosTransporte()          
            linha.append(locale.currency(coluna.getSalariosTransporte(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha    
    
    def getLinhatotalGeralAdministrativo(self):
        linha=[]
        linha.append("TOTAL ADM:")          
        total =0          
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.gettotalGeralAdministrativo()
            linha.append(locale.currency(coluna.gettotalGeralAdministrativo(),symbol=False, grouping=True))  
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))                 
        return linha  
   
    def  getLinhaDespesasFinanceiras(self):
        linha=[]
        linha.append("DESPESAS FINANCEIRAS:")
        total =0   
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getDespesasFinanceiras()               
            linha.append(locale.currency(coluna.getDespesasFinanceiras(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha  
  
    def  getLinhaJuros(self):
        linha=[]
        linha.append("JUROS:")
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getJuros()              
            linha.append(locale.currency(coluna.getJuros(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha  
    
    def  getLinhaJurosSobreAplicacoes(self):
        linha=[]
        linha.append("JUROS SOBRE APLICAÇÕES:")
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getJurosSobreAplicacoes()              
            linha.append(locale.currency(coluna.getJurosSobreAplicacoes(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha   
    
    def  getLinhaJurosSobreVendas(self):
        linha=[]
        linha.append("JUROS SOBRE VENDAS:")
        total =0           
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.getJurosSobreVendas()              
            linha.append(locale.currency(coluna.getJurosSobreVendas(),symbol=False, grouping=True)) 
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha      


           
    def getLinhatotalResultadoEzipa(self):
        linha=[]
        linha.append("TOTAL RESULTADO EZIPA:")    
        total =0                  
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.gettotalResultadoEzipa()              
            linha.append(locale.currency(coluna.gettotalResultadoEzipa(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))  
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))        
        return linha 
    
    def getlinhatotalWoodpaper(self):
        linha=[]
        linha.append("WOODPAPER")        
        total =0    
        for coluna in self.__colunasRelatorioGeral:
            total+=coluna.gettotalWoodpaper()                                            
            linha.append(locale.currency(coluna.gettotalWoodpaper(),symbol=False, grouping=True))
        media = total /self.__quantMeses
        linha.append(locale.currency(round(float(media),2),symbol=False, grouping=True))    
        linha.append(locale.currency(round(float(total),2),symbol=False, grouping=True))
        return linha     
    def getlinhatotalEzipaeWoodpaper(self):
        linha=[]
        total =0            
        linha.append("RESULTADO EZIPA E WOODPAPER")          
        for coluna in self.__colunasRelatorioGeral:  
            total+=coluna.gettotalEzipaeWoodpaper()
            linha.append(locale.currency(coluna.gettotalEzipaeWoodpaper(),symbol=False, grouping=True))
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
            
    def gerarCadastroWoodpaper(self, listaDeWoodpaper): 
        for woodpaper in listaDeWoodpaper:
            self.__vendaDao.CadastrodeDadosdaWoodpaper(woodpaper.getData(), woodpaper.getValor())
        
    def gerarDepreciacao(self, valor ): 
        self.__vendaDao.valorDeDepreciacao(valor)        
        
        pass
    
    def ApagarValorCusto(self):
      linhas = []
      self.__custoDao.TruncateValorCusto()#trncate  
      return linhas