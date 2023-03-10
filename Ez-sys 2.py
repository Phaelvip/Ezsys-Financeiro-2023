from tkinter.ttk import *
from  tkinter  import *
from pandas import *
from tkcalendar import*
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from tkinter import messagebox
from tkinter import ttk
from tkinter import Entry
import tkinter  as tk
from tkcalendar import DateEntry
from mysql.connector import Error
from entity.Woodpaper import Woodpaper
from entity.dto.RelatorioAnalitico import RelatorioAnalitico
from persistence.DbConnecion import DbConnection
from service.LoginService import LoginService
from service.VendaService import VendaService
from service.AnaliticoService import AnaliticoService
from service.AnaliticoCentralService import AnaliticoCentralService
from service.PermFinanceiraService import  PermFinanceiroService
from datetime import date, datetime
import csv
import win32api
from entity.Usuariopc import Usuario
import threading
import time
import pdfkit
from UI import limpaWidget
import os
import win32print
from  Leitura import Leitura
import configparser
from tkinter.filedialog import asksaveasfile, asksaveasfilename
#from tkPDFViewer import tkPDFViewer as pdf 
#import mysql.connector
#from reportlab.pdfgen import canvas
#from reportlab.lib.pagesizes import 
#import matplotlib.pyplot as plt
#import locale
#import win32com.client
#import tempfile
#import subprocess



#con = mysql.connector.connect(host='127.0.0.1', database='ez-sys', user='root', password='q9w8e7#MTB')
#con = mysql.connector.connect(host='127.0.0.1', database='ezsys', user='root', password='q9w8e7#MTB')
usuarioLogado=None
jan = Tk()
jan.title("Ez sys")
jan.geometry("1100x650+100+10")
jan.configure(background="White")
jan.resizable(width=False, height=False)
jan.attributes("-alpha", 1.1)



logo = PhotoImage(file='Resource\logo\Ezipa.gif')

logo2 = PhotoImage(file='Resource\logo\Ezipanovo.gif')

logo3 = PhotoImage(file='Resource\logo\Logo teste2.png')


jan.iconbitmap('Resource\Financeiro\Icone.bmp')


topframe = Frame(jan, width=9999, height=135, bg="#040D9E", relief="raise")
topframe.pack(side=TOP,fill = 'both', expand = True)
bottomframe = Frame(jan, width=9999, height=665,bg="white", relief="raise")
bottomframe.pack(side=BOTTOM,fill = 'both', expand = True)
LogoLabel = Label(topframe, image=logo, bg="#040D9E")
LogoLabel.place(x=940, y=0)
LogoLabel = Label(topframe, image=logo3, bg="#040D9E")
LogoLabel.place(x=0, y=0)
UserLabel = Label(bottomframe, text="Usuário", font=("Calibri",15), bg="white", fg="black" )
UserLabel.place(x=310,y=200)
UserEntry = ttk.Entry(bottomframe, width=25)
UserEntry.bind("<KeyPress>", lambda e: PassEntry.focus() if e.char == '\r' else None) 
UserEntry.place(x=390,y=205)
PassLabel = Label(bottomframe, text="Senha", font=("Calibri",15), bg="white", fg="black")
PassLabel.place(x= 590,y=200)
PassEntry = ttk.Entry(bottomframe, width=25, show="*")
PassEntry.bind("<KeyPress>", lambda e: Entrar() if e.char == '\r' else None)
PassEntry.place(x=660,y=205)
config = configparser.ConfigParser()
config.read('Leitura.ini')
db_name = config['DATABASE'].get('dbname')
# Check if a database connection exists
if db_name:
    # Create a label to display the database name
    db_name_label = Label(bottomframe, text=f"Connected to database: {db_name}", bg="white", fg="black")
    db_name_label.place(x=10,y=475)
else:
    # Display a message if no database connection exists
    message_label = Label(bottomframe, text="No database connection", bg="white", fg="black")
    message_label.place(x=10,y=475)


def Entrar():
   nome = UserEntry.get()  
   senha = PassEntry.get()  
   loginService = LoginService()
   usuarioLogado = loginService.realizarLoginComUsuarioESenha(nome, senha)
   if usuarioLogado !=False :
         
        messagebox.showinfo('Login', 'Logado com Sucesso')
       
   else:
        messagebox.showerror('Erro','verifique as credenciais de Login e senha')
        return
   permissoes=loginService.verificarPermissionamento(usuarioLogado.getusuarioid())
   usuarioLogado.setPermissoes(permissoes)   
   vendaService = VendaService()
   analiticoService=AnaliticoService()
   analiticoCentralService=AnaliticoCentralService()
          
      
          
   new=Toplevel()
  
   new.title("Sistema Ezipa")
   new.geometry("1100x650+100+10")
   new.grid()
   new.configure(background="white")
   new.resizable(width=True, height=True)
   new.attributes("-alpha", 1.1)
   jan.withdraw()
   

   

   Upframe = Frame(master=new, width=99999, height=200, bg="#040D9E", relief="raise")
   Upframe.pack(fill=tk.X, side=tk.TOP,  expand=True)
   #Upframe.pack(side=TOP) 
   LogoLabel = Label(Upframe, image=logo2, bg="#040D9E") 
   LogoLabel.place(x=1, y=1)   
   Downframe = Frame(master=new, width=99999, height=700,bg='white', relief="raise")
   Downframe.pack(fill=tk.X, side=tk.BOTTOM,  expand=True)


   #Downframe.pack(side=BOTTOM)
   EmpresaLabel = Label(Upframe, text="Empresa: Ezipa", font=("Calibri",8), bg="#040D9E", fg="White")
   EmpresaLabel.place(x= 550 ,y=30)   
   UsuarioLabel = Label(Upframe, text="Usuário: "+usuarioLogado.getNome(), font=("Calibri",8), bg="#040D9E", fg="White")
   UsuarioLabel.place(x= 550 ,y=50)
   Datalabel = Label(Upframe, text= "Data:" f"{datetime.now(): %d:%m:%y}", bg="#040D9E", fg="White")
   Datalabel.pack()
   Datalabel.place(x= 550 ,y=70)  

   

   def aperteEnter(event):
      event.widget.config( relief = RIDGE )
      btn.bind( "<Enter>", command=aperteEnter)
      
   def gestDeCadastro():
      limpaWidget.limpar(Downframe)   
      def aperteEnter(event):
         event.widget.config( relief = SOLID )
         btn.bind( "<Enter>", command=aperteEnter)
         
   
    

      grafico1 = PhotoImage(file="Resource\Cadastros\Contas Monetarias.png")
      grafico1= grafico1.subsample(2,2)
      figura = Label(image=grafico1)
      figura.image=grafico1
      if(usuarioLogado.getPermissoes().perm_cadfin_contMonetaria== True): 
       btn=Button(Downframe,image=grafico1, text="Contas Monetárias",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=150,y=115)
      
      grafico2 = PhotoImage(file="Resource\Cadastros\Plano de Receitas.png")
      grafico2= grafico2.subsample(2,2)
      figura = Label(image=grafico2)
      figura.image=grafico2
      if(usuarioLogado.getPermissoes().perm_cadfin_planReceitas== True): 
       btn=Button(Downframe,image=grafico2, text="Planos De Receitas",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=265,y=115)
      
      grafico3 = PhotoImage(file="Resource\Cadastros\Plano de Pagamentos.png")
      grafico3= grafico3.subsample(2,2)
      figura = Label(image=grafico3)
      figura.image=grafico3
      if(usuarioLogado.getPermissoes().perm_cadfin_PlanPagt== True): 
       btn=Button(Downframe,image=grafico3, text="Planos De Pagamentos",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=380,y=115)
      
      grafico4 = PhotoImage(file="Resource\Cadastros\Grupos de Despesas.png")
      grafico4= grafico4.subsample(2,2)
      figura = Label(image=grafico4)
      figura.image=grafico4
      if(usuarioLogado.getPermissoes().perm_cadfin_grupdespesas== True): 
       btn=Button(Downframe,image=grafico4, text="Grupos De Despesas",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=510,y=115)
      
      grafico5 = PhotoImage(file="Resource\Cadastros\Centro de Custos.png")
      grafico5= grafico5.subsample(2,2)
      figura = Label(image=grafico5)
      figura.image=grafico5
      if(usuarioLogado.getPermissoes().perm_cadfin_centdecustos== True): 
       btn=Button(Downframe,image=grafico5, text="Centro de Custos",bd=0, highlightthickness=0, bg="white", compound = TOP, relief="solid")
      btn.place(x=630,y=115)
      
      grafico6 = PhotoImage(file="Resource\Cadastros\Operações Financeiras.png")
      grafico6= grafico6.subsample(2,2)
      figura = Label(image=grafico6)
      figura.image=grafico6
      if(usuarioLogado.getPermissoes().perm_cadfin_opfinanceiras== True): 
       btn=Button(Downframe,image=grafico6, text="Operações Financeiras",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=730,y=115)
      
      grafico7 = PhotoImage(file="Resource\Cadastros\Códigos Fiscais.png")
      grafico7= grafico7.subsample(2,2)
      figura = Label(image=grafico7)
      figura.image=grafico7
      if(usuarioLogado.getPermissoes().perm_cadfin_codfiscais== True): 
       btn=Button(Downframe,image=grafico7, text="Códigos Fiscais",bd=0, highlightthickness=0, bg="white", compound = TOP) 
      btn.place(x=200,y=215)     
      
      grafico8 = PhotoImage(file="Resource\Cadastros\Formas de pagamento.png")
      grafico8= grafico8.subsample(2,2)
      figura = Label(image=grafico8)
      figura.image=grafico8
      if(usuarioLogado.getPermissoes().perm_cadfin_condpagamentos== True):       
       btn=Button(Downframe,image=grafico8, text="Condições de Pagamento",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=315,y=215)       
      
      grafico9 = PhotoImage(file="Resource\Cadastros\patrimonio.png")
      grafico9= grafico9.subsample(2,2)
      figura = Label(image=grafico9)
      figura.image=grafico9
      if(usuarioLogado.getPermissoes().perm_cadfin_benpatrimonios == True):       
       btn=Button(Downframe,image=grafico9, text="Bens e Patrimonios",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=465,y=215)        
          
      grafico10 = PhotoImage(file="Resource\Cadastros\Situações  Tributárias.png")
      grafico10= grafico10.subsample(2,2)
      figura = Label(image=grafico10)
      figura.image=grafico10
      if(usuarioLogado.getPermissoes().perm_cadfin_situtributarias== True):       
       btn=Button(Downframe,image=grafico10, text="Situações Tributárias",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=575,y=215)  
      
      grafico11 = PhotoImage(file="Resource\Cadastros\Classificações Fiscais.png")
      grafico11= grafico11.subsample(2,2)
      figura = Label(image=grafico11)
      figura.image=grafico11
      if(usuarioLogado.getPermissoes().perm_cadfin_classfiscais == True):       
       btn=Button(Downframe,image=grafico11, text="Classificações Fiscais(NCM)",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=695,y=215)   
      
      grafico12 = PhotoImage(file="Resource\Cadastros\Condições  de pagamento.png")
      grafico12= grafico12.subsample(2,2)
      figura = Label(image=grafico12)
      figura.image=grafico12
      if(usuarioLogado.getPermissoes().perm_cadfin_formspagamento == True):       
       btn=Button(Downframe,image=grafico12, text="Formas De Pagamento",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=850,y=215)                         
      
      def AlterarLogin():
         new.destroy()
         jan.deiconify()
         return 
      Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
      Backbutton.place(x=920, y=380)      
      def Voltando():
         new.destroy()
         jan.destroy()
         return 
      Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
      Backbutton.place(x=920, y=420)  

   def button_hover(e):
       btn1["bg"]="lightblue"
   def buttonhoverleave(e):
       btn1["bg"]="#040D9E"      
   Financeiro1 = PhotoImage(file="Resource\Financeiro\Cadastro.png")
   Financeiro1 = Financeiro1.subsample(2,2)
   figura = Label(image=Financeiro1)
   figura.image=Financeiro1
   if(usuarioLogado.getPermissoes().perm_cadfin == True):        
    btn1=Button(Upframe,image=Financeiro1, text="Cadastros Financeiros",bd=0, highlightthickness=0, bg="#040D9E", fg="white",  compound = TOP, relief="solid", command=gestDeCadastro) 
    btn1.bind("<Enter>", button_hover)  
    btn1.bind("<Leave>",buttonhoverleave)  
   btn1.place(x=221,y=120)
   
   
   def ContasaPagar():
    limpaWidget.limpar(Downframe)      
    def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
    Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
    Backbutton.place(x=920, y=380)      
    def Voltando():
       new.destroy()
       jan.destroy()
       return 
    Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
    Backbutton.place(x=920, y=420)   
   def button_hover(e):
       btn2["bg"]="lightblue"
   def buttonhoverleave(e):
       btn2["bg"]="#040D9E"     
   Financeiro2 = PhotoImage(file="Resource\Financeiro\Contas a Pagar.png")
   Financeiro2 = Financeiro2.subsample(2,2)
   figura = Label(image=Financeiro2)
   figura.image=Financeiro2
   if(usuarioLogado.getPermissoes().perm_contpagar == True):      
    btn2=Button(Upframe,image=Financeiro2, text="Contas A Pagar",bd=0, highlightthickness=0, bg="#040D9E", fg="white", compound = TOP, relief="solid", command= ContasaPagar)
    btn2.bind("<Enter>", button_hover)  
    btn2.bind("<Leave>",buttonhoverleave)   
   btn2.place(x=346,y=120)
 
   def ContasaReceber():
    limpaWidget.limpar(Downframe)      
    def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
    Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
    Backbutton.place(x=920, y=380)      
    def Voltando():
       new.destroy()
       jan.destroy()
       return 
    Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
    Backbutton.place(x=920, y=420)
   def button_hover(e):
       btn3["bg"]="lightblue"
   def buttonhoverleave(e):
       btn3["bg"]="#040D9E"     
   Financeiro3 = PhotoImage(file="Resource\Financeiro\Contas a receber.png")
   Financeiro3 = Financeiro3.subsample(2,2)
   figura = Label(image=Financeiro3)
   figura.image=Financeiro3
   if(usuarioLogado.getPermissoes().perm_contreceber == True):     
    btn3=Button(Upframe,image=Financeiro3, text="Contas a Receber",bd=0, highlightthickness=0, bg="#040D9E", fg="white", compound = TOP, relief="solid", command= ContasaReceber)
    btn3.bind("<Enter>", button_hover)  
    btn3.bind("<Leave>",buttonhoverleave)  
   btn3.place(x=441,y=120)
   
   def SaidasMonetarias():
    limpaWidget.limpar(Downframe)      
    def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
    Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
    Backbutton.place(x=920, y=380)      
    def Voltando():
       new.destroy()
       jan.destroy()
       return 
    Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
    Backbutton.place(x=920, y=420)   
   def button_hover(e):
       btn4["bg"]="lightblue"
   def buttonhoverleave(e):
       btn4["bg"]="#040D9E"             
   Financeiro4 = PhotoImage(file="Resource\Financeiro\Saidas monetarias.png")
   Financeiro4 = Financeiro4.subsample(2,2)
   figura = Label(image=Financeiro4)
   figura.image=Financeiro4
   if(usuarioLogado.getPermissoes().perm_saimonetarias == True):    
    btn4=Button(Upframe,image=Financeiro4, text="Saidas Monetárias",bd=0, highlightthickness=0, bg="#040D9E", fg="white", compound = TOP, relief="solid", command= SaidasMonetarias)
    btn4.bind("<Enter>", button_hover)  
    btn4.bind("<Leave>",buttonhoverleave)     
   btn4.place(x=541,y=120)
   
   def EntradasMonetarias():
    limpaWidget.limpar(Downframe)      
    def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
    Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
    Backbutton.place(x=920, y=380)      
    def Voltando():
       new.destroy()
       jan.destroy()
       return 
    Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
    Backbutton.place(x=920, y=420)
   def button_hover(e):
       btn5["bg"]="lightblue"
   def buttonhoverleave(e):
       btn5["bg"]="#040D9E"     
   Financeiro5 = PhotoImage(file="Resource\Financeiro\Entradas Monetárias.png")
   Financeiro5 = Financeiro5.subsample(2,2)
   figura = Label(image=Financeiro5)
   figura.image=Financeiro5
   if(usuarioLogado.getPermissoes().perm_entmonetarias == True):      
    btn5=Button(Upframe,image=Financeiro5, text="Entradas Monetárias",bd=0, highlightthickness=0, bg="#040D9E", fg="white", compound = TOP, relief="solid", command= EntradasMonetarias)
    btn5.bind("<Enter>", button_hover)  
    btn5.bind("<Leave>",buttonhoverleave)   
   btn5.place(x=651,y=120)
   
   def cheques():
    limpaWidget.limpar(Downframe)      
    def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
    Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
    Backbutton.place(x=920, y=380)      
    def Voltando():
       new.destroy()
       jan.destroy()
       return 
    Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
    Backbutton.place(x=920, y=420)
   def button_hover(e):
       btn6["bg"]="lightblue"
   def buttonhoverleave(e):
       btn6["bg"]="#040D9E"     
   Financeiro6 = PhotoImage(file="Resource\Financeiro\Cheques 2.png")
   Financeiro6 = Financeiro6.subsample(2,2)
   figura = Label(image=Financeiro6)
   figura.image=Financeiro6
   if(usuarioLogado.getPermissoes().perm_cheques == True):      
    btn6=Button(Upframe,image=Financeiro6, text="Cheques",bd=0, highlightthickness=0, bg="#040D9E", fg="white",  compound = TOP, relief="solid", command= cheques)
    btn6.bind("<Enter>", button_hover)  
    btn6.bind("<Leave>",buttonhoverleave)           
   btn6.place(x=781,y=120)  
   def gestFinanceira():
 
      def aperteEnter(event):
         event.widget.config( relief = GROOVE )
         btn.bind( "<Enter>", command=aperteEnter)
         
      limpaWidget.limpar(Downframe)   
      grafico1 = PhotoImage(file="Resource\Gest\Extrato de contas.png")
      grafico1= grafico1.subsample(2,2)
      figura = Label(image=grafico1)
      figura.image=grafico1
      if(usuarioLogado.getPermissoes().perm_gestfin_extcontas == True):      
       btn=Button(Downframe,image=grafico1, text="Extrato  de  Contas",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=150, y=180)
      
      grafico2 = PhotoImage(file="Resource\Gest\Fluxo de Caixa.png")
      grafico2= grafico2.subsample(2,2)
      figura = Label(image=grafico2)
      figura.image=grafico2
      if(usuarioLogado.getPermissoes().perm_gestfin_fluxocaixa == True):
       btn=Button(Downframe,image=grafico2, text="Fluxo de caixa",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=270, y=180)
      
      grafico3 = PhotoImage(file="Resource\Gest\Receitas  e Entradas.png")
      grafico3= grafico3.subsample(2,2)
      figura = Label(image=grafico3)
      figura.image=grafico3
      if(usuarioLogado.getPermissoes().perm_gestfin_receitaseentradas == True):
       btn=Button(Downframe,image=grafico3, text="Receitas e Entradas",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=360, y=180)
      
      grafico4 = PhotoImage(file="Resource\Gest\Despesas.png")
      grafico4= grafico4.subsample(2,2)
      figura = Label(image=grafico4)
      figura.image=grafico4
      if(usuarioLogado.getPermissoes().perm_gestfin_despesas == True):
       btn=Button(Downframe,image=grafico4, text="Despesas",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=470, y=180)
      
      grafico5 = PhotoImage(file="Resource\Gest\Despesas centro de custo.png")
      grafico5= grafico5.subsample(2,2)
      figura = Label(image=grafico5)
      figura.image=grafico5
      if(usuarioLogado.getPermissoes().perm_gestfin_despcentrodecusto == True):
       btn=Button(Downframe,image=grafico5, text="Desp. por Centro de Custo",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=550, y=180)
      
      grafico6 = PhotoImage(file="Resource\Gest\Posição Monetária  e financeira.png")
      grafico6= grafico6.subsample(2,2)
      figura = Label(image=grafico6)
      figura.image=grafico6
      if(usuarioLogado.getPermissoes().perm_gestfin_posmonetfin == True):
       btn=Button(Downframe,image=grafico6, text="Pos. Monetária e Financeira",bd=0, highlightthickness=0, bg="white", compound = TOP)
      btn.place(x=700, y=180)
      
      def ResultadoFinanceiro():

        Fin = tk.Toplevel()
        Fin.title("Resultado financeiro Por Competência")
        Fin.geometry("300x300")
        Fin.configure(background="white")
        Fin.resizable(width=True, height=False)
        Fin.attributes("-alpha", 1.1)
        Fin.grid()
        ValorLabel = Label(Fin, text="Depreciação", font=("Calibri",11), bg="white", fg="black")
        ValorLabel.place(x=1,y=25)
        ValorEntry = ttk.Entry(Fin, width=20)
        ValorEntry.place(x=85,y=25)  
        def Salva():
            valor=ValorEntry.get() 
            vendaService.gerarDepreciacao(valor)
            if ValorEntry != cadastroDepreciacaobutton:
              messagebox.showinfo(title = "Cadastrado", message="Depreciação cadastrada com sucesso")

            else:
              messagebox.showerror(title = "ERRO", message="Não pode ser cadastrado")
        cadastroDepreciacaobutton= ttk.Button(Fin, text="Inserir", width=10, command=Salva)
        cadastroDepreciacaobutton.bind("<KeyPress>",lambda e:Salva())
        cadastroDepreciacaobutton.place(x=215, y=25) 

        
        
        
        
        
        def cadasstroWoodpaper():
            cadastroWood=Tk()
            cadastroWood.title("cadastro WoodPaper")
            cadastroWood.geometry("500x500")
            cadastroWood.configure(background="lightblue")
            cadastroWood.resizable(width=True, height=False)
            cadastroWood.attributes("-alpha", 1.0)
            cadastroWoodupframe = Frame(cadastroWood, width=9999, height=9999,bg="lightgrey", relief="raise")
            cadastroWoodupframe.grid() 
            Fin.withdraw()
            AnoLabel = Label(cadastroWood, text="Ano", font=("Calibri",10), bg="lightgrey", fg="black")
            AnoLabel.place(x=10,y=55)
            AnoEntry = ttk.Entry(cadastroWood, width=15)
            AnoEntry.place(x=70,y=55)
            JaneiroLabel = Label(cadastroWood, text="Janeiro", font=("Calibri",10), bg="lightgrey", fg="black")
            JaneiroLabel.place(x=10,y=75)
            JaneiroEntry = ttk.Entry(cadastroWood, width=20)
            JaneiroEntry.place(x=70,y=75)
            FevereiroLabel = Label(cadastroWood, text="Fevereiro", font=("Calibri",10), bg="lightgrey", fg="black")
            FevereiroLabel.place(x= 10,y=95)
            FevereiroEntry = ttk.Entry(cadastroWood, width=20)
            FevereiroEntry.place(x=70,y=95)
            MarcoLabel = Label(cadastroWood, text="Marco", font=("Calibri",10), bg="lightgrey", fg="black")
            MarcoLabel.place(x= 10,y=115)
            MarcoEntry = ttk.Entry(cadastroWood, width=20)
            MarcoEntry.place(x=70,y=115)
            AbrilLabel = Label(cadastroWood, text="Abril", font=("Calibri",10), bg="lightgrey", fg="black")
            AbrilLabel.place(x= 10,y=135)
            AbrilEntry = ttk.Entry(cadastroWood, width=20)
            AbrilEntry.place(x=70,y=135)
            MaioLabel = Label(cadastroWood, text="Maio", font=("Calibri",10), bg="lightgrey", fg="black")
            MaioLabel.place(x= 10,y=155)
            MaioEntry = ttk.Entry(cadastroWood, width=20)
            MaioEntry.place(x=70,y=155)
            JunhoLabel = Label(cadastroWood, text="junho", font=("Calibri",10), bg="lightgrey", fg="black")
            JunhoLabel.place(x= 10,y=175)
            JunhoEntry = ttk.Entry(cadastroWood, width=20)
            JunhoEntry.place(x=70,y=175)
            JulhoLabel = Label(cadastroWood, text="julho", font=("Calibri",10), bg="lightgrey", fg="black")
            JulhoLabel.place(x= 10,y=195)
            JulhoEntry = ttk.Entry(cadastroWood, width=20)
            JulhoEntry.place(x=70,y=195)
            AgostoLabel = Label(cadastroWood, text="Agosto", font=("Calibri",10), bg="lightgrey", fg="black")
            AgostoLabel.place(x= 10,y=215)
            AgostoEntry = ttk.Entry(cadastroWood, width=20)
            AgostoEntry.place(x=70,y=215)
            SetembroLabel = Label(cadastroWood, text="Setembro", font=("Calibri",10), bg="lightgrey", fg="black")
            SetembroLabel.place(x= 10,y=235)
            SetembroEntry = ttk.Entry(cadastroWood, width=20)
            SetembroEntry.place(x=70,y=235)
            OutubroLabel = Label(cadastroWood, text="Outubro", font=("Calibri",10), bg="lightgrey", fg="black")
            OutubroLabel.place(x= 10,y=255)
            OutubroEntry = ttk.Entry(cadastroWood, width=20)
            OutubroEntry.place(x=70,y=255)  
            NovembroLabel = Label(cadastroWood, text="Novembro", font=("Calibri",10), bg="lightgrey", fg="black")
            NovembroLabel.place(x= 10,y=275)
            NovembroEntry = ttk.Entry(cadastroWood, width=20)
            NovembroEntry.place(x=70,y=275)
            DezembroLabel = Label(cadastroWood, text="Dezembro", font=("Calibri",10), bg="lightgrey", fg="black")
            DezembroLabel.place(x= 10,y=295)
            DezembroEntry = ttk.Entry(cadastroWood, width=20)
            DezembroEntry.place(x=70,y=295)         
    
    
            def Salvo():
               listaDeWoodaper=[]
               Ano = AnoEntry.get()
               
               if JaneiroEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-01-01", JaneiroEntry.get())
                  listaDeWoodaper.append(woodpaper)    
               
               if FevereiroEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-02-01", FevereiroEntry.get())
                  listaDeWoodaper.append(woodpaper)   
                  
               if MarcoEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-03-01", MarcoEntry.get())
                  listaDeWoodaper.append(woodpaper)                                
               if AbrilEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-04-01", AbrilEntry.get())
                  listaDeWoodaper.append(woodpaper)
               if MaioEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-05-01", MaioEntry.get())
                  listaDeWoodaper.append(woodpaper)                  
               if JunhoEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-06-01", JunhoEntry.get())
                  listaDeWoodaper.append(woodpaper)      
               if JulhoEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-07-01", JulhoEntry.get())
                  listaDeWoodaper.append(woodpaper)
               if AgostoEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-08-01", AgostoEntry.get())
                  listaDeWoodaper.append(woodpaper)
               if SetembroEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-09-01", SetembroEntry.get())
                  listaDeWoodaper.append(woodpaper)
               if OutubroEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-10-01", OutubroEntry.get())
                  listaDeWoodaper.append(woodpaper)
               if NovembroEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-11-01", NovembroEntry.get())
                  listaDeWoodaper.append(woodpaper)
               if DezembroEntry.get() != '':
                  woodpaper = Woodpaper(AnoEntry.get()+"-12-01", DezembroEntry.get())
                  listaDeWoodaper.append(woodpaper)
                                    
               vendaService.gerarCadastroWoodpaper(listaDeWoodaper)
               messagebox.showinfo(title = "cadastrado", message="cadastrado com Sucesso")
               
            cadastroWoodastrobutton= ttk.Button(cadastroWood, text="cadastrar", width=10, command=Salvo)
            cadastroWoodastrobutton.bind ("<KeyPress>", lambda e: Salvo()) 
            cadastroWoodastrobutton.place(x=50, y=400) 
            def sairWood():
             cadastroWood.destroy()
             return ResultadoFinanceiro()
            sairdaWoodbutton= ttk.Button(cadastroWood, text="Sair", width=10, command=sairWood)
            sairdaWoodbutton.place(x=150, y=400) 
        ok_btn=ttk.Button(Fin, text="Woodpaper", command=cadasstroWoodpaper)
        ok_btn.place(x=1, y=50)                 
 
        

        def relatorioCompetencia():
           vendaService.ApagarValorCusto() 
           Competencia = tk.Toplevel()
           Competencia.title("Relatorio de vendas por periodo")
           Competencia.configure(background="white")
           Competencia.attributes("-alpha", 1.1)
           Competencia.grid()
           Fin.withdraw()
           screen_width = Competencia.winfo_screenwidth()
           screen_height = Competencia.winfo_screenheight()

            # set window size and position based on screen size
           width_ratio = 0.8  # adjust as desired
           height_ratio = 0.7  # adjust as desired
           x_pos = int((screen_width - (screen_width * width_ratio)) / 2)
           y_pos = int((screen_height - (screen_height * height_ratio)) / 2)
           width = int(screen_width * width_ratio)
           height = int(screen_height * height_ratio)

            # set window size and position
           Competencia.geometry(f"{width}x{height}+{x_pos}+{y_pos}")
           Competencia.resizable(width=True, height=True)           
           
           Compframe = Frame(Competencia, width=99999, height=135, bg="Grey", relief="raise")
           Compframe.pack(side=TOP)
           Periodo = tk.Label(Compframe,text="Competência de:", font=("Calibri",8), bg="White", fg="black")
           Periodo.place(x=1 ,y=105) 
           sv = StringVar()
           periodoEntry=tk.Entry(Compframe,  width=10, textvariable=sv)
           sv.trace("w", lambda name, index, mode, sv=sv: entryUpdateEndHour(periodoEntry))
           def entryUpdateEndHour(Entry):
              text = periodoEntry.get()
              if len(text) in (2,7):
                Entry.insert(END,'/')
                Entry.icursor(len(text)+1)
              if len(text) > 6:
                 Entry.delete(0,END)
                 Entry.insert(0,text[:7])
           periodoEntry.place(x=91, y=104) 
           sv2 = StringVar()              
           Competencia2 = tk.Label(Compframe, text="Competência até", font=("Calibri",8), bg="White", fg="black")
           Competencia2.place(x=191 ,y=105)
           competenciaEntry = tk.Entry(Compframe, width=10,textvariable=sv2 )
           sv2.trace("w", lambda name, index, mode, sv2=sv2: entryUpdateEndHour2(competenciaEntry))
           def entryUpdateEndHour2(Entry):
              text = competenciaEntry.get()
              if len(text) in (2,7):
                Entry.insert(END,'/')
                Entry.icursor(len(text)+1)
              if len(text) > 6:
                 Entry.delete(0,END)
                 Entry.insert(0,text[:7])           
           competenciaEntry.place(x=281,y=104)
           competenciaEntry.bind("<KeyPress>", lambda e: TabelaCompetencia() if e.char == '\r' else None)

            
           tree=ttk.Treeview(Competencia, selectmode="browse", height=21, show='headings')

           tree.tag_configure('linhaimpar', background="white")
           tree.tag_configure('linhapar', background="lightblue")
           tree.tag_configure('linhatitulo', background="lightgrey") 
           tree.place(x=1,y=138)
           treescrollbar = ttk.Scrollbar(Competencia, orient="vertical", command=tree.yview)
           treescrollbar.pack(expand = 0, side='right', fill='y')
           tree.configure(yscrollcommand=treescrollbar.set)
           treescrollbar = ttk.Scrollbar(Competencia, orient="horizontal", command=tree.xview)
           treescrollbar.pack(expand = 0, side='bottom', fill='x')
           tree.configure(xscrollcommand=treescrollbar.set)
           quantMes=[]
           def TabelaCompetencia():  
                 
                 periodoCompete = datetime.strptime(periodoEntry.get(), '%m/%Y')#seta a string como data
                 #digits = filter(str.isdigit, PeriodoEntry.get())
                 #PeriodoEntry.set('{}/{}'.format(digits[2]/digits[4]))
                 #PeriodoEntry.icursor(END)

                 competenciaCompete = datetime.strptime(competenciaEntry.get(), '%m/%Y')
                 #digits = filter(str.isdigit, competenciaEntry.get())   
                 #competenciaEntry.set('{}/{}'.format(digits[2]/digits[4]))                               
                 #competenciaEntry.icursor(END)
                 if periodoCompete.year != competenciaCompete.year:
                      messagebox.showerror('ERRO', 'Você não pode fazer essa operação')#permite  o erro  em  caso  de uma busca maior que a planejada
                      return Competencia

      
  


                 def gerarRelatorioPorCompetencia(): 
                     quantMes.clear()
                     testeLabel = Progressbar(Competencia, orient = 'horizontal', length = 250, mode = 'determinate')
                     testeLabel.place(x=600,y=500)                    
                     testeLabel['value'] = 25
                     Competencia.update_idletasks()
                     time.sleep(1)                    
                    
                     resultSet = vendaService.gerarTabelaPorCompete(periodoCompete.year, periodoCompete.month, competenciaCompete.month)
                     tamanho = int(len(resultSet))
                     
                     
                     my_list = []
                     tamanhoColunas = tamanho + 4 
                     tree['columns'] =list(range(1, tamanhoColunas))
                     
                     tree.column("1", width=150, minwidth=60,stretch=NO)
                     tree.heading("#1", text='Índices')     
                     quantMes.append('Indices') 
                     



  
                     
                     #GeraOsMeses
                     posicao=0
                     for j in range(0, tamanho, 1):
                        posicao=j+2
                        result = resultSet[j]
                        tree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                        tree.heading("#"+str(posicao), text=result.getNomeMes())
                        quantMes.append(result.getNomeMes())
                     
                     posicao+=1   
                     tree.column(int(posicao), width=100, minwidth=50,stretch=NO)
                     tree.heading("#"+str(posicao), text='Média')
                     quantMes.append('Média')
                     posicao+=1 
                     tree.column(int(posicao), width=100, minwidth=50,stretch=NO)
                     tree.heading("#"+str(posicao), text='Total')
                     quantMes.append('Total')
                     
                     

                     for i in tree.get_children():
                        tree.delete(i)
                     

                     tree.insert('', 'end',  values=vendaService.getLinhaVendas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendasCNVendas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendasCNServicos(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendas_S_N(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaTotalDeVendas(),tags=('linhatitulo'))                 
                     tree.insert('', 'end',  values=vendaService.getLinhaSimples(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaTotalLiq(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaVendas_Devolvidas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendas_Canceladas(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaCortesias(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDescErroOperacional(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDescErroComercial(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaComissoesInternas(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaComissoesExternas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaTotalVendasGeral(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values='',tags=('linhaimpar'))
                     tree.insert('', 'end',  values=('Despesas',),tags=('linhatitulo')) #CONSERTAR AQUI
                     tree.insert('', 'end',  values=vendaService.getLinhaProLaboreGerencia(),tags=('linhaimpar'))  
                    
                     tree.insert('', 'end',  values=('ESTOQUE',),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaLaminas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMadeiras(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVazadores(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaBorrachas(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaPertinax(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMaterialPDestaque(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMaquinasParaRevenda(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaOutros(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaServicosExternos(),tags=('linhaimpar'))
                     
                     tree.insert('', 'end',  values='',tags=('linhapar'))                     
                     tree.insert('', 'end',  values=('CLIENTES',),tags=('linhatitulo'))                     
                     tree.insert('', 'end',  values=vendaService.getLinhaLaminasClientes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMadeirasClientes(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVazadoresClientes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaBorrachasClientes(),tags=('linhaipar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaPertinaxClientes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMaterialPDestaqueClientes(),tags=('linhapar'))                     
                     tree.insert('', 'end',  values=vendaService.getLinhaCliches(),tags=('linhaimpar'))  
                                          
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDesenho(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosDesenho(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHE(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestes(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscates(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadas(),tags=('linhapar'))                     
                     tree.insert('', 'end',  values=vendaService.getLinhausoeConsumoDesenho(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhamanutencao(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaPlotter(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaUsoeConsumoPlotter(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManutencaoPlotter(),tags=('linhaimpar'))

                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaLaser(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosLaser(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHELaser(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestesLaser(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscatesLaser(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadasLaser(),tags=('linhapar'))                      
                     tree.insert('', 'end',  values=vendaService.getLinhaUsoeConsumoLaser(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManutencaoLaser(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaGaspLaser(),tags=('linhapar'))  
                     tree.insert('', 'end',  values='',tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMontagem(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosMontagem(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHEMontagem(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestesMontagem(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscatesMontagem(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadasMontagem(),tags=('linhaimpar'))                        
                     tree.insert('', 'end',  values=vendaService.getLinhausoEconsumoMontagem(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManuTencaoMontagem(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaBorracha(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosBorracha(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHEBorracha(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestesBorracha(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscatesBorracha(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadasBorracha(),tags=('linhaimpar'))                              
                     tree.insert('', 'end',  values=vendaService.getLinhaUsoeConsumoBorracha(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManutencaoBorracha(),tags=('linhapar'))   
                     tree.insert('', 'end',  values='',tags=('linhapar')) 
                     tree.insert('', 'end',  values=vendaService.getLinhatotalGeralServiços(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values='',tags=('linhapar')) 
                     tree.insert('', 'end',  values=vendaService.getLinhatotaldeResultados1(),tags=('linhatitulo')) 
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDepreciacao(),tags=('linhatitulo')) 
                     tree.insert('', 'end',  values='',tags=('linhapar')) 
                     tree.insert('', 'end',  values=vendaService.getLinhatotaldeResultados2(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaDespesasAdministrativas(),tags=('linhaimpar'))                      
                     tree.insert('', 'end',  values=vendaService.getLinhaDespesasCentrais(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosAdm(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosComercial(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosTransporte(),tags=('linhaimpar'))   
                     tree.insert('', 'end',  values='',tags=('linhapar'))                      
                     tree.insert('', 'end',  values=vendaService.getLinhatotalGeralAdministrativo(),tags=('linhatitulo'))    
                     tree.insert('', 'end',  values=vendaService.getLinhaDespesasFinanceiras(),tags=('linhaimpar'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaJuros(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaJurosSobreAplicacoes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaJurosSobreVendas(),tags=('linhapar'))
                     tree.insert('', 'end',  values='',tags=('linhaimpar'))                                                    
                     tree.insert('', 'end',  values=vendaService.getLinhatotalResultadoEzipa(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getlinhatotalWoodpaper(),tags=('linhaimpar'))
                     
                     testeLabel['value'] = 75
                     Competencia.update_idletasks()
                     time.sleep(1)                        
                            
                     testeLabel['value'] = 100   
                     testeLabel.destroy()
                     if testeLabel == 100:
                        return messagebox.showinfo("RELATÓRIO", "Relatório gerado com sucesso")

 
                     return                 
                 threading.Thread(target=gerarRelatorioPorCompetencia).start()

     

           ok_btn=ttk.Button(Competencia, text="ok", command=TabelaCompetencia)
           ok_btn.place(x=381, y=103)    

                  

           
           def competenciaExcel():
  
            #leitura=Leitura()
            if(os.path.exists('temporario.csv')):
               os.remove('temporario.csv')            
                      
            colunas = quantMes
            path= 'temporario.csv'
            excel_name=tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Excel','*.xlsx')],filetypes=[('Excel','*.xlsx')])#Relatorio_Ezipa_Competencia_17-08-2022
            if not excel_name or excel_name == '/': # If the user closes the dialog without choosing location
               messagebox.showerror('Error','Choose a location to save')
               return 
            #leitura.excel2+ str( date.today()) +'Relatorio Financeiro Competencia.xlsx' #Relatorio_Ezipa_Competencia_17-08-2022
            lista=[]
            with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                     csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                     for row_id in tree.get_children():
                        row = tree.item(row_id,'values')
                        lista.append(row)
                     lista = list(map(list,lista)) #faz um mapa das linhas
                     lista.insert(0,colunas)  #insere as colunas
                     for row in lista:
                        csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário
            writer = pd.ExcelWriter(excel_name) #cria um novo arquivo excel 
            df = pd.read_csv(path, encoding='latin') #define o encoding (nota: estavamos com um problema de utf-8, isso pode ser um problema no futuro...)
            df.to_excel(writer,'Analitico',index=False) #escreve as informações no excel com o nome interno da planilha pre-definido
            writer.save() #efetivamente cria o arquivo excel

           excel_btn= PhotoImage(file="Resource\Financeiro\gerarExcel.png")
           excel_btn = excel_btn.subsample(2,2)
           figura = Label(image=excel_btn)
           figura.image=excel_btn
           excelbtn=Button(Competencia,image=excel_btn, text="Excel",bd=0, highlightthickness=0, bg="grey", compound = TOP, command=competenciaExcel)
           excelbtn.place(x=1, y=14)
           
           def competenciaPdf():
            #leitura=Leitura()
            if(os.path.exists('temporario.csv')):
               os.remove('temporario.csv')
            Caminhopdf2= tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Arquivo Pdf','*.pdf')],filetypes=[('Arquivo pdf','*.pdf')])#Relatorio_Ezipa_Competencia_17-08-2022
            if not Caminhopdf2 or Caminhopdf2 == '/': # If the user closes the dialog without choosing location
               messagebox.showerror('Error','Choose a location to save')
               return # Stop the function    
            #leitura.PDF2 + str( date.today())+'Relatorio Financeiro Competencia.pdf'
              
            colunas = quantMes  
            path= 'temporario.csv'
            lista=[]
            with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
               csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
               for row_id in tree.get_children():
                  row = tree.item(row_id,'values')
                  lista.append(row)
               lista = list(map(list,lista)) #faz um mapa das linhas
               lista.insert(0,colunas)  #insere as colunas
               for row in lista:
                  csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário   
            a = pd.read_csv(path,  encoding="latin-1") #le o csv
            a.to_html('Relatório por Competencia.htm') #cria o html que  gerará o pdf
            html_file = a.to_html() 
            
            path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'#necessita para a criação do pdf
            config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
            options={'encoding': "UTF-8" ,  'cookie': [('cookie-empty-value', '""')]}#cria  opções para o pdf
            pdfkit.from_file('relatório por Competencia.htm', Caminhopdf2, configuration=config, options=options)#gera o caminho do  pdf
           pdf_btn= PhotoImage(file="Resource\Financeiro\gerarpdf.png")
           pdf_btn = pdf_btn.subsample(2,2)
           figura = Label(image=pdf_btn)
           figura.image=pdf_btn
           pdf_btn=Button(Competencia,image=pdf_btn,  text="Pdf",bd=0, highlightthickness=0, bg="grey", compound = TOP,command=competenciaPdf)
           pdf_btn.place(x=81, y=14)  
           
           def imprimirCompetencia():
             answer= messagebox.askyesno("Impressão","Deseja Imprimir") 
             if answer == YES:  
             #Leitura=Leiturapdf()
               if(os.path.exists('temporario.csv')):
                  os.remove('temporario.csv')
               # Stop the function    
               Caminhopdf2= 'Relatorio Financeiro Competencia.pdf'
               #Leitura.PDF2 + str( date.today())+'Relatorio Financeiro Competencia.pdf'
               
               colunas = quantMes  
               path= 'temporario.csv'
               lista=[]
               with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                  csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                  for row_id in tree.get_children():
                     row = tree.item(row_id,'values')
                     lista.append(row)
                  lista = list(map(list,lista)) #faz um mapa das linhas
                  lista.insert(0,colunas)  #insere as colunas
                  for row in lista:
                     csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário   
               a = pd.read_csv(path,  encoding="latin-1") #le o csv
               a.to_html('Impressao.html') #cria o html que  gerará o pdf
               html_file = a.to_html() 
               path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'#necessita para a criação do pdf
               config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
               options={'encoding': "UTF-8" ,  'cookie': [('cookie-empty-value', '""')]}#cria  opções para o pdf
               pdfkit.from_file('impressao.html', Caminhopdf2, configuration=config, options=options)#gera o caminho do  pdf               
               

               filename = Caminhopdf2
               win32api.ShellExecute (0,"print",filename,'/d:"%s"' % win32print.GetDefaultPrinter (),".",0)                              
             else:
                return Competencia
           print_btn= PhotoImage(file="Resource\Financeiro\gerarimpressão.png")
           print_btn= print_btn.subsample(2,2)
           figura = Label(image=print_btn)
           figura.image=print_btn
           print_btn=Button(Competencia, image=print_btn, text="Imprimir", bd=0, highlightthickness=0, bg="grey", compound = TOP, command=imprimirCompetencia)
           print_btn.place(x=161,y=14) 
           
           def gerarAnaliticoAdministrativo(): 
             answer= messagebox.askyesno("Analítico Admiministrativo","Deseja Gerar Analítico Administrativo") 
             if answer == YES:

               AnaliticoCompetencia = tk.Toplevel()
               titulo = f"Analitico ADM de {periodoEntry.get()} até {competenciaEntry.get()}"
               AnaliticoCompetencia.title(titulo)
               AnaliticoCompetencia.geometry("950x750+350+10")
               AnaliticoCompetencia.configure(background="white")
               AnaliticoCompetencia.resizable(width=True, height=True)
               AnaliticoCompetencia.attributes("-alpha", 1.1)
               Analiticoframe = Frame(AnaliticoCompetencia, width=99999, height=135, bg="#8c8c8c", relief="raise")
               Analiticoframe.pack(side=TOP)
               AnaliticoCompetencia.grid
               screen_width = AnaliticoCompetencia.winfo_screenwidth()
               screen_height = AnaliticoCompetencia.winfo_screenheight()

                  # set window size and position based on screen size
               width_ratio = 0.8  # adjust as desired
               height_ratio = 0.7  # adjust as desired
               x_pos = int((screen_width - (screen_width * width_ratio)) / 2)
               y_pos = int((screen_height - (screen_height * height_ratio)) / 2)
               width = int(screen_width * width_ratio)
               height = int(screen_height * height_ratio)

                  # set window size and position
               AnaliticoCompetencia.geometry(f"{width}x{height}+{x_pos}+{y_pos}")
               AnaliticoCompetencia.resizable(width=True, height=True)                
               
               
                                                            
               Analiticotree=ttk.Treeview(AnaliticoCompetencia, selectmode="browse",  height=23, show='headings')
               Analiticotree.tag_configure('linhaimpar', background="white")
               Analiticotree.tag_configure('linhapar', background="lightblue")
               Analiticotree.tag_configure('linhatitulo', background="lightyellow")
               treescrollbar = ttk.Scrollbar(AnaliticoCompetencia, orient="vertical", command=Analiticotree.yview)
               treescrollbar.pack(expand = 0, side='right', fill='y')
               Analiticotree.configure(yscrollcommand=treescrollbar.set) 
               treescrollbar = ttk.Scrollbar(AnaliticoCompetencia, orient="horizontal", command=Analiticotree.xview)
               treescrollbar.pack(expand = 0, side='bottom', fill='x')
               Analiticotree.configure(xscrollcommand=treescrollbar.set) 
               quantMes=[]
                  
               def gerarTabelaAnalitico():
                     quantMes.clear()
                     analiticoService.Apagaranaliticoadm()
                     analiticoService.ApagarValorAnalitico()                     
                     testeLabel = Progressbar(AnaliticoCompetencia,  orient = 'horizontal', length = 250, mode = 'determinate')
                     testeLabel.place(x=600,y=500)    
                     testeLabel['value'] = 25
                     AnaliticoCompetencia.update_idletasks()
                     time.sleep(1)
                     
                     periodoCompete = datetime.strptime(periodoEntry.get(), '%m/%Y')#seta a string como data
                     #digits = filter(str.isdigit, PeriodoEntry.get())
                     #PeriodoEntry.set('{}/{}'.format(digits[2]/digits[4]))
                     #PeriodoEntry.icursor(END)

                     competenciaCompete = datetime.strptime(competenciaEntry.get(), '%m/%Y')
                     #digits = filter(str.isdigit, competenciaEntry.get())   
                     #competenciaEntry.set('{}/{}'.format(digits[2]/digits[4]))                               
                     #competenciaEntry.icursor(END)
                     resultado = analiticoService.GerarRelatorioDetalhadoCompetenciaAnalitico(periodoCompete.year, periodoCompete.month, competenciaCompete.month)
                     if periodoCompete.year != competenciaCompete.year:
                           messagebox.showerror('ERRO', 'Você não pode fazer essa operação')#permite  o erro  em  caso  de uma busca maior que a planejada
                           return Competencia

                           
                           
                     resultSet = analiticoService.GerarRelatorioAnaliticoCompetencia(periodoCompete.year, periodoCompete.month, competenciaCompete.month)                                     
                     tamanho = int(len(resultSet))
                           

                     tamanhoColunas = tamanho + 4
                     Analiticotree['columns'] =list(range(0, tamanhoColunas))
                     
                        
                        #GeraOsMeses
                  
                     
                        
                     Analiticotree.column("1", width=120, minwidth=100,stretch=NO)
                     Analiticotree.heading("#1", text='Nome das Despesas')     
                     quantMes.append('Nome das Despesas') 
                        
                        #GeraOsMeses
                     posicao=1
                     
                     testeLabel['value'] = 50
                     AnaliticoCompetencia.update_idletasks()
                     time.sleep(1)
                     
                     for j in range(0, tamanho, 1):
                        posicao=j+2
                        result = resultSet[j]
                        Analiticotree.column(str(posicao), width=80, minwidth=50,stretch=NO)
                        Analiticotree.heading("#"+str(posicao), text=result.getNomeMes())
                        quantMes.append(result.getNomeMes())                    
                     posicao+=1   
                     Analiticotree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     Analiticotree.heading("#"+str(posicao), text='Média')
                     quantMes.append('Média')
                                      
                     posicao+=1 
                     Analiticotree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     Analiticotree.heading("#"+str(posicao), text='Total')
                     quantMes.append('Total')   
                                          
                     
                        
                     for i in Analiticotree.get_children():
                        Analiticotree.delete(i)           
                                 
                        
                     Analiticotree.place(x=1,y=135)
                     Analiticotree.identify_region(10,100)
                  
                     Analiticotree.insert('', 'end', values=analiticoService.getLinhaAluguel(),tags=('linhaimpar','3'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIptu(),tags=('linhapar','4'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCedae(),tags=('linhimpar','5'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaLight(),tags=('linhapar','6'))                     
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCeg(),tags=('linhaimpar','7'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTelefonia(),tags=('linhapar','8')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPnP(),tags=('linhaimpar','9')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaContabilidade(),tags=('linhapar','11'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSigraf(),tags=('linhaimpar','12')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaInformatica(),tags=('linhapar','14'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMaterialdeEscritorio(),tags=('linhaimpar','15')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMaterialLimpeza(),tags=('linhapar','16')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCartuchosParaImpressoras(),tags=('linhaimpar','17')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOnibusContinuo(),tags=('linhapar','18')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCorreios(),tags=('linhaimpar','19')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDespesasBancarias(),tags=('linhapar','20')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutrosAdministrativos(),tags=('linhaimpar','21')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaKombi(),tags=('linhapar','23')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMotoboy(),tags=('linhaimpar','24')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutrosFretes(),tags=('linhapar','25')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTranpostadoras(),tags=('linhaimpar','26')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSedexRemessaVenda(),tags=('linhapar','27')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPedagioFreteVenda(),tags=('linhaimpar','30')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSimples(),tags=('linhapar','31'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutrosImpostos(),tags=('linhaimpar','34')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSalarios(),tags=('linhapar','35'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaRecisoes(),tags=('linhaimpar','36')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaValeTransporte(),tags=('linhapar','37'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaValeAlimentacao(),tags=('linhaimpar','38')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaInss(),tags=('linhapar','39')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFgts(),tags=('linhaimpar','40')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaAjudaDeCusto(),tags=('linhapar','41')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFerias(),tags=('linhaimpar','42')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaHoraExtra(),tags=('linhapar','43')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTestes(),tags=('linhaimpar','44')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCestaBasica(),tags=('linhapar','45')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaLanches(),tags=('linhaimpar','46'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhapesquisaDesenvolvimento(),tags=('linhapar','68'))                      
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFarmacia(),tags=('linhaimpar','47'))     
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTreinamento(),tags=('linhapar','48')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBiscate(),tags=('linhaimpar','49'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutros(),tags=('linhapar','50')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaEmprestimosInternos(),tags=('linhaimpar','71'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaResultados(),tags=('linhapar','72')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaProLaboreDiretoria(),tags=('linhaimpar','73')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaProLaboreGerencia(),tags=('linhapar','74'))     
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDespesasCadastraisClientes(),tags=('linhaimpar','75')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaExames(),tags=('linhapar','77')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCipa(),tags=('linhaimpar','78'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaServSeguranca(),tags=('linhapar','79'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMinformatica(),tags=('linhaimpar','86'))                      
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPlanoDeSaude(),tags=('linhapar','91'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFeriasTrabalhadas(),tags=('linhaimpar','92'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaEpiUniformes(),tags=('linhapar','93'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFestasEConfraternizações(),tags=('linhaimpar','95'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMPredial(),tags=('linhapar','97'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMEletrica(),tags=('linhaimpar','98'))                    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMArCondicionado(),tags=('linhapar','99'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMBebedouro(),tags=('linhaimpar','100'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMMoveis(),tags=('linhapar','101'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMOutrosManut(),tags=('linhaimpar','102'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMHidraulica(),tags=('linhapar','103'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBMaquinaseEquip(),tags=('linhaimpar','104'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaObrasEMelhorias(),tags=('linhapar','105'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBMoveis(),tags=('linhaimpar','106'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBInformaticaEHardware(),tags=('linhapar','107')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBFerramentas(),tags=('linhaimpar','109'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaAjudaCombustivel(),tags=('linhapar','110'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDespRepresentacao(),tags=('linhaimpar','111'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPropaganda(),tags=('linhapar','112'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBrindes(),tags=('linhaimpar','113'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBInformaticaSoft(),tags=('linhapar','120'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIMaquinasEquips(),tags=('linhaimpar','134')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIComputadoresHardware(),tags=('linhapar','135'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIOutrosInvestimentos(),tags=('linhaimpar','140')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPremioseGratificacoes(),tags=('linhapar','141'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDevolucaoVenda(),tags=('linhaimpar','144')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIrrf(),tags=('linhapar','148'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIof(),tags=('linhaimpar','149'))                                                                                             
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaJurosSEmprestimos(),tags=('linhapar','151')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaLicenciamentoseCertidoes(),tags=('linhapar','163'))                       
                     
                                                                                                                                                                                                                                                       

                     testeLabel['value'] = 75 
                     def doubleClick(event):

                        widget = event.widget
                        tree = widget.identify('item', event.x, event.y)
                        desp = widget.item(tree, 'tags')[1]

                        resultados = analiticoService.AnaliticoCompetenciaDetalhado(periodoCompete.year, periodoCompete.month, competenciaCompete.month, desp)
                        listaDeSubLinhas = widget.get_children(tree)
                        for resultado in  range(0,len(resultados),1) :
                           widget.insert(parent=tree, index=resultado,  values=resultados[resultado],tags=('linhatitulo'))   
                        for linha in listaDeSubLinhas:
                           widget.delete(linha)    
                        return

                     Analiticotree.bind("<Double-1>", doubleClick)                    
                                                                                                                                                                     

                     AnaliticoCompetencia.update_idletasks()
                     time.sleep(1)                     
                     testeLabel['value'] = 100   
                     testeLabel.destroy()  
                   
      
               threading.Thread(target=gerarTabelaAnalitico).start()
                          
               def gerarexcel():
               #leitura=Leitura()
                if(os.path.exists('temporario3.csv')):
                     os.remove('temporario3.csv')

                colunas = quantMes
                path= 'temporario3.csv'
                excel_name= tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Excel','*.xlsx')],filetypes=[('Excel','*.xlsx')])#Relatorio_Ezipa_Competencia_17-08-2022
                if not excel_name or excel_name == '/': # If the user closes the dialog without choosing location
                   messagebox.showerror('Error','Choose a location to save')
                   return 
               #leitura.excel+ str( date.today()) +'Relatorio AnaliticoCompetencia.xlsx' #Relatorio_Ezipa_Competencia_17-08-2022
                lista=[]
                with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                   csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                  #csvwriter.writerows(int('temporario3'))          
                   for row_id in Analiticotree.get_children():
                      row = Analiticotree.item(row_id,'values')
                      lista.append(row)
                   #for i, row in enumerate('values'):
                     # if i > 0:
                       # row[1] = int(row[1])
                   #lista.append(row) 
                   lista = list(map(list,lista)) #faz um mapa das linhas
                   lista.insert(0,colunas)  #insere as colunas
                   for row in lista:
                     csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário
                     #csvwriter.writerow(map(lambda x: [x], row))

                writer = pd.ExcelWriter(excel_name) #cria um novo arquivo excel 
                df = pd.read_csv(path, encoding= "ISO-8859-1") #define o encoding (nota: estavamos com um problema de utf-8, isso pode ser um problema no futuro...)
                #pd.to_numeric(df, errors = 'coerce')
                #df = df.astype('int')
                df.to_excel(writer,'Analitico',index=False) #escreve as informações no excel com o nome interno da planilha pre-definido.
                #df[row_id] = df[row].astype(int)
                writer.save() #efetivamente cria o arquivo excel
               excelpequeno_btn= PhotoImage(file="Resource\Financeiro\gerarExcelpequeno.png")
               excelpequeno_btn = excelpequeno_btn.subsample(2,2)
               figura = Label(image=excelpequeno_btn)
               figura.image=excelpequeno_btn
               excelbtn=Button(AnaliticoCompetencia,image=excelpequeno_btn, text="Excel",bd=0, highlightthickness=0, bg="#8c8c8c", compound = TOP, relief="solid", command=gerarexcel)
               excelbtn.place(x=1, y=40)
               def sairDoAnalitico():
                  AnaliticoCompetencia.destroy() 
                  return Competencia
               btn_Sair=PhotoImage(file="Resource\Financeiro\SaidaPequeno.png")
               btn_Sair = btn_Sair.subsample(2,2)
               figura = Label(image=btn_Sair)
               figura.image=btn_Sair  
               btn_Sair=Button(AnaliticoCompetencia, image= btn_Sair, text='Sair', bd=0, highlightthickness=0, bg="#8c8c8c", compound = TOP, command=sairDoAnalitico)
               btn_Sair.place(x=81, y=40 )             
             else:
                return Competencia
           btn_Analitico=PhotoImage(file="Resource\Financeiro\Analitico.png")
           btn_Analitico = btn_Analitico.subsample(2,2)
           figura = Label(image=btn_Analitico)
           figura.image=btn_Analitico  
           btn_Analitico=Button(Competencia, image= btn_Analitico, text='Analitico Adm.', bd=0, highlightthickness=0, bg="grey", compound = TOP, relief="solid", command=gerarAnaliticoAdministrativo)
           btn_Analitico.place(x=231, y=14 )                 

              
           def sairDaCompetencia():
               Competencia.destroy() 
               return ResultadoFinanceiro()
           print_btn= PhotoImage(file="Resource\Financeiro\Saida.png")
           print_btn= print_btn.subsample(2,2)
           figura = Label(image=print_btn)
           figura.image=print_btn
           print_btn=Button(Competencia, image=print_btn, text="Sair", bd=0, highlightthickness=0, bg="grey", compound = TOP, command=sairDaCompetencia)
           print_btn.place(x=321,y=14)
                         
        ok_btn=ttk.Button(Fin, text="Por competência", command=relatorioCompetencia)
        ok_btn.place(x=1, y=80)                

        def relatorioPeriodo(): 
            vendaService.ApagarValorCusto()

            Relatorio = tk.Toplevel()
            Relatorio.title("Resultado financeiro por periodo")
            Relatorio.configure(background="white")
            Relatorio.attributes("-alpha", 1.1)
            Relatorioframe = Frame(Relatorio, width=99999, height=135, bg="Grey", relief="raise")
            Relatorioframe.pack(side=TOP)
            Fin.withdraw()

            # get screen size
            screen_width = Relatorio.winfo_screenwidth()
            screen_height = Relatorio.winfo_screenheight()

            # set window size and position based on screen size
            width_ratio = 0.8  # adjust as desired
            height_ratio = 0.7  # adjust as desired
            x_pos = int((screen_width - (screen_width * width_ratio)) / 2)
            y_pos = int((screen_height - (screen_height * height_ratio)) / 2)
            width = int(screen_width * width_ratio)
            height = int(screen_height * height_ratio)

            # set window size and position
            Relatorio.geometry(f"{width}x{height}+{x_pos}+{y_pos}")
            Relatorio.resizable(width=True, height=True)
            
            periodo = tk.Label(Relatorio, text="Período de:", font=("Calibri",8), bg="grey", fg="white")
            periodo.place(x=1 ,y=104) 
            sv = StringVar()
            periodoEntry=DateEntry (Relatorio,  selectmode='day', pattern='dd/mm/yyyy' ,locale='pt_br', textvariable=sv)
            sv.trace("w", lambda name, index, mode, sv=sv: entryUpdateEndHour(periodoEntry))
            def entryUpdateEndHour(Entry):
              text = periodoEntry.get()
              if len(text) in (2,5):
                Entry.insert(END,'/')
                Entry.icursor(len(text)+1)
              if len(text) > 10:
                  Entry.delete(0,END)
                  Entry.insert(0,text[:10])
            periodoEntry.place(x=81, y=104)          

            ateLabel = tk.Label(Relatorio, text="Até:", font=("Calibri",8), bg="grey", fg="white")
            ateLabel.place(x= 181 ,y=104) 
            sv2 = StringVar()
            ateEntry= DateEntry(Relatorio, selectmode='day', pattern='dd/mm/yyyy',text=" /  /    ", locale='pt_br',textvariable=sv2)
            sv2.trace("w", lambda name, index, mode, sv2=sv2: entryUpdateEndHour2(ateEntry))            
            def entryUpdateEndHour2(Entry):
              text = ateEntry.get()
              if len(text) in (2,5):
                Entry.insert(END,'/')
                Entry.icursor(len(text)+1)
              if len(text) > 10:
                  Entry.delete(0,END)
                  Entry.insert(0,text[:10])            
            ateEntry.place(x=261,y=104)
            ateEntry.bind("<KeyPress>", lambda e: gerarRelatorioPorPeriodo() if e.char == '\r' else None)

            tree=ttk.Treeview(Relatorio, selectmode="browse", height=21, show='headings')
            tree.tag_configure('linhaimpar', background="white")
            tree.tag_configure('linhapar', background="lightblue")
            tree.tag_configure('linhatitulo', background="lightgrey")
            treescrollbar = ttk.Scrollbar(Relatorio, orient="vertical", command=tree.yview)
            treescrollbar.pack(expand = 0, side='right', fill='y')
            tree.configure(yscrollcommand=treescrollbar.set) 
            treescrollbar = ttk.Scrollbar(Relatorio, orient="horizontal", command=tree.xview)
            treescrollbar.pack(expand = 0, side='bottom', fill='x')
            tree.configure(xscrollcommand=treescrollbar.set) 
            quantMes=[]
            

            def gerarRelatorioPorPeriodo():
                 if periodoEntry.get_date() > ateEntry.get_date():
                     messagebox.showerror('ERRO', 'Você não pode fazer essa operação')   
                     return Relatorio
               
                 if periodoEntry.get_date().year != ateEntry.get_date().year:
                      messagebox.showerror('ERRO', 'Você não pode fazer essa operação')
                      return Relatorio
                 if periodoEntry.get_date().month > ateEntry.get_date().month:
                      messagebox.showerror('ERRO', "O mês selecionado está  Incorreto")
                      return Relatorio
    
                
                 def gerarTabelPorPeriodo():
                     quantMes.clear()
                     
                     testeLabel = Progressbar(Relatorio, orient = 'horizontal', length = 250, mode = 'determinate')
                     testeLabel.place(x=600,y=500)    
                     testeLabel['value'] = 25
                     Relatorio.update_idletasks()
                     time.sleep(1)
                        
                    
                     resultSet = vendaService.talvezGerarRelatorioGeralAnual(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month, periodoEntry.get_date().day,ateEntry.get_date().day )
                                     
                     tamanho = int(len(resultSet))
                        
                     tamanhoColunas = tamanho + 4 
                     tree['columns'] =list(range(1, tamanhoColunas))
                     tree.column("1", width=140, minwidth=60,stretch=NO)
                     tree.heading("#1", text='Índices')     
                     quantMes.append('Indices') 
                     
                     #GeraOsMeses
                     posicao=0
                     for j in range(0, tamanho, 1):
                        posicao=j+2
                        result = resultSet[j]
                        tree.column(str(posicao), width=80, minwidth=50,stretch=NO)
                        tree.heading("#"+str(posicao), text=result.getNomeMes())
                        quantMes.append(result.getNomeMes())
                     
                     posicao+=1   
                     tree.column(str(posicao), width=90, minwidth=50,stretch=NO)
                     tree.heading("#"+str(posicao), text='Média')
                     quantMes.append('Média')                 
                     
                     posicao+=1 
                     tree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     tree.heading("#"+str(posicao), text='Total')
                     quantMes.append('Total')   
                                       
                     testeLabel['value'] = 50
                     Relatorio.update_idletasks()
                     time.sleep(1)
                       
                     for i in tree.get_children():
                          tree.delete(i)           
                              
                     
                     tree.place(x=1,y=135)
                     tree.identify_region(10,100)
               


                     tree.insert('', 'end',  values=vendaService.getLinhaVendas(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendasCNVendas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendasCNServicos(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendas_S_N(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaTotalDeVendas(),tags=('linhatitulo'))                 
                     tree.insert('', 'end',  values=vendaService.getLinhaSimples(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaTotalLiq(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaVendas_Devolvidas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVendas_Canceladas(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaCortesias(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDescErroOperacional(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDescErroComercial(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaComissoesInternas(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaComissoesExternas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaTotalVendasGeral(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values='',tags=('linhaimpar'))
                     tree.insert('', 'end',  values=('Despesas',),tags=('linhatitulo')) #CONSERTAR AQUI
                     tree.insert('', 'end',  values=vendaService.getLinhaProLaboreGerencia(),tags=('linhaimpar'))  
                     tree.insert('', 'end',  values=('ESTOQUE',),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaLaminas(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMadeiras(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVazadores(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaBorrachas(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaPertinax(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMaterialPDestaque(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMaquinasParaRevenda(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaOutros(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaServicosExternos(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))                     
                     tree.insert('', 'end',  values=('CLIENTES',),tags=('linhatitulo'))                     
                     tree.insert('', 'end',  values=vendaService.getLinhaLaminasClientes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMadeirasClientes(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaVazadoresClientes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaBorrachasClientes(),tags=('linhaipar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaPertinaxClientes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMaterialPDestaqueClientes(),tags=('linhapar'))                     
                     tree.insert('', 'end',  values=vendaService.getLinhaCliches(),tags=('linhaimpar'))                        
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDesenho(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosDesenho(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHE(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestes(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscates(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadas(),tags=('linhapar'))                     
                     tree.insert('', 'end',  values=vendaService.getLinhausoeConsumoDesenho(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhamanutencao(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaPlotter(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaUsoeConsumoPlotter(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManutencaoPlotter(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaLaser(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosLaser(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHELaser(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestesLaser(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscatesLaser(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadasLaser(),tags=('linhaimpar'))                      
                     tree.insert('', 'end',  values=vendaService.getLinhaUsoeConsumoLaser(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManutencaoLaser(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaGaspLaser(),tags=('linhapar'))  
                     tree.insert('', 'end',  values='',tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaMontagem(),tags=('linhatitulo'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosMontagem(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHEMontagem(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestesMontagem(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscatesMontagem(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadasMontagem(),tags=('linhaimpar'))                        
                     tree.insert('', 'end',  values=vendaService.getLinhausoEconsumoMontagem(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManuTencaoMontagem(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaBorracha(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosBorracha(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaHEBorracha(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhatestesBorracha(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhabiscatesBorracha(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaFeriasTrabalhadasBorracha(),tags=('linhaimpar'))                           
                     tree.insert('', 'end',  values=vendaService.getLinhaUsoeConsumoBorracha(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaManutencaoBorracha(),tags=('linhaimpar'))   
                     tree.insert('', 'end',  values='',tags=('linhapar')) 
                     tree.insert('', 'end',  values=vendaService.getLinhatotalGeralServiços(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values='',tags=('linhapar')) 
                     tree.insert('', 'end',  values=vendaService.getLinhatotaldeResultados1(),tags=('linhatitulo')) 
                     tree.insert('', 'end',  values='',tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaDepreciacao(),tags=('linhatitulo')) 
                     tree.insert('', 'end',  values='',tags=('linhapar')) 
                     tree.insert('', 'end',  values=vendaService.getLinhatotaldeResultados2(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaDespesasAdministrativas(),tags=('linhaimpar'))                      
                     tree.insert('', 'end',  values=vendaService.getLinhaDespesasCentrais(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosAdm(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosComercial(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaSalariosTransporte(),tags=('linhaimpar'))   
                     tree.insert('', 'end',  values='',tags=('linhapar'))                      
                     tree.insert('', 'end',  values=vendaService.getLinhatotalGeralAdministrativo(),tags=('linhatitulo'))    
                     tree.insert('', 'end',  values=vendaService.getLinhaDespesasFinanceiras(),tags=('linhaimpar'))  
                     tree.insert('', 'end',  values=vendaService.getLinhaJuros(),tags=('linhapar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaJurosSobreAplicacoes(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getLinhaJurosSobreVendas(),tags=('linhapar'))
                     tree.insert('', 'end',  values='',tags=('linhaimpar'))                                                    
                     tree.insert('', 'end',  values=vendaService.getLinhatotalResultadoEzipa(),tags=('linhatitulo'))  
                     tree.insert('', 'end',  values=vendaService.getlinhatotalWoodpaper(),tags=('linhaimpar'))
                     tree.insert('', 'end',  values=vendaService.getlinhatotalEzipaeWoodpaper(),tags=('linhatitulo'))
                                       
                     
                     testeLabel['value'] = 75
                     Relatorio.update_idletasks()
                     time.sleep(1)                     
                     testeLabel['value'] = 100   
                     messagebox.showinfo("RELATÓRIO", "Relatório gerado com sucesso")                       
                     testeLabel.destroy()  
                     return
                 threading.Thread(target=gerarTabelPorPeriodo).start()

                        
            ok_btn=ttk.Button(Relatorio, text="ok", command=gerarRelatorioPorPeriodo)
            ok_btn.place(x=501, y=104)
                              
            def relatorioExcel():
             #leitura=Leitura()
             if(os.path.exists('temporario2.csv')):
                   os.remove('temporario2.csv')

             colunas = quantMes
             path= 'temporario2.csv'
             excel_name= tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Excel','*.xlsx')],filetypes=[('Excel','*.xlsx')])#Relatorio_Ezipa_Competencia_17-08-2022
             if not excel_name or excel_name == '/': # If the user closes the dialog without choosing location
               messagebox.showerror('Error','Choose a location to save')
               return 
             lista=[]
             with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                     csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                     for row_id in tree.get_children():
                        row = tree.item(row_id,'values')
                        lista.append(row)
                     lista = list(map(list,lista)) #faz um mapa das linhas
                     lista.insert(0,colunas)  #insere as colunas
                     for row in lista:
                        csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário
             writer = pd.ExcelWriter(excel_name) #cria um novo arquivo excel 
             df = pd.read_csv(path, encoding='latin') #define o encoding (nota: estavamos com um problema de utf-8, isso pode ser um problema no futuro...)
             df.to_excel(writer,'Analitico',index=False) #escreve as informações no excel com o nome interno da planilha pre-definido
             writer.save() #efetivamente cria o arquivo excel

            excel_btn= PhotoImage(file="Resource\Financeiro\gerarExcel.png")
            excel_btn = excel_btn.subsample(2,2)
            figura = Label(image=excel_btn)
            figura.image=excel_btn
            excelbtn=Button(Relatorio,image=excel_btn, text="Excel",bd=0, highlightthickness=0, bg="grey", compound = TOP, command=relatorioExcel)
            excelbtn.place(x=1, y=14)

            def relatorioPdf():
             #leitura=Leitura()
             if(os.path.exists('temporario2.csv')):
               os.remove('temporario2.csv')
             Caminhopdf= tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Arquivo Pdf','*.pdf')],filetypes=[('Arquivo pdf','*.pdf')])#Relatorio_Ezipa_Competencia_17-08-2022
             if not Caminhopdf or Caminhopdf == '/': # If the user closes the dialog without choosing location
               messagebox.showerror('Error','Choose a location to save')
               return # Stop the function   
             #leitura.PDF+ str( date.today())+'Relatorio Financeiro Periodo.pdf'
   
              
             colunas = quantMes             

             path= 'temporario2.csv'
             lista=[]
             with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
               csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
               for row_id in tree.get_children():
                  row = tree.item(row_id,'values')
                  lista.append(row)    
               lista = list(map(list,lista)) #faz um mapa das linhas
               lista.insert(0,colunas)  #insere as colunas
               for row in lista:
                  csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário
  
 
             a = pd.read_csv(path,  encoding='"latin-1"') 
             a.to_html('Relatório por Periodo.html') 
             html_file = a.to_html() 
             path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'#exe  necessário para rodar o pdf
             config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)#cria a configuração  do  pdf
             options={'encoding': "UTF-8"  }
             pdfkit.from_file('relatório por Periodo.html', Caminhopdf, configuration=config, options=options)#cria o arquivo em pdf
   
             
            pdf_btn= PhotoImage(file="Resource\Financeiro\gerarpdf.png")
            pdf_btn = pdf_btn.subsample(2,2)
            figura = Label(image=pdf_btn)
            figura.image=pdf_btn
            pdf_btn=Button(Relatorio,image=pdf_btn,  text="Pdf",bd=0, highlightthickness=0, bg="grey", compound = TOP,command=relatorioPdf)
            pdf_btn.place(x=81, y=14)  
            
            def imprimieRelatorio():
             answer= messagebox.askyesno("Impressão","Deseja Imprimir") 
             if answer == YES:  
             #Leitura=Leiturapdf()
               if(os.path.exists('temporario2.csv')):
                  os.remove('temporario2.csv')
               # Stop the function    
               Caminhopdf2='Relatorio Financeiro .pdf' 
               #str( date.today())+'Relatorio Financeiro .pdf'
               
               colunas = quantMes  
               path= 'temporario2.csv'
               lista=[]
               with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                  csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                  for row_id in tree.get_children():
                     row = tree.item(row_id,'values')
                     lista.append(row)
                  lista = list(map(list,lista)) #faz um mapa das linhas
                  lista.insert(0,colunas)  #insere as colunas
                  for row in lista:
                     csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário   
               a = pd.read_csv(path,  encoding="latin-1") #le o csv
               a.to_html('Impressao.html') #cria o html que  gerará o pdf
               html_file = a.to_html() 
               path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'#necessita para a criação do pdf
               config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
               options={'encoding': "UTF-8" ,  'cookie': [('cookie-empty-value', '""')]}#cria  opções para o pdf
               pdfkit.from_file('impressao.html', Caminhopdf2, configuration=config, options=options)#gera o caminho do  pdf               
               

               filename = Caminhopdf2
               win32api.ShellExecute (0,"print",filename,'/d:"%s"' % win32print.GetDefaultPrinter (),".",0)        
             else:
                return Relatorio   
            print_btn= PhotoImage(file="Resource\Financeiro\gerarimpressão.png")
            print_btn= print_btn.subsample(2,2)
            figura = Label(image=print_btn)
            figura.image=print_btn
            print_btn=Button(Relatorio, image=print_btn, text="Imprimir", bd=0, highlightthickness=0, bg="grey", compound = TOP, command=imprimieRelatorio)
            print_btn.place(x=161,y=16) 
            
            def gerarAnaliticoAdministrativo(): 
             answer= messagebox.askyesno("Analítico Admiministrativo","Deseja Gerar Analítico Administrativo") 
             if answer == YES:

               Analitico = tk.Toplevel()
               titulo = f"Analitico ADM de {datetime.strftime(periodoEntry.get_date(),'%d/%m/%Y')} até {datetime.strftime(ateEntry.get_date(),'%d/%m/%Y')}"
               Analitico.title(titulo)

               Analitico.geometry("950x750+350+10")
               Analitico.configure(background="white")
               Analitico.resizable(width=True, height=True)
               Analitico.attributes("-alpha", 1.1)
               Analiticoframe = Frame(Analitico, width=99999, height=135, bg="#8c8c8c", relief="raise")
               Analiticoframe.pack(side=TOP)
               Analitico.grid
               screen_width = Analitico.winfo_screenwidth()
               screen_height = Analitico.winfo_screenheight()               
               
               width_ratio = 0.8  # adjust as desired
               height_ratio = 0.7  # adjust as desired
               x_pos = int((screen_width - (screen_width * width_ratio)) / 2)
               y_pos = int((screen_height - (screen_height * height_ratio)) / 2)
               width = int(screen_width * width_ratio)
               height = int(screen_height * height_ratio)

                  # set window size and position
               Analitico.geometry(f"{width}x{height}+{x_pos}+{y_pos}")
               Analitico.resizable(width=True, height=True)               
                                                            
               Analiticotree=ttk.Treeview(Analitico, selectmode="browse",  height=20, show='headings')
               Analiticotree.tag_configure('linhaimpar', background="white")
               Analiticotree.tag_configure('linhapar', background="lightblue")
               Analiticotree.tag_configure('linhatitulo', background="lightyellow")
               treescrollbar = ttk.Scrollbar(Analitico, orient="vertical", command=Analiticotree.yview)
               treescrollbar.pack(expand = 0, side='right', fill='y')
               Analiticotree.configure(yscrollcommand=treescrollbar.set) 
               treescrollbar = ttk.Scrollbar(Analitico, orient="horizontal", command=Analiticotree.xview)
               treescrollbar.pack(expand = 0, side='bottom', fill='x')
               Analiticotree.configure(xscrollcommand=treescrollbar.set) 
               quantMes=[]
                  
               def gerarTabelaAnalitico():
                     quantMes.clear()
                     analiticoService.Apagaranaliticoadm()
                     analiticoService.ApagarValorAnalitico()                     
                     
                     testeLabel = Progressbar(Analitico,  orient = 'horizontal', length = 250, mode = 'determinate')
                     testeLabel.place(x=600,y=500)    
                     testeLabel['value'] = 25
                     Analitico.update_idletasks()
                     time.sleep(1)
                     
                     resultado = analiticoService.GerarRelatorioDetalhadoAnalitico(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month)
                     if periodoEntry.get_date() > ateEntry.get_date():
                        messagebox.showerror('ERRO', 'Você não pode fazer essa operação')   
                        return Relatorio
                  
                     if periodoEntry.get_date().year != ateEntry.get_date().year:
                        messagebox.showerror('ERRO', 'Você não pode fazer essa operação')
                        return Relatorio
                     if periodoEntry.get_date().month > ateEntry.get_date().month:
                        messagebox.showerror('ERRO', "O mês selecionado está  Incorreto")
                        return Relatorio                  
                     

                           
                           
                     resultSet = analiticoService.GerarRelatorioAnalitico(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month, periodoEntry.get_date().day,ateEntry.get_date().day)                                     
                     tamanho = int(len(resultSet))
                           

                     tamanhoColunas = tamanho + 4
                     Analiticotree['columns'] =list(range(0, tamanhoColunas))
                     
                        
                        #GeraOsMeses
                  
                     
                        
                     Analiticotree.column("1", width=120, minwidth=100,stretch=NO)
                     Analiticotree.heading("#1", text='Nome das Despesas')     
                     quantMes.append('Nome das Despesas') 
                        
                        #GeraOsMeses
                     posicao=1
                     
                     testeLabel['value'] = 50
                     Analitico.update_idletasks()
                     time.sleep(1)
                     
                     for j in range(0, tamanho, 1):
                        posicao=j+2
                        result = resultSet[j]
                        Analiticotree.column(str(posicao), width=80, minwidth=50,stretch=NO)
                        Analiticotree.heading("#"+str(posicao), text=result.getNomeMes())
                        quantMes.append(result.getNomeMes())
                     
                     posicao+=1   
                     Analiticotree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     Analiticotree.heading("#"+str(posicao), text='Média')
                        
                  
                     
                     quantMes.append('Média')                 
                     posicao+=1 
                     Analiticotree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     Analiticotree.heading("#"+str(posicao), text='Total')

                     quantMes.append('Total')   
                                          
                     
                        
                     for i in Analiticotree.get_children():
                        Analiticotree.delete(i)           
                                 
                        
                     Analiticotree.place(x=1,y=135)
                     Analiticotree.identify_region(10,100)
                  
                     Analiticotree.insert('', 'end', values=analiticoService.getLinhaAluguel(),tags=('linhaimpar','3'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIptu(),tags=('linhapar','4'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCedae(),tags=('linhimpar','5'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaLight(),tags=('linhapar','6'))                     
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCeg(),tags=('linhaimpar','7'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTelefonia(),tags=('linhapar','8')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPnP(),tags=('linhaimpar','9')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaContabilidade(),tags=('linhapar','11'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSigraf(),tags=('linhaimpar','12')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaInformatica(),tags=('linhapar','14'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMaterialdeEscritorio(),tags=('linhaimpar','15')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMaterialLimpeza(),tags=('linhapar','16')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCartuchosParaImpressoras(),tags=('linhaimpar','17')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOnibusContinuo(),tags=('linhapar','18')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCorreios(),tags=('linhaimpar','19')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDespesasBancarias(),tags=('linhapar','20')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutrosAdministrativos(),tags=('linhaimpar','21')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaKombi(),tags=('linhapar','23')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMotoboy(),tags=('linhaimpar','24')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutrosFretes(),tags=('linhapar','25')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTranpostadoras(),tags=('linhaimpar','26')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSedexRemessaVenda(),tags=('linhapar','27')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPedagioFreteVenda(),tags=('linhaimpar','30')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSimples(),tags=('linhapar','31'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutrosImpostos(),tags=('linhaimpar','34')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaSalarios(),tags=('linhapar','35'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaRecisoes(),tags=('linhaimpar','36')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaValeTransporte(),tags=('linhapar','37'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaValeAlimentacao(),tags=('linhaimpar','38')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaInss(),tags=('linhapar','39')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFgts(),tags=('linhaimpar','40')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaAjudaDeCusto(),tags=('linhapar','41')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFerias(),tags=('linhaimpar','42')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaHoraExtra(),tags=('linhapar','43')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTestes(),tags=('linhaimpar','44')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCestaBasica(),tags=('linhapar','45')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaLanches(),tags=('linhaimpar','46'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhapesquisaDesenvolvimento(),tags=('linhapar','68'))                      
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFarmacia(),tags=('linhaimpar','47'))     
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaTreinamento(),tags=('linhapar','48')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBiscate(),tags=('linhaimpar','49'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaOutros(),tags=('linhapar','50')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaEmprestimosInternos(),tags=('linhaimpar','71'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaResultados(),tags=('linhapar','72')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaProLaboreDiretoria(),tags=('linhaimpar','73')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaProLaboreGerencia(),tags=('linhapar','74'))     
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDespesasCadastraisClientes(),tags=('linhaimpar','75')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaExames(),tags=('linhapar','77')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaCipa(),tags=('linhaimpar','78'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaServSeguranca(),tags=('linhapar','79'))
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMinformatica(),tags=('linhaimpar','86'))                      
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPlanoDeSaude(),tags=('linhapar','91'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFeriasTrabalhadas(),tags=('linhaimpar','92'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaEpiUniformes(),tags=('linhapar','93'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaFestasEConfraternizações(),tags=('linhaimpar','95'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMPredial(),tags=('linhapar','97'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMEletrica(),tags=('linhaimpar','98'))                    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMArCondicionado(),tags=('linhapar','99'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMBebedouro(),tags=('linhaimpar','100'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMMoveis(),tags=('linhapar','101'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMOutrosManut(),tags=('linhaimpar','102'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaMHidraulica(),tags=('linhapar','103'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBMaquinaseEquip(),tags=('linhaimpar','104'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaObrasEMelhorias(),tags=('linhapar','105'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBMoveis(),tags=('linhaimpar','106'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBInformaticaEHardware(),tags=('linhapar','107')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBFerramentas(),tags=('linhaimpar','109'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaAjudaCombustivel(),tags=('linhapar','110'))   
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDespRepresentacao(),tags=('linhaimpar','111'))    
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPropaganda(),tags=('linhapar','112'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBrindes(),tags=('linhaimpar','113'))  
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaBInformaticaSoft(),tags=('linhapar','120'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIMaquinasEquips(),tags=('linhaimpar','134')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIComputadoresHardware(),tags=('linhapar','135'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIOutrosInvestimentos(),tags=('linhaimpar','140')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaPremioseGratificacoes(),tags=('linhapar','141'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaDevolucaoVenda(),tags=('linhaimpar','144')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIrrf(),tags=('linhapar','148'))                                                                                                                 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaIof(),tags=('linhaimpar','149'))                                                                                             
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaJurosSEmprestimos(),tags=('linhapar','151')) 
                     Analiticotree.insert('', 'end',  values=analiticoService.getLinhaLicenciamentoseCertidoes(),tags=('linhapar','163'))                       
                     
                                                                                                                                                                                                                                                       

                     testeLabel['value'] = 75 
                     def doubleClick(event):

                        widget = event.widget
                        tree = widget.identify('item', event.x, event.y)
                        desp = widget.item(tree, 'tags')[1]

                        resultados = analiticoService.AnaliticoDetalhado(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month, desp)
                        listaDeSubLinhas = widget.get_children(tree)
                        for resultado in  range(0,len(resultados),1) :
                           widget.insert(parent=tree, index=resultado,  values=resultados[resultado],tags=('linhatitulo'))   
                        for linha in listaDeSubLinhas:
                           widget.delete(linha)    
                        return

                     Analiticotree.bind("<Double-1>", doubleClick)                    
                                                                                                                                                                     

                     Analitico.update_idletasks()
                     time.sleep(1)                     
                     testeLabel['value'] = 100   
                     testeLabel.destroy()  
                   
      
               threading.Thread(target=gerarTabelaAnalitico).start()
                          
               def gerarexcel():
               #leitura=Leitura()
                if(os.path.exists('temporario3.csv')):
                     os.remove('temporario3.csv')

                colunas = quantMes
                path= 'temporario3.csv'
                excel_name= tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Excel','*.xlsx')],filetypes=[('Excel','*.xlsx')])#Relatorio_Ezipa_Competencia_17-08-2022
                if not excel_name or excel_name == '/': # If the user closes the dialog without choosing location
                   messagebox.showerror('Error','Choose a location to save')
                   return 
               #leitura.excel+ str( date.today()) +'Relatorio Analitico.xlsx' #Relatorio_Ezipa_Competencia_17-08-2022
                lista=[]
                with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                   csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                  #csvwriter.writerows(int('temporario3'))          
                   for row_id in Analiticotree.get_children():
                      row = Analiticotree.item(row_id,'values')
                      lista.append(row)
                   #for i, row in enumerate('values'):
                     # if i > 0:
                       # row[1] = int(row[1])
                   #lista.append(row) 
                   lista = list(map(list,lista)) #faz um mapa das linhas
                   lista.insert(0,colunas)  #insere as colunas
                   for row in lista:
                     csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário
                     #csvwriter.writerow(map(lambda x: [x], row))

                writer = pd.ExcelWriter(excel_name) #cria um novo arquivo excel 
                df = pd.read_csv(path, encoding= "ISO-8859-1") #define o encoding (nota: estavamos com um problema de utf-8, isso pode ser um problema no futuro...)
                #pd.to_numeric(df, errors = 'coerce')
                #df = df.astype('int')
                df.to_excel(writer,'Analitico',index=False) #escreve as informações no excel com o nome interno da planilha pre-definido.
                #df[row_id] = df[row].astype(int)
                writer.save() #efetivamente cria o arquivo excel
               excelpequeno_btn= PhotoImage(file="Resource\Financeiro\gerarExcelpequeno.png")
               excelpequeno_btn = excelpequeno_btn.subsample(2,2)
               figura = Label(image=excelpequeno_btn)
               figura.image=excelpequeno_btn
               excelbtn=Button(Analitico,image=excelpequeno_btn, text="Excel",bd=0, highlightthickness=0, bg="#8c8c8c", compound = TOP, relief="solid", command=gerarexcel)
               excelbtn.place(x=1, y=40)
               def sairDoAnalitico():
                  Analitico.destroy() 
                  return Relatorio
               btn_Sair=PhotoImage(file="Resource\Financeiro\SaidaPequeno.png")
               btn_Sair = btn_Sair.subsample(2,2)
               figura = Label(image=btn_Sair)
               figura.image=btn_Sair  
               btn_Sair=Button(Analitico, image= btn_Sair, text='Sair', bd=0, highlightthickness=0, bg="#8c8c8c", compound = TOP, command=sairDoAnalitico)
               btn_Sair.place(x=81, y=40 )             
             else:
                return Relatorio
            btn_Analitico=PhotoImage(file="Resource\Financeiro\Analitico.png")
            btn_Analitico = btn_Analitico.subsample(2,2)
            figura = Label(image=btn_Analitico)
            figura.image=btn_Analitico  
            btn_Analitico=Button(Relatorio, image= btn_Analitico, text='Analitico Adm.', bd=0, highlightthickness=0, bg="grey", compound = TOP, relief="solid", command=gerarAnaliticoAdministrativo)
            btn_Analitico.place(x=231, y=14 )                  


            def gerarAnaliticoCentral(): 
             answer= messagebox.askyesno("Analítico Central","Deseja Gerar Analítico Central") 
             if answer == YES:

               AnaliticoCentral = tk.Toplevel()
               AnaliticoCentral.title("Analitico Central")
               AnaliticoCentral.geometry("950x750+350+10")
               AnaliticoCentral.configure(background="white")
               AnaliticoCentral.resizable(width=True, height=True)
               AnaliticoCentral.attributes("-alpha", 1.1)
               AnaliticoCentralframe = Frame(AnaliticoCentral, width=99999, height=135, bg="#8c8c8c", relief="raise")
               AnaliticoCentralframe.pack(side=TOP)
               AnaliticoCentral.grid
                                                            
               AnaliticoCentraltree=ttk.Treeview(AnaliticoCentral, selectmode="browse",  height=20, show='headings')
               AnaliticoCentraltree.tag_configure('linhaimpar', background="white")
               AnaliticoCentraltree.tag_configure('linhapar', background="lightblue")
               AnaliticoCentraltree.tag_configure('linhatitulo', background="lightyellow")
               treescrollbar = ttk.Scrollbar(AnaliticoCentral, orient="vertical", command=AnaliticoCentraltree.yview)
               treescrollbar.pack(expand = 0, side='right', fill='y')
               AnaliticoCentraltree.configure(yscrollcommand=treescrollbar.set) 
               treescrollbar = ttk.Scrollbar(AnaliticoCentral, orient="horizontal", command=AnaliticoCentraltree.xview)
               treescrollbar.pack(expand = 0, side='bottom', fill='x')
               AnaliticoCentraltree.configure(xscrollcommand=treescrollbar.set) 
               quantMes=[]
               
               def gerarTabelaAnaliticoCentral():
                     quantMes.clear()
                     analiticoService.Apagaranaliticoadm()
                     testeLabel = Progressbar(AnaliticoCentral,  orient = 'horizontal', length = 250, mode = 'determinate')
                     testeLabel.place(x=400,y=500)    
                     testeLabel['value'] = 25
                     AnaliticoCentral.update_idletasks()
                     time.sleep(1)
                     
                     resultado = analiticoCentralService.GerarRelatorioDetalhadoAnaliticoCentral(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month)
                     if periodoEntry.get_date() > ateEntry.get_date():
                        messagebox.showerror('ERRO', 'Você não pode fazer essa operação')   
                        return Relatorio
                  
                     if periodoEntry.get_date().year != ateEntry.get_date().year:
                        messagebox.showerror('ERRO', 'Você não pode fazer essa operação')
                        return Relatorio
                     if periodoEntry.get_date().month > ateEntry.get_date().month:
                        messagebox.showerror('ERRO', "O mês selecionado está  Incorreto")
                        return Relatorio                  
                     

                           
                           
                     resultSet = analiticoService.GerarRelatorioAnalitico(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month, periodoEntry.get_date().day,ateEntry.get_date().day)                                     
                     tamanho = int(len(resultSet))
                           

                     tamanhoColunas = tamanho + 4
                     AnaliticoCentraltree['columns'] =list(range(0, tamanhoColunas))
                     
                        
                        #GeraOsMeses
                  
                     
                        
                     AnaliticoCentraltree.column("1", width=120, minwidth=100,stretch=NO)
                     AnaliticoCentraltree.heading("#1", text='Nome das Despesas')     
                     quantMes.append('Nome das Despesas') 
                        
                        #GeraOsMeses
                     posicao=1
                     
                     testeLabel['value'] = 50
                     AnaliticoCentral.update_idletasks()
                     time.sleep(1)
                     
                     for j in range(0, tamanho, 1):
                        posicao=j+2
                        result = resultSet[j]
                        AnaliticoCentraltree.column(str(posicao), width=80, minwidth=50,stretch=NO)
                        AnaliticoCentraltree.heading("#"+str(posicao), text=result.getNomeMes())
                        quantMes.append(result.getNomeMes())
                     
                     posicao+=1   
                     AnaliticoCentraltree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     AnaliticoCentraltree.heading("#"+str(posicao), text='Média')
                        
                  
                     
                     quantMes.append('Média')                 
                     posicao+=1 
                     AnaliticoCentraltree.column(str(posicao), width=100, minwidth=50,stretch=NO)
                     AnaliticoCentraltree.heading("#"+str(posicao), text='Total')

                     quantMes.append('Total')   
                                          
                     
                        
                     for i in AnaliticoCentraltree.get_children():
                        AnaliticoCentraltree.delete(i)           
                                 
                        
                     AnaliticoCentraltree.place(x=1,y=200)
                     AnaliticoCentraltree.identify_region(10,100)  
                     
                     testeLabel['value'] = 75 
                     def doubleClick(event):

                        widget = event.widget
                        tree = widget.identify('item', event.x, event.y)
                        desp = widget.item(tree, 'tags')[1]

                        resultados = analiticoService.AnaliticoDetalhado(periodoEntry.get_date().year, ateEntry.get_date().month, periodoEntry.get_date().month, desp)
                        listaDeSubLinhas = widget.get_children(tree)
                        for resultado in  range(0,len(resultados),1) :
                           widget.insert(parent=tree, index=resultado,  values=resultados[resultado],tags=('linhatitulo'))   
                        for linha in listaDeSubLinhas:
                           widget.delete(linha)    
                        return

                     AnaliticoCentraltree.bind("<Double-1>", doubleClick)                    
                                                                                                                                                                     

                     AnaliticoCentral.update_idletasks()
                     time.sleep(1)                     
                     testeLabel['value'] = 100   
                     testeLabel.destroy()  
                   
      
               threading.Thread(target=gerarTabelaAnaliticoCentral).start()                                  
               
               def analiticoCentralGerarExcel():
               #leitura=Leitura()
                if(os.path.exists('temporario3.csv')):
                     os.remove('temporario3.csv')

                colunas = quantMes
                path= 'temporario3.csv'
                excel_name= tk.filedialog.asksaveasfilename(title='Salvar como',defaultextension=[('Excel','*.xlsx')],filetypes=[('Excel','*.xlsx')])#Relatorio_Ezipa_Competencia_17-08-2022
                if not excel_name or excel_name == '/': # If the user closes the dialog without choosing location
                   messagebox.showerror('Error','Choose a location to save')
                   return 
               #leitura.excel+ str( date.today()) +'Relatorio Analitico.xlsx' #Relatorio_Ezipa_Competencia_17-08-2022
                lista=[]
                with open(path, "w", newline='') as myFile: #aqui eu to gerando um arquivo em memória
                   csvwriter = csv.writer(myFile, delimiter=',') #aqui Eu abro o arquivo em memoria para gravação
                  #csvwriter.writerows(int('temporario3'))          
                   for row_id in AnaliticoCentraltree.get_children():
                      row = AnaliticoCentraltree.item(row_id,'values')
                      lista.append(row)
                   #for i, row in enumerate('values'):
                     # if i > 0:
                       # row[1] = int(row[1])
                   #lista.append(row) 
                   lista = list(map(list,lista)) #faz um mapa das linhas
                   lista.insert(0,colunas)  #insere as colunas
                   for row in lista:
                     csvwriter.writerow(row) #escreve efetivamente as informações no arquivo temporário
                     #csvwriter.writerow(map(lambda x: [x], row))

                writer = pd.ExcelWriter(excel_name) #cria um novo arquivo excel 
                df = pd.read_csv(path, encoding= "ISO-8859-1") #define o encoding (nota: estavamos com um problema de utf-8, isso pode ser um problema no futuro...)
                #pd.to_numeric(df, errors = 'coerce')
                #df = df.astype('int')
                df.to_excel(writer,'Analitico',index=False) #escreve as informações no excel com o nome interno da planilha pre-definido.
                #df[row_id] = df[row].astype(int)
                writer.save() #efetivamente cria o arquivo excel
               excelpequeno_btn= PhotoImage(file="Resource\Financeiro\gerarExcelpequeno.png")
               excelpequeno_btn = excelpequeno_btn.subsample(2,2)
               figura = Label(image=excelpequeno_btn)
               figura.image=excelpequeno_btn
               excelbtn=Button(AnaliticoCentral,image=excelpequeno_btn, text="Excel",bd=0, highlightthickness=0, bg="#8c8c8c", compound = TOP, relief="solid", command=analiticoCentralGerarExcel)
               excelbtn.place(x=1, y=40)               
               
             else:
                return Relatorio
            btn_AnaliticoCentral=PhotoImage(file="Resource\Financeiro\Analitico.png")
            btn_AnaliticoCentral = btn_AnaliticoCentral.subsample(2,2)
            figura = Label(image=btn_AnaliticoCentral)
            figura.image=btn_AnaliticoCentral  
            btn_AnaliticoCentral=Button(Relatorio, image= btn_AnaliticoCentral, text='Analitico Central', bd=0, highlightthickness=0, bg="grey", compound = TOP, relief="solid", command=gerarAnaliticoCentral)
            btn_AnaliticoCentral.place(x=321, y=14 )                   

            def sairDoPeriodo():
                Relatorio.destroy() 
                return ResultadoFinanceiro()
            btn_Sair=PhotoImage(file="Resource\Financeiro\Saida.png")
            btn_Sair = btn_Sair.subsample(2,2)
            figura = Label(image=btn_Sair)
            figura.image=btn_Sair  
            btn_Sair=Button(Relatorio, image= btn_Sair, text='Sair', bd=0, highlightthickness=0, bg="grey", compound = TOP, command=sairDoPeriodo)
            btn_Sair.place(x=431, y=14 )  
            

        ok_btn=ttk.Button(Fin,  text="Por Periodo", command=relatorioPeriodo)
        ok_btn.place(x=1,y=110)   

  
      grafico7 = PhotoImage(file="Resource\Gest\Resultado  FInanceiro.png")
      grafico7 = grafico7.subsample(2,2)
      figura = Label(image=grafico7)
      figura.image=grafico7
      if(usuarioLogado.getPermissoes().perm_gestfin_resultfin == True):
         Result_btn=Button(Downframe,image=grafico7, text="Resultado Financeiro",bd=0, highlightthickness=0, bg="white", compound = TOP, command=ResultadoFinanceiro) 
      Result_btn.place(x=860, y=180)    
      
      def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
      Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
      Backbutton.place(x=920, y=380)       

      def Voltando():
        new.destroy()
        jan.destroy()
        return 
      Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
      Backbutton.place(x=920, y=420)
      new.mainloop()
 
  
   def button_hover(e):
       btn["bg"]="lightblue"
   def buttonhoverleave(e):
       btn["bg"]="#040D9E"   
   Financeiro7 = PhotoImage(file="Resource\Financeiro\Gestão Financeira.gif")
   Financeiro7 = Financeiro7.subsample(2,2)
   figura = Label(image=Financeiro7)
   figura.image=Financeiro7
   btn=Button(Upframe,image=Financeiro7, text="Gestão Financeira",bd=0, highlightthickness=0, bg="#040D9E",fg="white", compound = TOP, command=gestFinanceira)
   btn.bind("<Enter>", button_hover)  
   btn.bind("<Leave>",buttonhoverleave)   
   btn.place(x=871,y=120)
 
   Versaolabel= Label(Downframe, text="Versão 1.1", font=("Calibri",10), fg="black")
   Versaolabel.place(x= 10 ,y=420)    
   def AlterarLogin():
       new.destroy()
       jan.deiconify()
       return 
   Backbutton = ttk.Button(Downframe, text="Alterar Usuario", width=25, command=AlterarLogin)
   Backbutton.place(x=920, y=380)      

   def Voltando():
     new.destroy()
     jan.destroy()
     return 
   Backbutton = ttk.Button(Downframe, text="Sair", width=25, command=Voltando)
   Backbutton.place(x=920, y=420)
   new.mainloop()
  


LoginButton = ttk.Button(bottomframe, text="Entrar", width=25, command=Entrar)
LoginButton.bind("<KeyPress>", lambda e: Entrar() if e.char == '\r' else None) 
LoginButton.place(x=485,y=275)   

def Sair():
    jan.destroy()
    return
    
ExitButton = ttk.Button(bottomframe, text="Sair", width=25, command=Sair)
ExitButton.place(x=850,y=475)





jan.mainloop()





