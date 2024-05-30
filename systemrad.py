import tkinter as tk
from tkinter import ttk
import pandas as pd

class PrincipalRAD:
    def __init__(self, win):
        
        self.lblNome=tk.Label(win, text='Nome do Aluno:')
        self.lblNota1=tk.Label(win, text='Nota 1')
        self.lblNota2=tk.Label(win, text='Nota 2')
        self.lblMedia=tk.Label(win, text='Média')
        self.txtNome=tk.Entry(bd=3)
        self.txtNota1=tk.Entry()
        self.txtNota2=tk.Entry()
        self.btnCalcular=tk.Button(win, text='Calcular Média', command=self.fCalcularMedia)
        
        self.dadosColunas = ("Aluno", "Nota 1", "Nota2", "Média", "Situação")
        
        self.treeMedias = ttk.Treeview(win,
                                       columns=self.dadosColunas,
                                       selectmode='browse')
        
        self.verscrlbar = ttk.Scrollbar(win,
                                        orient="vertical",
                                        command=self.treeMedias.yview)
        
        self.verscrlbar.pack(side = 'rigth', fill = 'x')
        
        
        
        self.treeMedias.configure(yscrollcommand=self.verscrlbar.set)

        self.treeMedias.heading("Aluno", text='Aluno')
        self.treeMedias.heading("Nota 1", text='Nota 1')
        self.treeMedias.heading("Nota 2", text='Nota 2')
        self.treeMedias.heading("Média", text='Media')
        self.treeMedias.heading("Situação", text='Situção')
        
        self.treeMedias.pack(padx=10, pady=10)
        
        self.lblNome.place(x=100, y=50)
        self.lblNome.place(x=200, y=50)
        
        self.lblNota1.place(x=100, y=100)
        self.lblNota1.place(x=200, y=100)
        
        self.lblNota2.place(x=100, y=150)
        self.lblNota2.place(x=200, y=150)
        
        self.btnCalcular.place(x=100, y=200)
        
        self.treeMedias.place(x=100, y=300)
        self.verscrlbar.place(x=805, y=300, height=225)
        
        self.id = 0
        self.id = 0
        
        self.carregarDadosIniciais()
        
    def carregarDadosIniciais(self):
        try:
            fsave = 'planilhaAlunos.xlsx'
            dados = pd.read_excel(fsave)
            #fsave = 'planilhaAlunos.csv'
            #dados = pd.read_csv(fsave)
            print("********* dados disponíveis ***********")
            print(dados)
            
            u=dados.count()
            print('u:'+str(u))
            nn=len(dados['Aluno'])
            for i in range(nn):
                nome = dados['Aluno'][i]
                nota1 = str(dados['Nota1'][i])
                nota2 = str(dados['Nota2'][i])
                media = str(dados['Média'][i])
                situacao = dados['Situação'][i]
                
                self.treeMedias.insert('', 'end',
                                       iid=self.iid,
                                       values=(nome,
                                               nota1,
                                               nota2,
                                               media,
                                               situacao))
                
                self.iid = self.iid + 1
                self.id = self.id + 1
        except:
            print('Ainda não existem dados para carregar')
                
    def fSalvarDados(self):
          try:
              fsave = 'planilhaAlunos.xlsx'
              #fsave = 'PlanilhaAlunos.csv'
              dados=[]
              
              for line in self.treeMedias.get_children():
                  lstDados=[]
                  for value in self.treeMedias.item(line)['values']:
                      lstDados.append(value)
                      
                      dados.append(lstDados)
                      
                      df = pd.DataFrame(data=dados,columns=self.dadosColunas)
                 
                      planilha = pd.ExcelWriter(fsave)
                      
                      df.to_excel(planilha, 'Inconsistencias', index=False)
                 
                      planilha.save()
                      print('Dados salvos')
                      
          except:
              print('Não foi possível salvar os dados')
              
    def fCalcularMedia(self):
        try:
            nome = self.txtNome.get()
            nota1 = float(self.txtNota1.get())
            nota2 = float(self.txtNota2.get())
            media, situacao = self.fVerificarSituacao(nota1, nota2)
            
            
            self.treeMedias.insert('', 'end',
                                   iid=self.iid,
                                   values=(nome,
                                   str(nota1),
                                   str(nota2),
                                   str(media),
                                   situacao))
            
            
            self.iid = self.iid + 1
            self.id = self.id + 1
            
            self.fSalvarDados()
        except ValueError:
            print('Entre com valores válidos')
        finally:
            self.txtNome.delete(0, 'end')
            self.txtNota1.delete(0, 'end')
            self.txtNota2.delete(0, 'end')

        janela=tk.Tk()
        principal=PrincipalRAD(janela)
        janela.title('Bem vindo ao RAD')
        janela.geometry("820x600+10+10")
        janela.mainloop()
            
            