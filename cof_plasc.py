import sqlite3
import tkinter as tk
import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from time import sleep


janela = tk.Tk()
janela.title('Registro de COF') 
janela. geometry("825x500")
    
    

def cadastrar_cof():
    conexao = sqlite3.connect('cof.db')
    c = conexao.cursor()
    #Inserir dados na tabela:
    c.execute("INSERT INTO cof VALUES (:data,:op,:resultado,:descrição)",
              {
                  'data': entry_data.get(),
                  'op': entry_op.get(),
                  'resultado': entry_resultado.get(),
                  'descrição': entry_descricao.get()
              })
    # Inserir dados na tabela:
    c.execute("SELECT *, oid FROM cof")
    cof_cadastrados = c.fetchall()
    
    # add data to the treeview
    
    tree.insert('', tk.END, values=cof_cadastrados[-1])
    # Commit as mudanças:
    conexao.commit()
    # Fechar o banco de dados:
    conexao.close()
    
    # #Apaga os valores das caixas de entrada
    entry_data.delete(0,"end")
    entry_op.delete(0,"end")
    entry_resultado.delete(0,"end")
    entry_descricao.delete(0,"end")

def exporta_cof():
    conexao = sqlite3.connect('cof.db')
    c = conexao.cursor()

    # Inserir dados na tabela:
    c.execute("SELECT *, oid FROM cof")
    cof_cadastrados = c.fetchall()
    cof_cadastrados = pd.DataFrame(cof_cadastrados,columns=['data','op','resultado','descrição','Id_banco'])
    cof_cadastrados.to_excel('cof.xlsx')
    print('dados exportado para excel')
    # Commit as mudanças:
    conexao.commit()

    # Fechar o banco de dados:
    conexao.close()
    sleep(0.5)
    Label(janela, text="Dados exportados com sucesso!",  
    font="arial 12 bold", bg='white').place(x=300, y=450) 

    
# define columns
columns = ('data', 'op', 'resultado', 'descricao')

tree = ttk.Treeview(janela, columns=columns, show='headings')

# define headings
tree.heading('data', text='Data')
tree.heading('op', text='OP')
tree.heading('resultado', text='Resultado')
tree.heading('descricao', text='Descrição')

conexao = sqlite3.connect('cof.db')
c = conexao.cursor()

# Inserir dados na tabela:
c.execute("SELECT *, oid FROM cof")
cof_cadastrados = c.fetchall()

# add data to the treeview
for cof in cof_cadastrados:
    tree.insert('', tk.END, values=cof)

tree.grid(row=0, column=1, sticky='ns')

# add a scrollbar
scrollbar = ttk.Scrollbar(janela, orient=tk.VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.grid(row=0, column=2, sticky='ns')

# Enviar dados para o Email
def Enviar():
    root = Tk()
    root.geometry("375x450")
    root.configure(bg='white')
    root.title(" Enviar email")

    Label(root, text="Enviar email", font="arial 15 bold", bg='white').pack()

    Msg = StringVar()
    Ass = StringVar()
    Dest = StringVar()

    Label(root, text='Assunto:', font='font 10 bold', bg='white smoke').place(x=20,y=35)
    mail_assunto = Entry(root, textvariable=Ass, width='50', bg='white smoke')
    mail_assunto.place(x=20, y=60)

    Label(root, text='Destinatário:', font='font 10 bold', bg='white smoke').place(x=20,y=100)
    destinatario = Entry(root, textvariable=Dest, width='50', bg='white smoke')
    destinatario.place(x=20, y=125)

    Label(root, text='Observação:', font='font 10 bold', bg='white smoke').place(x=20,y=170)
    mail_texto = Text(root, bg='white smoke', font='font 10 bold')
    mail_texto.place(x=20, y=205, width=300, height=100)

    Label(root, text='Arquivo em anexo: cof.xlsx', font='font 10 bold', bg='white smoke').place(x=20,y=350)
   


    def EnviarEmail():
        mensagem = mail_texto.get("1.0","end")
        assunto = mail_assunto.get()
        endereco_gmail = "desouza850@gmail.com"
        senha_app = "soukxnrfzvkstnfg"
        mail_de = "desouza850@gmail.com"
        mail_para = destinatario.get()
        excelName = 'cof.xlsx'
        
        fp = open(excelName, 'rb')
        anexo = MIMEApplication(fp.read(), _subtype="xlsx")
        fp.close()
        anexo.add_header('Content-Disposition', 'attachment', filename=excelName)
        
        mimemsg = MIMEMultipart()
        mimemsg['From'] = mail_de
        mimemsg['To'] = mail_para
        mimemsg['Subject'] = assunto
        mimemsg.attach(MIMEText(mensagem, 'plain')) 
        mimemsg.attach(anexo)
        

        connection = smtplib.SMTP(host='smtp.gmail.com', port=587)
        connection.starttls()
        connection.login(endereco_gmail, senha_app)
        connection.send_message(mimemsg)
        connection.quit()
        sleep(0.5)
        Label(root, text="Seu email foi enviado com sucesso!", 
            font="arial 8 bold", bg='white').place(x=20, y=300)  
        sleep(1)
        root.destroy()
        return
    
    def Sair():
        root.destroy()


    Button(root, text='Enviar', font='arial 10 bold', command=EnviarEmail,
        bg='white smoke').place(x=20, y=400,  width='80')
    Button(root, text="Sair", font='arial 10 bold', command=Sair, bg='red').place(x=105, y=400,  width='80')



    root.mainloop()

def Sair():
    janela.destroy()
    
def deletar():
    conexao = sqlite3.connect('cof.db')
    try:
        selecionado=tree.selection()[0]
        tree.delete(selecionado)  
    except: 
        messagebox.showinfo(title="ERRO", message='Selecione a linha que deseja deletar')
#Rótulos Entradas:
label_data = tk.Label(janela, text='Data')
label_data.place(x=150, y=250, width=120)

label_op = tk.Label(janela, text='OP')
label_op.place(x=150, y=300, width=120)

label_resultado = tk.Label(janela, text='Resultado')
label_resultado.place(x=400, y=250, width=120)

label_descricao = tk.Label(janela, text='Descrição')
label_descricao.place(x=400, y=300, width=120)

#Caixas Entradas:
entry_data = tk.Entry(janela , width=40)
entry_data.place(x=250, y=250, width=120)

entry_op = tk.Entry(janela, width=40)
entry_op.place(x=250, y=300, width=120)

entry_resultado = tk.Entry(janela, width=40)
entry_resultado.place(x=500, y=250, width=120)

entry_descricao = tk.Entry(janela, width=40)
entry_descricao.place(x=500, y=300, width=120)


# Botão de Cadastrar

botao_cadastrar = tk.Button(text='Cadastrar COF', command=cadastrar_cof, bg='blue')
botao_cadastrar.place(x=250, y=350, width=120)

# Botão de Exportar

botao_exportar = tk.Button(text='Exportar para Excel', command=exporta_cof, bg='green')
botao_exportar.place(x=200, y=400, width=120)

# Botão de Visualizar

botao_enviar = tk.Button(text='Enviar para email', command=Enviar, bg='yellow')
botao_enviar.place(x=350, y=400, width=120)

botao_deletar = tk.Button(text='Deletar', command=deletar , bg='red')
botao_deletar.place(x=500, y=350, width=120)

botao_deletar = tk.Button(text='Sair', command=Sair)
botao_deletar.place(x=500, y=400, width=120)

janela.mainloop()