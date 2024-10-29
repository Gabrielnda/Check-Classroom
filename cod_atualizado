from tkinter import *
from tkinter import messagebox
import tkinter as tk
import subprocess
import sqlite3
from datetime import datetime
import pandas as pd

window = Tk()

# Criação da janela principal:

class bdd():
    def conecta_bdd(self):
        self.conn = sqlite3.connect("check.bd")
        self.cursor = self.conn.cursor()
        print('Conectando ao BDD')
    
    def desconecta_bdd(self):
        self.conn.close()
        print('Desconectando ao BDD')
        

    def montaTabelas(self):
        self.conecta_bdd()
        # Criar tabela se não existir
        self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS verifica_sala (
                        Data TEXT,
                        Sala TEXT,
                        Org TEXT,
                        Anti TEXT,
                        SCCM TEXT,
                        Cam TEXT,
                        Proj TEXT,
                        GLPI TEXT,
                        Som TEXT,
                        Tel TEXT,
                        Gop TEXT,
                        Mic TEXT,
                        Observacoes TEXT
                );
        """)
        self.conn.commit()
        print('Banco de dados criado')
        self.desconecta_bdd()

    def add_bdd(self):
        self.conecta_bdd()

# Adiconando data
        data_atual = datetime.now().strftime("%d-%m-%Y")
        
# Adicionando valores ao banco de dados + FUNÇÃO GET PARA CORREÇÃO DO TIPO DE INFORMAÇÃO QUE É ARMAZENADA NO BDD
        self.cursor.execute(""" 
                INSERT INTO verifica_sala (Data, Sala, Org, Anti, SCCM, Cam, Proj, GLPI, Som, Tel, Gop, Mic, Observacoes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, ( 
            data_atual,
            self.combo_sala.get(),
            self.org.get(),
            self.bit.get(),
            self.status.get(),
            self.camera.get(),
            self.video.get(),
            self.glpi.get(),
            self.som.get(),
            self.tel.get(),
            self.go.get(),
            self.mic.get(),
            self.obs.get("1.0", END).strip()
        ))
        
        self.conn.commit()
        self.desconecta_bdd()

# Exportação para excel
    def export_to_excel(self, filename="Verificação_salas.xlsx"):
            self.conecta_bdd()
            query = "SELECT * FROM verifica_sala"
            df = pd.read_sql_query(query, self.conn)  # Lê os dados do banco de dados para um DataFrame pandas
            df.to_excel(filename, index=False)  # Salva o DataFrame como arquivo Excel
            self.desconecta_bdd()

#####

# Classe da aplicação principal
class Application(bdd):
    def __init__(self):
        self.window = window
        self.tela()
        self.frames_da_tela()
        self.action()
        
# Inicializando StringVar para os OptionMenu (Devido ao método que utilizei para captção dos dados precisou-se definir um sistema para que consiguisse colocar no BDD)
        self.combo_sala = StringVar()
        self.org = StringVar()
        self.bit = StringVar()
        self.status = StringVar()
        self.camera = StringVar()
        self.glpi = StringVar()
        self.som = StringVar()
        self.video = StringVar()
        self.mic = StringVar()
        self.peri = StringVar()
        self.tel = StringVar()
        self.go = StringVar()
        
        self.obs = Text(self.frame_1, height=5, width=30)  
        
        self.widgets_frame1()
        self.montaTabelas()
        self.action()  # Mova esta chamada para o final
        window.mainloop()

    
# Função que chama o método add_bdd para salvar os dados no banco
    def salvar_dados(self):
        try:
            print("Salvando dados...")
            self.add_bdd()
            messagebox.showinfo("Sucesso", "Dados salvos no banco de dados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar dados: {e}")

# Criação das telas
    def tela(self):
        self.window.title("Classroom Check")                
        self.window.configure(background='#707271')
        self.window.geometry("500x670")
        self.window.resizable(False,False)

#Criação janela interna
    def frames_da_tela(self):
        self.frame_1 = Frame(self.window, bd=4, highlightbackground= "#292421", highlightthickness=2)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.95)


# Botões de Procedimentos:
    def widgets_frame1(self):

        ## sala
        self.sala_menu = StringVar()
        self.sala_menu = Label(self.frame_1, text="Selecione a Sala:", font=('Courier New',10, 'bold')).grid(row=0, pady=(8,0), column=0)
        self.sala_menu = OptionMenu(self.frame_1, self.combo_sala, "Sala 91", "Sala 92", "Sala 81", "Sala 82", "Sala 71", "Sala 72", "Sala 61", "Sala 62", "Sala 51", "Sala 52", "Sala 41", "Sala 42")
        self.sala_menu.grid(row=0,sticky=W,columnspan=2,pady=(8,0), column=1)
 
        # Procedimentos
        
        self.tittle_Proce = Label(self.frame_1, text="---------------------------------------", font=('Times New Roman', 12, 'bold'))
        self.tittle_Proce.grid(row=1,columnspan=3,sticky=W, pady=(2,1), column=0)

        ## Bitdefender
        self.bit = StringVar()
        self.antivirus_menu = Label(self.frame_1, text="BitDefender:",font=('Courier New',10,'bold')).grid(row=2,sticky=E, column=0)
        self.antivirus_menu = OptionMenu(self.frame_1, self.bit, "OK", "FF", "DF")
        self.antivirus_menu.grid(row=2, column=1)
        ## Button Bitdefender
        self.button_bit = tk.Button(self.frame_1, text="Atualizar Antivírus", command=lambda: subprocess.run([r'C:\Program Files\Bitdefender\Endpoint Security\EPconsole.exe', '--update']))
        self.button_bit.grid(row=2,sticky=E+W, column=2)

        ## Câmera
        self.camera_menu = Label(self.frame_1, text="Câmera:", font=('Courier New',10, 'bold')).grid(row=4,sticky=E, column=0)
        self.camera = StringVar()
        self.camera_menu = OptionMenu(self.frame_1, self.camera, "OK", "FF", "DF")
        self.camera_menu.grid(row=4, column=1)
        ## Button Camera 
        self.button_cam = tk.Button(self.frame_1, text="Abrir Câmera", command=lambda: subprocess.run(['powershell', 'start', 'microsoft.windows.camera:']))
        self.button_cam.grid(row=4,sticky=E+W, column=2)
                ## O subprocess.run() executa um comando do sistema. 
                ## O lambda permite embutir esse comando diretamente no botão.
        
        ## GLPI 
        self.glpi_menu = Label(self.frame_1, text="GLPI:", font=('Courier New',10, 'bold')).grid(row=6,sticky=E, column=0)
        self.glpi = StringVar()
        self.glpi_menu = OptionMenu(self.frame_1, self.glpi, "OK", "FF", "DF")
        self.glpi_menu.grid(row=6, column=1)
        ## Button GLPI 
        self.button_glpi = tk.Button(self.frame_1, text="Verificar GLPI", command=lambda: subprocess.run(['powershell', 'start', 'http://localhost:62354/']))
        self.button_glpi.grid(row=6,sticky=E+W, column=2)

        ## Som
        self.som_menu = Label(self.frame_1, text="Som:", font=('Courier New',10, 'bold')).grid(row=7,sticky=E,column=0)
        self.som = StringVar()
        self.som_menu = OptionMenu(self.frame_1, self.som, "OK", "FF", "DF")
        self.som_menu.grid(row=7, column=1)
        ## Button som
        self.button_som = tk.Button(self.frame_1, text="Verificar SOM", command=lambda: subprocess.run(['powershell', '-c', '(New-Object Media.SoundPlayer "C:\\Windows\\Media\\Ring06.wav").PlaySync();']))
        self.button_som.grid(row=7,sticky=E+W, column=2)

        ## Projetor
        self.video_menu = Label(self.frame_1, text="Projetor:", font=('Courier New',10, 'bold')).grid(row=8,sticky=E, column=0)
        self.video = StringVar()
        self.video_menu = OptionMenu(self.frame_1, self.video, "OK", "FF", "DF")
        self.video_menu.grid(row=8, column=1)

        ## Microfone
        self.mic_menu = Label(self.frame_1, text="Microfone:", font=('Courier New',10, 'bold')).grid(row=9,sticky=E, column=0)
        self.mic = StringVar()
        self.mic_menu = OptionMenu(self.frame_1, self.mic, "OK", "FF", "DF")
        self.mic_menu.grid(row=9, column=1)

        ## Perfiféricos
        self.peri_menu = Label(self.frame_1, text="Periféricos:", font=('Courier New',10, 'bold')).grid(row=10,sticky=E,column=0)
        self.peri = StringVar()
        self.peri_menu = OptionMenu(self.frame_1, self.peri, "OK", "FF", "DF")
        self.peri_menu.grid(row=10, column=1)

        ## Telefone
        self.tel_menu = Label(self.frame_1, text="Telefone:", font=('Courier New',10, 'bold')).grid(row=11,sticky=E,column=0)
        self.tel = StringVar()
        self.tel_menu = OptionMenu(self.frame_1, self.tel, "OK", "FF", "DF")
        self.tel_menu.grid(row=11, column=1)

        ## Organização
        self.org_menu = Label(self.frame_1, text="Organização:", font=('Courier New',10, 'bold')).grid(row=12,sticky=E,column=0)
        self.org = StringVar()
        self.org_menu = OptionMenu(self.frame_1, self.org, "OK", "FF", "DF")
        self.org_menu.grid(row=12,columnspan=1, column=1)

        ## Gopresence
        self.go_menu = Label(self.frame_1, text="GoPresence:", font=('Courier New',10, 'bold')).grid(row=13,sticky=E,column=0)
        self.go = StringVar()
        self.go_menu = OptionMenu(self.frame_1, self.go, "OK", "FF", "DF")
        self.go_menu.grid(row=13,columnspan=1, column=1)

        ## Observação
        self.obs = Label(self.frame_1, text="Observações:", font=('Courier New',10, 'bold')).grid(row=14,pady=(20,0),sticky=W, column=0)
        self.obs = Text(self.frame_1,bd = 2,width=50, height=6)
        self.obs.grid(row=15,columnspan=6, pady=(12,0),padx=(20,0), column=0)
        
        # Botão para exportar dados
        self.bt_exportar = Button(self.frame_1, text="Exportar Dados", bd=3, bg='White', fg='black', command=self.export_to_excel)
        self.bt_exportar.grid(row=16, pady=(20,0), sticky=E+W, padx=(15,0), column=0)

    ## SCCM
        self.sccm_menu = Label(self.frame_1, text="SCCM:",font=('Courier New',10,'bold')).grid(row=3,sticky=E, column=0)
        self.status = StringVar()
        self.sccm_menu = OptionMenu(self.frame_1, self.status, "OK", "FF", "DF")
        self.sccm_menu.grid(row=3, column=1)
        ## Button SCCM
        self.button_sccm = tk.Button(self.frame_1, text="Executar SCCM", command=self.executar_sccm)
        self.button_sccm.grid(row=3, sticky=E+W, column=2)

# column 3
        ## Install antivírus
        self.button_sccm = tk.Button(self.frame_1, text="Install", font=('Courier New',7, 'bold'), bd=3, bg='black', fg='White', command=lambda: subprocess.run(['powershell', 'start', 'e:\Instalar\Antivirús\epskit_x64.exe']))
        self.button_sccm.grid(row=2, sticky=E, column=3)

        ## Install GLPI
        self.button_sccm = tk.Button(self.frame_1, text="Install", font=('Courier New',7, 'bold'), bd=3, bg='black', fg='White', command=lambda: subprocess.run(['powershell', 'start', 'e:\Instalar\glpi\GLPI-Agent-1.4-x64.msi']))
        self.button_sccm.grid(row=6, sticky=E, column=3)

        ## Button Forçar inventário
        self.button_sccm = tk.Button(self.frame_1, text="Force Invent", command=lambda: subprocess.run(['powershell', 'start', 'gpupdate /force']))
        self.button_sccm.grid(row=13, sticky=E, column=3)

# Execução do script do SCCM, verificar se está funcional ainda
    def executar_sccm(self):
        # Comando PowerShell que verifica e executa o ccmsetup.exe
        script = r'''
        $FileToDetect = "C:\Client\ccmsetup.exe"
        if (Test-Path $FileToDetect -PathType leaf) {
            cd 'C:\Client'
            .\ccmsetup.exe /mp:SCCM.dominio.interno SMSSITECODE=DHC FSP=SCCM
            Write-Output "Arquivo existe"
        } else {
            New-Item -ItemType directory -Path 'C:\Client'
            Copy-Item -Path \\10.10.15.112\Client\* -Destination 'C:\Client'
            cd 'C:\Client'
            .\ccmsetup.exe /mp:SCCM.dominio.interno SMSSITECODE=DHC FSP=SCCM
        }
        '''
        result = subprocess.run(['powershell', '-Command', script], shell=True)
        if result.returncode == 0:
            messagebox.showinfo("Sucesso", "O SCCM foi executado com sucesso!")
        else:
            messagebox.showerror("Erro", "Houve um erro ao executar o SCCM.")


# Botão de salvamento
    def action(self):
        self.bt_salvar = Button(self.frame_1, text="Salvar",font=('Courier New',10, 'bold'), bd=3, bg='green', fg='White', command=self.salvar_dados)
        self.bt_salvar.grid(row=16, pady=(20,0),sticky=E+W, padx=(0,0), column=3)

       


Application() 
