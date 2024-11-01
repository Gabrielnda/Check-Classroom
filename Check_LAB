from tkinter import *
from tkinter import messagebox
import tkinter as tk
import subprocess

window = Tk()

# Criação da janela principal:


# Classe da aplicação principal
class Application():
    def __init__(self):
        self.window = window
        self.tela()
        self.frames_da_tela()
        self.widgets_frame1()
        window.mainloop()

    # Função que chama o método add_bdd para salvar os dados no banco
    def tela(self):
        self.window.title("Lab Check")                
        self.window.configure(background='#B72727')#031A4F
        self.window.geometry("143x280")
        self.window.resizable(False,False)

    def frames_da_tela(self):
        self.frame_1 = Frame(self.window, bd=5, highlightbackground= "#292421", highlightthickness=2,background='#031A4F')
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.95)


# Botões de Procedimentos: 

    def widgets_frame1(self):

        ## Button Bitdefender
        self.button_bit = tk.Button(self.frame_1, text="Atualizar Antivírus",bd=4, command=lambda: subprocess.run([r'C:\Program Files\Bitdefender\Endpoint Security\EPconsole.exe', '--update']))
        self.button_bit.grid(row=1,sticky=E+W,pady=(10,0),padx=(5,0), column=0)

         
        ## Button GLPI 
        self.button_glpi = tk.Button(self.frame_1, text="Verificar GLPI",bd=4, command=lambda: subprocess.run(['powershell', 'start', 'http://localhost:62354/']))
        self.button_glpi.grid(row=3,sticky=E+W,pady=(10,0),padx=(5,0), column=0)

         ## Button som
        self.button_som = tk.Button(self.frame_1, text="Verificar SOM",bd=4, command=lambda: subprocess.run(['powershell', '-c', '(New-Object Media.SoundPlayer "C:\\Windows\\Media\\Ring06.wav").PlaySync();'], capture_output=True).returncode == 0 
        and messagebox.showinfo("Sucesso", "Som verificado com sucesso!")
    )
 
        self.button_som.grid(row=4,sticky=E+W,pady=(10,0),padx=(5,0), column=0)

        ## Projetor
        self.video_menu = tk.Button(self.frame_1, text="Verificar Projetor",bd=4, command=lambda: subprocess.run(['powershell', '-C', 'Start-Process "ms-settings:display"'], capture_output=True).returncode == 0 
        and messagebox.showinfo("Nota", "Verificar se projetor está na configuração PADRÃO!")
    )
 
        self.video_menu.grid(row=5,sticky=E+W,pady=(10,0),padx=(5,0), column=0)
   
        ## Button SCCM
        self.button_sccm = tk.Button(self.frame_1, text="Executar SCCM",bd=4, command=self.executar_sccm)
        self.button_sccm.grid(row=6, sticky=E+W,pady=(10,0),padx=(5,0), column=0)

        ## Button Forçar inventário
        self.button_invent = tk.Button(self.frame_1, text="Force Invent",bd=4, command=lambda: subprocess.run(['powershell', 'start', 'gpupdate /force'], capture_output=True).returncode == 0 
        and messagebox.showinfo("Sucesso", "Inventário executado com sucesso!")
    )
 
        self.button_invent.grid(row=7, sticky=E+W,pady=(10,0),padx=(5,0), column=0)

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

        


Application() 
