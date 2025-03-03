import tkinter as tk
from tkinter import ttk
import webbrowser
import subprocess
import os

# Dados dos produtos
produtos = {
    "O365ProPlusRetail": {
        "aplicativos": ["Access", "Excel", "Lync", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word", "OneDrive"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/O365ProPlusRetail.img"
    },
    "O365BusinessRetail": {
        "aplicativos": ["Access", "Excel", "Lync", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word", "OneDrive"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365BusinessRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/O365BusinessRetail.img"
    },
    "O365EduCloudRetail": {
        "aplicativos": ["Excel", "OneNote", "PowerPoint", "Word", "OneDrive"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365EduCloudRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "",  # Não fornecido
        "offline": ""  # Não fornecido
    },
    "O365HomePremRetail": {
        "aplicativos": ["Access", "Excel", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word", "OneDrive"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365HomePremRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/O365HomePremRetail.img"
    },
    "O365SmallBusPremRetail": {
        "aplicativos": ["Access", "Excel", "Lync", "OneNote", "Outlook", "PowerPoint", "Publisher", "Word", "OneDrive"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365SmallBusPremRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365SmallBusPremRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": ""  # Não fornecido
    },
    "AccessRuntimeRetail": {
        "aplicativos": ["Access"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=AccessRuntimeRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=AccessRuntimeRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": ""  # Não fornecido
    },
    "ProjectProRetail": {
        "aplicativos": ["Project"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectProRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectProRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/ProjectProRetail.img"
    },
    "VisioProRetail": {
        "aplicativos": ["Visio", "OneDrive"],
        "online_x64": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioProRetail&platform=x64&language=pt-br&version=O16GA",
        "online_x32": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioProRetail&platform=x86&language=pt-br&version=O16GA",
        "offline": "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/pt-br/VisioProRetail.img"
    }
}

class InstaladorOffice:
    def __init__(self, root):
        self.root = root
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        self.root.title("Instalador Office by Techrizen")
        self.produto_selecionado = tk.StringVar()
        self.tipo_instalacao = tk.StringVar()
        self.plataforma = tk.StringVar()

        # Seleção do produto
        tk.Label(root, text="Selecione o produto:").pack()
        for produto in produtos.keys():
            tk.Radiobutton(root, text=produto, variable=self.produto_selecionado, value=produto).pack()

        # Seleção do tipo de instalação
        tk.Label(root, text="Selecione a arquitetura:").pack()
        tk.Radiobutton(root, text="Online (x64)", variable=self.tipo_instalacao, value="online_x64").pack()
        tk.Radiobutton(root, text="Online (x32)", variable=self.tipo_instalacao, value="online_x32").pack()
        tk.Radiobutton(root, text="Offline", variable=self.tipo_instalacao, value="offline").pack()

        # Botão para iniciar a instalação
        tk.Button(root, text="Iniciar Instalação", command=self.iniciar_instalacao).pack()

        # Botão para executar script PowerShell
        tk.Button(root, text="Ativador", command=self.executar_powershell).pack()


    def iniciar_instalacao(self):
        produto = self.produto_selecionado.get()
        tipo_instalacao = self.tipo_instalacao.get()

        if tipo_instalacao == "online_x64":
            url = produtos[produto]["online_x64"]
        elif tipo_instalacao == "online_x32":
            url = produtos[produto]["online_x32"]
        elif tipo_instalacao == "offline":
            url = produtos[produto]["offline"]

        if url:
            webbrowser.open(url)
        else:
            print("URL não disponível para o tipo de instalação escolhido.")

    def executar_powershell(self):
        try:
            subprocess.run(
                ["powershell", "-Command", "Start-Process powershell -Verb RunAs -ArgumentList '-Command irm https://get.activated.win | iex'"],
                check=True
            )
        except subprocess.CalledProcessError as e:
            print(f"Erro ao executar o PowerShell: {e}")
        except FileNotFoundError:
            print("PowerShell não encontrado.")


if __name__ == "__main__":
    root = tk.Tk()
    app = InstaladorOffice(root)
    root.mainloop()
