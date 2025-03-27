from urllib.parse import quote
import time
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import re

# ------------------------- Classe para Conexão com WhatsApp -------------------------
class WhatsAppBot:
    def __init__(self):
        self.navegador = None

    def iniciar(self):
        """Inicia o WhatsApp Web no navegador"""
        self.navegador = webdriver.Chrome()
        self.navegador.get("https://web.whatsapp.com")
        messagebox.showinfo("Escaneie o QR Code", "Escaneie o QR Code no WhatsApp Web e clique em OK.")

    def enviar_mensagem(self, telefone, mensagem):
        """Envia uma mensagem de texto para um número específico"""
        self.navegador.get(f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}")

        while len(self.navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)

        input_box = self.navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div/p/span')
        input_box.send_keys(Keys.ENTER)

        time.sleep(2)

    def enviar_imagem(self, telefone, imagem_path, lengenda=""):
        """Envia uma imagem para um número específico"""
        self.navegador.get(f"https://web.whatsapp.com/send?phone={telefone}&text={quote(lengenda)}")

        while len(self.navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)

        while len(self.navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button')) < 1:
            time.sleep(1)

        self.navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button').click()

        while len(self.navegador.find_elements(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input')) < 1:
            time.sleep(1)

        self.navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input').send_keys(imagem_path)

        while len(self.navegador.find_elements(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div')) < 1:
            time.sleep(1)

        self.navegador.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div').click()

        time.sleep(3)

    def fechar(self):
        """Fecha o navegador"""
        if self.navegador:
            self.navegador.quit()

# ------------------------- Classe para Construção de Mensagens -------------------------
class MessageBuilder:
    @staticmethod
    def substituir_variaveis(texto, linha, headers):
        """Substitui variáveis {variavel} pelos valores da planilha"""
        variaveis = re.findall(r"\{(.*?)\}", texto)
        for var in variaveis:
            if var in headers:
                texto = texto.replace(f"{{{var}}}", str(linha[headers[var]].value))
            else:
                texto = texto.replace(f"{{{var}}}", f"[{var} NÃO ENCONTRADO]")
        return texto

# ------------------------- Classe para Interface Gráfica -------------------------
class AppUI:
    def __init__(self, root, main_app):
        self.root = root
        self.main_app = main_app
        self.message_list = []
        self.file_path = None

        self.criar_interface()

    def criar_interface(self):
        """Cria a interface gráfica"""
        self.root.title("WhatsApp Sender Message")
        self.root.geometry("600x500")

        # Botão para selecionar arquivo Excel
        tk.Button(self.root, text="Selecionar Lista de Contatos", command=self.selecionar_arquivo).pack(pady=5)
        self.file_label = tk.Label(self.root, text="Nenhum arquivo selecionado", fg="gray")
        self.file_label.pack()

        # Área para adicionar mensagens
        tk.Label(self.root, text="Digite uma mensagem (use {variavel} para personalizar):", font=("Arial", 10)).pack(pady=5)
        self.text_entry = tk.Text(self.root, height=3, width=50)
        self.text_entry.pack()

        tk.Button(self.root, text="➕ Adicionar Texto", command=self.adicionar_texto, bg="lightblue").pack(pady=5)
        tk.Button(self.root, text="🖼️ Adicionar Imagem", command=self.adicionar_imagem, bg="lightgreen").pack(pady=5)

        # Área de visualização da sequência
        tk.Label(self.root, text="📜 Sequência das Mensagens:", font=("Arial", 10, "bold")).pack(pady=5)
        self.preview_frame = tk.Frame(self.root)
        self.preview_frame.pack(fill="both", expand=True)

    def selecionar_arquivo(self):
        """Seleciona o arquivo de contatos"""
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.file_label.config(text=f"Arquivo: {self.file_path.split('/')[-1]}" if self.file_path else "Nenhum arquivo selecionado")

    def adicionar_texto(self):
        """Adiciona um texto à lista de mensagens e atualiza a interface"""
        text = self.text_entry.get("1.0", tk.END).strip()
        if text:
            self.message_list.append({"type": "text", "content": text})
            self.text_entry.delete("1.0", tk.END)
            self.atualizar_lista_mensagens()

    def adicionar_imagem(self):
        """Adiciona uma imagem e permite adicionar uma legenda"""
        file_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.jpg;*.jpeg;*.png")])
        if file_path:
            legenda = simpledialog.askstring("Legenda", "Digite a legenda para a imagem (opcional):")
            self.message_list.append({"type": "image", "content": file_path, "caption": legenda if legenda else ""})
            self.atualizar_lista_mensagens()

    def atualizar_lista_mensagens(self):
        """Atualiza a visualização da lista de mensagens"""
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        for i, msg in enumerate(self.message_list):
            text = f"{i + 1}. "
            if msg["type"] == "text":
                text += f"📝 Texto: {msg['content']}"
            elif msg["type"] == "image":
                text += f"🖼️ Imagem: {msg['content'].split('/')[-1]} (Legenda: {msg['caption']})"

            lbl = tk.Label(self.preview_frame, text=text, anchor="w", justify="left", bg="lightgray", padx=5, pady=3)
            lbl.pack(fill="x", padx=10, pady=2)

# ------------------------- Classe Principal -------------------------
class MainApp:
    def __init__(self):
        self.whatsapp_bot = WhatsAppBot()
        self.root = tk.Tk()
        self.ui = AppUI(self.root, self)

    def run(self):
        """Inicia a interface gráfica"""
        self.root.mainloop()

# ------------------------- Iniciar a Aplicação -------------------------
if __name__ == "__main__":
    app = MainApp()
    app.run()