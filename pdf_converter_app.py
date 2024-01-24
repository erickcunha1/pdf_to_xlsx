import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path
from leitorContas import processar_arquivos
from leitorLuz import processar_arquivos_luz
import os

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")

        self.pasta_origem = ""
        self.pasta_destino = ""
        self.pasta_excel = ""
        self.tipo_conta = tk.StringVar()  # Variável para armazenar o tipo de conta
        self.tipo_conta.set("Agua")  # Valor padrão
        self.exist = True
        # Botões
        self.btn_selecionar_origem = tk.Button(root, text="Selecionar Pasta de Origem", command=self.selecionar_pasta_origem)
        self.btn_selecionar_origem.pack(pady=10)

        # Botão de seleção do tipo de conta
        self.lbl_tipo_conta = tk.Label(root, text="Selecione o tipo de conta:")
        self.lbl_tipo_conta.pack()

        self.radio_agua = tk.Radiobutton(root, text="Água", variable=self.tipo_conta, value="Agua")
        self.radio_agua.pack()

        self.radio_luz = tk.Radiobutton(root, text="Luz", variable=self.tipo_conta, value="Luz")
        self.radio_luz.pack()

        self.btn_converter = tk.Button(root, text="Converter PDF's para Excel", command=self.converter_pdfs)
        self.btn_converter.pack(pady=10)

    def selecionar_pasta_origem(self):
        self.pasta_origem = filedialog.askdirectory(initialdir=Path.home(), title="Selecione a pasta")
        messagebox.showinfo("Pasta de Origem", f"Pasta de Origem selecionada: {self.pasta_origem}")

        # Cria as pastas de destino e Excel na área de trabalho se não existirem
        self.pasta_destino = os.path.join(Path.home(), 'Desktop', f"PDF's Concluidos - {self.tipo_conta.get()}")
        self.pasta_excel = os.path.join(Path.home(), 'Desktop', f"PDF's - EXCEL - {self.tipo_conta.get()}")
        self.criar_diretorio_se_nao_existir(self.pasta_destino)
        self.criar_diretorio_se_nao_existir(self.pasta_excel)
        
        if not self.exist:
            messagebox.showinfo("Pastas Criadas", f"As pastas de destino e Excel foram criadas na área de trabalho.")

    def converter_pdfs(self):
        if not self.pasta_origem:
            messagebox.showerror("Erro", "Por favor, selecione a pasta de origem dos PDFs.")
            return
        
        if self.tipo_conta.get() == 'Agua':
            processar_arquivos(self.pasta_origem, self.pasta_destino, self.pasta_excel)
            messagebox.showinfo("Concluído", "PDFs convertidos com sucesso!")
        elif self.tipo_conta.get() == 'Luz':
            processar_arquivos_luz(self.pasta_origem, self.pasta_destino, self.pasta_excel)
            messagebox.showinfo("Concluído", "PDFs convertidos com sucesso!")
        else:
            print('Não foi possivel efetuar a conversão!')
 
    def criar_diretorio_se_nao_existir(self, diretorio):
        caminho_completo = os.path.join(Path.home(), 'Desktop', diretorio)
    
        if not Path(caminho_completo).exists():
            Path(caminho_completo).mkdir(parents=True)
            self.exist = False

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry('500x200')
    app = PDFConverterApp(root)
    root.mainloop()
