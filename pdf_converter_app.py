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
        self.exist = True
        self.pasta_origem = ""
        self.pasta_destino = ""
        self.pasta_excel = ""
        self.tipo_conta = tk.StringVar()  # Variável para armazenar o tipo de conta
        self.tipo_conta.set("Nenhum")  # Valor padrão para nenhum radiobutton selecionado

        # Frame para organizar os widgets
        frame = tk.Frame(root)
        frame.pack()

        # Botão de seleção do tipo de conta
        self.lbl_tipo_conta = tk.Label(frame, text="Selecione o tipo de conta:")
        self.lbl_tipo_conta.pack()

        self.radio_agua = tk.Radiobutton(frame, text="Água", variable=self.tipo_conta, value="Agua", command=self.ativar_botao_origem)
        self.radio_agua.pack()

        self.radio_luz = tk.Radiobutton(frame, text="Luz", variable=self.tipo_conta, value="Luz", command=self.ativar_botao_origem)
        self.radio_luz.pack()

        # Botão de seleção de pasta de origem
        self.btn_selecionar_origem = tk.Button(frame, text="Selecionar Pasta de Origem", command=self.selecionar_pasta_origem)
        self.btn_selecionar_origem.pack(pady=10)
        self.btn_selecionar_origem["state"] = "disabled"  # Desativar inicialmente

        # Botão de conversão
        self.btn_converter = tk.Button(frame, text="Converter PDF's para Excel", command=self.converter_pdfs)
        self.btn_converter.pack(pady=10)
        self.btn_converter["state"] = "disabled"  # Desativar inicialmente

    def ativar_botao_origem(self):
        # Ativar o botão de seleção de pasta de origem quando o tipo de conta for escolhido
        self.btn_selecionar_origem["state"] = "active"

    def selecionar_pasta_origem(self):
        if self.tipo_conta.get() == "Nenhum":
            messagebox.showwarning("Aviso", "Por favor, selecione o tipo de conta primeiro.")
            return

        self.pasta_origem = filedialog.askdirectory(initialdir=Path.home(), title="Selecione a pasta")
        messagebox.showinfo("Pasta de Origem", f"Pasta de Origem selecionada: {self.pasta_origem}")
        self.btn_converter["state"] = "active"  # Ativar o botão de conversão
        
        tipo_conta_str = self.tipo_conta.get()
        self.pasta_destino = os.path.join(Path.home(), 'Desktop', f"PDF's Concluidos - {tipo_conta_str}")
        self.pasta_excel = os.path.join(Path.home(), 'Desktop', f"PDF's - EXCEL - {tipo_conta_str}")

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
            self.btn_converter["state"] = "disabled"  
            self.btn_selecionar_origem["state"] = "disabled"
            self.tipo_conta.set("Nenhum")

        elif self.tipo_conta.get() == 'Luz':
            processar_arquivos_luz(self.pasta_origem, self.pasta_destino, self.pasta_excel)
            messagebox.showinfo("Concluído", "PDFs convertidos com sucesso!")
            self.btn_converter["state"] = "disabled"  
            self.btn_selecionar_origem["state"] = "disabled"
            self.tipo_conta.set("Nenhum")

        else:
            print('Não foi possível efetuar a conversão!')

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
