import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from tkinter import ttk

def calcular_imc():
    peso_str = entry_peso.get().replace(',', '.')
    altura_str = entry_altura.get().replace(',', '.')

    try:
        peso = float(peso_str)
        altura = float(altura_str)
    except ValueError:
        label_resultado["text"] = "Por favor, insira números válidos para peso e altura."
        return

    imc = peso / (altura ** 2)
    imc = round(imc, 2)
    label_resultado["text"] = f"Seu IMC é: {imc}"

    workbook_path = entry_caminho.get() 

    if workbook_path:
        workbook = load_workbook(workbook_path)
        sheet = workbook.active
        sheet.append([peso_str, altura_str, imc])
        workbook.save(workbook_path)
        
    else:
        label_resultado["text"] = "Por favor, insira o caminho para o arquivo da planilha."
    
    entry_peso.delete(0, 'end')
    entry_altura.delete(0, 'end')

def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
    entry_caminho.delete(0, tk.END)
    entry_caminho.insert(0, arquivo)

root = tk.Tk()
root.title("Calculadora de IMC")
root.geometry("500x300")
root.resizable(False, False)
root.configure(bg="#4CAF50")


style = ttk.Style()
style.theme_use("default")
style.configure("Green.TEntry", fieldbackground="#2E7D32", foreground="black", font=('Arial', 10, 'bold'))  # Configuração da fonte para a Entry
style.map("Green.TEntry",
          background=[('disabled', '#2E7D32'), ('readonly', '#2E7D32')])

frame = tk.Frame(root, bg="#4CAF50") 
frame.pack(expand=True, pady=30)

font_bold = ('Arial', 10, 'bold')  

label_peso = tk.Label(frame, text="Digite seu peso (kg):", bg="#4CAF50", fg="black", font=font_bold)  # Texto preto sobre fundo verde
label_peso.grid(row=0, column=0, padx=10, pady=5)

entry_peso = ttk.Entry(frame, style="Green.TEntry")  
entry_peso.grid(row=0, column=1, padx=10, pady=5)

label_altura = tk.Label(frame, text="Digite sua altura (m):", bg="#4CAF50", fg="black", font=font_bold)
label_altura.grid(row=1, column=0, padx=10, pady=5)

entry_altura = ttk.Entry(frame, style="Green.TEntry")
entry_altura.grid(row=1, column=1, padx=10, pady=5)

label_caminho = tk.Label(frame, text="Caminho para o arquivo da planilha:", bg="#4CAF50", fg="black", font=font_bold)
label_caminho.grid(row=2, column=0, padx=10, pady=5)

entry_caminho = ttk.Entry(frame, style="Green.TEntry")
entry_caminho.grid(row=2, column=1, padx=10, pady=5)

button_selecionar_arquivo = tk.Button(frame, text="Selecionar Arquivo", command=selecionar_arquivo)
button_selecionar_arquivo.grid(row=3, columnspan=2, padx=10, pady=5)

button_calcular = tk.Button(frame, text="Calcular IMC", command=calcular_imc)
button_calcular.grid(row=4, columnspan=2, padx=10, pady=5)

label_resultado = tk.Label(frame, text="", bg="#4CAF50", fg="black", font=font_bold)  # Texto preto sobre fundo verde
label_resultado.grid(row=5, columnspan=2, padx=10, pady=5)


root.mainloop()
