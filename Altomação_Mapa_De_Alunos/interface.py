import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os

def iniciar_automacao():
    tipo_contrato = contrato_var.get()
    trimestre = trimestre_var.get()
    login = login_var.get()
    senha = senha_var.get()

    if not tipo_contrato or not trimestre or not login or not senha:
        messagebox.showwarning("Atenção", "Por favor, preencha todos os campos antes de continuar.")
        return

    pasta_automacao = os.path.join(os.getcwd(), "automacao")
    os.makedirs(pasta_automacao, exist_ok=True)

    caminho_filtros = os.path.join(pasta_automacao, "filtros_selecionados.txt")
    with open(caminho_filtros, "w", encoding="utf-8") as f:
        f.write(f"{tipo_contrato}\n{trimestre}\n{login}\n{senha}")

    root.destroy()

    script_automacao = os.path.join("automacao", "extrair_dados.py")

    try:
        subprocess.run(["python", script_automacao], check=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Erro ao executar a automação:\n{e}")

def alternar_senha():
    if senha_entry.cget('show') == '':
        senha_entry.config(show='*')
        btn_mostrar_senha.config(text='Mostrar senha')
    else:
        senha_entry.config(show='')
        btn_mostrar_senha.config(text='Ocultar senha')

# --- Interface Gráfica ---
root = tk.Tk()
root.title("Filtro para Automação")
root.geometry("460x460")
root.configure(bg="#f0f4f8")

# Estilo geral
entry_style = {
    "bd": 1,
    "relief": "solid",
    "highlightthickness": 1,
    "highlightbackground": "#000000",
    "highlightcolor": "#000000",
    "font": ("Arial", 10),
    "width": 34
}

btn_style = {
    "bg": "#ffffff",
    "fg": "#000000",
    "borderwidth": 1,
    "highlightthickness": 1,
    "highlightbackground": "#000000",
    "relief": "solid",
    "font": ("Arial", 10),
    "activebackground": "#e6e6e6"
}

# Título
tk.Label(root, text="Configuração de Filtros", font=("Arial", 16, "bold"), bg="#f0f4f8").pack(pady=15)

# Tipo de Contrato
tk.Label(root, text="Tipo de Contrato:", bg="#f0f4f8", font=("Arial", 11)).pack(pady=(10, 3))
contrato_var = tk.StringVar()
contrato_cb = ttk.Combobox(root, textvariable=contrato_var, state="readonly", width=30,
                           values=["Geral", "Contrato Ativo", "Contrato Cancelado"])
contrato_cb.pack()

# Trimestre
tk.Label(root, text="Trimestre:", bg="#f0f4f8", font=("Arial", 11)).pack(pady=(15, 3))
trimestre_var = tk.StringVar()
trimestres = [f"{i}/{ano}" for ano in range(23, 26) for i in range(1, 5)] + ["Todos"]
trimestre_cb = ttk.Combobox(root, textvariable=trimestre_var, state="readonly", width=30,
                            values=trimestres)
trimestre_cb.pack()

# Login
tk.Label(root, text="Login:", bg="#f0f4f8", font=("Arial", 11)).pack(pady=(15, 3))
login_var = tk.StringVar()
tk.Entry(root, textvariable=login_var, **entry_style).pack()

# Senha
tk.Label(root, text="Senha:", bg="#f0f4f8", font=("Arial", 11)).pack(pady=(15, 3))
senha_var = tk.StringVar()
senha_entry = tk.Entry(root, textvariable=senha_var, show='*', **entry_style)
senha_entry.pack()

# Botão Mostrar/Ocultar senha
btn_mostrar_senha = tk.Button(root, text="Mostrar senha", command=alternar_senha, **btn_style)
btn_mostrar_senha.pack(pady=(5, 15))

# Botão de Iniciar
tk.Button(root, text="Iniciar Automação", command=iniciar_automacao, **btn_style, width=25).pack(pady=10)

root.mainloop()
