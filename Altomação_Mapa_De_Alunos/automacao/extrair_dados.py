from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import os
from time import sleep
from datetime import datetime
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
chromedriver_autoinstaller.install()
import subprocess
import tkinter as tk
from tkinter import messagebox
import sys
from selenium.common.exceptions import TimeoutException


# Função para pegar os dados dos filtros selecionados
def ler_filtros():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    caminho_arquivo = os.path.join(base_dir, "filtros_selecionados.txt")
    
    with open(caminho_arquivo, "r") as arquivo:
        tipo_contrato = arquivo.readline().strip()
        trimestre = arquivo.readline().strip()
        login = arquivo.readline().strip()
        senha = arquivo.readline().strip()
    return tipo_contrato, trimestre, login, senha


tipo_contrato, trimestre, login, senha = ler_filtros()


# Faz a altomação em segundo plano

usar_headless = True  # ou False para desabilitar o modo headless

options = Options()
if usar_headless:
    options.add_argument("--headless=new")  # Usa o novo modo headless do Chrome
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")  # Tamanho de tela virtual
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

navegador = webdriver.Chrome(options=options)

# Entrar no site do site_exemplo.com e ir até a página de filtros.
navegador.get("site_exemplo.com")  # Substitua pelo URL correto do Pincel Atômico


wait = WebDriverWait(navegador, 10)

# Colocar o meu usuário
usuario = wait.until(EC.presence_of_element_located((By.ID, "usuario")))
usuario.click()
usuario.send_keys(login)

# Colocar a minha senha
colocar_senha = wait.until(EC.presence_of_element_located((By.ID, "senha")))
colocar_senha.click()
colocar_senha.send_keys(senha)

# Clicando no botão de acesso
acessar = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value = 'Acessar']")))
acessar.click()
sleep(2)

# Tenta detectar se apareceu mensagem de erro de login em até 5 segundos
try:
    erro_login = WebDriverWait(navegador, 5).until(
        EC.presence_of_element_located((By.CLASS_NAME, "alert-danger"))
    )

    if erro_login.is_displayed():
        navegador.quit()

        def mostrar_erro_login():
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erro de Login", "Usuário ou senha incorretos. Reinicie a automação e tente novamente.")
            root.destroy()

        mostrar_erro_login()
        sys.exit()

except TimeoutException:
    pass  # Não apareceu erro de login, pode continuar normalmente

# Página de filtros
navegador.get("site_exemplo.com/filtros")  # Substitua pelo URL correto da página de filtros

# Abrir os dados do curso/matrícula
dados_curso_matricula = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-target='#camposCurso']")))
dados_curso_matricula.click()
navegador.execute_script("arguments[0].scrollIntoView(true);", dados_curso_matricula)

# Marcando as opções de filtros
curso = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and contains(@value, \"Curso\")]")))
navegador.execute_script("arguments[0].scrollIntoView(true);", curso)
curso.click()

turma = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and contains(@value, \" TURMA\")]")))
navegador.execute_script("arguments[0].scrollIntoView(true);", turma)
turma.click()

turno = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and contains(@value, \" Turno\")]")))
navegador.execute_script("arguments[0].scrollIntoView(true);", turno)
turno.click()

grupo = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and contains(@value, \" 'Grupo'\")]")))
navegador.execute_script("arguments[0].scrollIntoView(true);", grupo)
grupo.click()

data_matricula = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox' and contains(@value, \" 'Data de Matrícula'\")]")))
navegador.execute_script("arguments[0].scrollIntoView(true);", data_matricula)
data_matricula.click()

# Escolhendo gerar o arquivo em Excel
excel = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='radio'][value='excel']")))
navegador.execute_script("arguments[0].scrollIntoView(true);", excel)
excel.click()

# Colocando o nível (Curso Livre)
nivel = wait.until(EC.presence_of_element_located((By.ID, "nivel")))
nivel.click()
curso_livre = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[@value='Curso Livre']")))
curso_livre.click()

grupo = wait.until(EC.presence_of_element_located((By.ID, "grupo")))
navegador.execute_script("arguments[0].scrollIntoView(true);", grupo)
grupo.click()

#Colocando tratamento para todos os trimestres.
if trimestre != "Todos":
    try:
        select_grupo = Select(grupo)
        select_grupo.select_by_value(trimestre)
    except Exception as e:
        print(f"[ERRO] Não foi possível selecionar o trimestre '{trimestre}': {e}")

# Selecionar a unidade
unidade = wait.until(EC.presence_of_element_located((By.ID, "unidade")))
navegador.execute_script("arguments[0].scrollIntoView(true);", unidade)

# Usando Select para interagir com o <select> 'unidade' corretamente
select_unidade = Select(unidade)
select_unidade.select_by_value("unidade_exemplo")

# Coloca a data de inicio de matriculas para 01/01/2023
inicio_matricula = wait.until(EC.presence_of_element_located((By.ID, "inicio")))
inicio_matricula.click()
inicio_matricula.send_keys("01012023")

# Coloca a data de termino de matriculas para a data de hoje
data_atual = datetime.now().strftime("%d-%m-%Y")
final_matricula = wait.until(EC.presence_of_element_located((By.ID, "fim")))
final_matricula.click()
final_matricula.send_keys(data_atual)

# Filtrando o tipo de contrato
situacao_contrato = wait.until(EC.presence_of_element_located((By.ID, "situacaoC")))
navegador.execute_script("arguments[0].scrollIntoView(true);", situacao_contrato)

# Usando Select para interagir com o <select> 'situacaoC' corretamente
select_situacao = Select(situacao_contrato)

# Dicionário para mapear os nomes amigáveis para os valores do HTML
mapa_contratos = {
    "Geral": "GC",
    "Contrato Ativo": "ativo",
    "Contrato Cancelado": "cancelado"
}

# Define o valor correto com base no filtro selecionado
valor_contrato = mapa_contratos.get(tipo_contrato)

# Define o valor no <select> e dispara o evento change
select_situacao.select_by_value(valor_contrato)

gerar_listagem = wait.until(EC.presence_of_element_located((By.ID, "Filtrar")))
gerar_listagem.click()
sleep(5)
navegador.quit()


# Define o nome da pasta onde está o script que você quer rodar
nome_pasta = "tratamento_de_planilhas"

# Caminho completo da pasta (dentro do diretório atual)
pasta_tratamento = os.path.join(os.getcwd(), nome_pasta)

# Nome do script Python dentro dessa pasta
nome_script = "transformar_em_planilha.py"

# Caminho completo do script dentro da pasta
caminho_script = os.path.join(pasta_tratamento, nome_script)

# Agora roda o script Python usando o subprocess
resultado = subprocess.run(["python", caminho_script], capture_output=True, text=True)


