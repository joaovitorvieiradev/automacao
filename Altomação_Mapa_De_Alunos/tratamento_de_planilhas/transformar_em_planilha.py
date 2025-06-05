import shutil
import os
import pandas as pd
import subprocess

caminho = "C://Users/usuário/Downloads"
lista_arquivos = os.listdir(caminho)

lista_datas = []
for arquivo in lista_arquivos:
    #Descobrir a data desse arquivo
    if "listagemAlunos" in arquivo:
        caminho_arquivo = os.path.join(caminho, arquivo)
        if os.path.isfile(caminho_arquivo):
            data = os.path.getmtime(f"{caminho}/{arquivo}")
            lista_datas.append((data, arquivo))


lista_datas.sort(reverse=True)
ultimo_arquivo = lista_datas[0]
nome_arquivo = lista_datas[0][1]
print(nome_arquivo)


 # Novo nome do arquivo
novo_nome = "BaseDeDados.xls"


# Caminho da pasta onde está este script
pasta_codigo = os.getcwd()
pasta_destino = os.path.join(f"{pasta_codigo}/tratamento_de_planilhas")


# Caminho completo de origem e destino
caminho_origem = os.path.join(caminho, nome_arquivo)
caminho_destino = os.path.join(pasta_destino, "Arquivo_HTML")

# Verifica se o arquivo realmente existe antes de mover
if os.path.exists(caminho_origem):
    shutil.move(caminho_origem, caminho_destino)
    print(f"Arquivo movido para: {caminho_destino}")
else:
    print(f"Arquivo não encontrado: {caminho_origem}")



# Caminho absoluto do arquivo .xls (na verdade é um HTML disfarçado)
arquivo_html = os.path.join(os.getcwd(),caminho_destino)

# Lê todas as tabelas do "HTML"
tabelas = pd.read_html(arquivo_html)

# Pega a segunda tabela, que contém os dados dos alunos
df = tabelas[1]

# Limpa colunas de espaços extras
df.columns = df.columns.str.strip()

# Limpa valores dentro do DataFrame
for col in df.select_dtypes(include='object').columns:
    df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)


# Renomeia colunas para facilitar o trabalho
df.columns = ["Matricula", "Nome", "Curso", "Turma", "Turno", "Grupo", "Data_Matricula"]

# Converte datas para o formato padrão do pandas
df["Data_Matricula"] = pd.to_datetime(df["Data_Matricula"], dayfirst=True, errors="coerce")


caminho_excel = os.path.join(pasta_destino, "BaseDeDados.xlsx")

df.to_excel(caminho_excel, index=False)
print(f"Arquivo salvo em: {caminho_excel}")




# Define o nome da pasta onde está o script que você quer rodar
nome_pasta = "tratamento_de_planilhas"

# Caminho completo da pasta (dentro do diretório atual)
pasta_tratamento = os.path.join(os.getcwd(), nome_pasta)

# Nome do script Python dentro dessa pasta
nome_script = "tratamento_de_BaseDeDados.py"

# Caminho completo do script dentro da pasta
caminho_script = os.path.join(pasta_tratamento, nome_script)

# Agora roda o script Python usando o subprocess
resultado = subprocess.run(["python", caminho_script], capture_output=True, text=True)

