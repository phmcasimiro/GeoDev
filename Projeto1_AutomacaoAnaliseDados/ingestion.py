"""
PROJETO: AUTOMAÇÃO E ANÁLISE DE DADOS COM PYTHON
------------------------------------------------
Este script realiza o download automatizado de uma base de dados no Google Drive
e aciona a geração de um relatório estatístico via Quarto.
"""

import os
import time
import logging
import subprocess
from pathlib import Path

import pyautogui
from dotenv import load_dotenv

# Configuração de Logging (Registra as ações e erros)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# Variáveis Globais
DRIVE_LINK = os.getenv("DRIVE_LINK")
NOME_ARQUIVO = os.getenv("NOME_ARQUIVO_VENDAS")
PASTA_DOWNLOADS = Path.home() / "Downloads"
CAMINHO_ARQUIVO = PASTA_DOWNLOADS / NOME_ARQUIVO

# Tempo padrão entre comandos do pyautogui
pyautogui.PAUSE = 1

def abrir_navegador_e_acessar(link):
    """Abre o Microsoft Edge e navega até o link especificado."""
    logging.info("Abrindo o navegador Edge...")
    pyautogui.press('win')
    pyautogui.write('edge')
    pyautogui.press('enter')
    time.sleep(3) # Aguarda o navegador abrir

    logging.info(f"Acessando o link: {link}")
    pyautogui.write(link)
    pyautogui.press('enter')
    time.sleep(5) # Aguarda o site carregar

def baixar_planilha():
    """Navega na interface do Google Drive para realizar o download."""
    logging.info("Iniciando navegação para download da planilha...")
    
    # Duplo click na pasta do drive
    pyautogui.click(x=356, y=385, clicks=2)
    time.sleep(2)
    
    # Clicar em 3 pontos para abrir a lista de ações
    pyautogui.click(x=925, y=382)
    time.sleep(2)
    
    # Clicar em transferir/download
    logging.info("Iniciando o download do arquivo...")
    pyautogui.click(x=967, y=347)
    
    # Aguarda o download concluir (ajustável dependendo da internet)
    time.sleep(15)

def gerar_relatorio():
    """Chama o Quarto para compilar o relatório .qmd."""
    logging.info("Iniciando a geração do relatório via Quarto...")
    caminho_qmd = Path("Projeto1_AutomacaoAnaliseDados") / "sells_report.qmd"
    
    # Executa o Quarto de forma síncrona
    subprocess.run(["quarto", "render", str(caminho_qmd)], shell=True, check=True)
    logging.info("Relatório gerado com sucesso!")

def limpar_arquivos_temporarios(caminho):
    """Deleta a planilha baixada após o uso."""
    logging.info("Limpando arquivos temporários...")
    if caminho.exists():
        caminho.unlink()
        logging.info("Arquivo da base de dados deletado com segurança.")
    else:
        logging.warning("O arquivo não foi encontrado para deleção.")

def main():
    """Função principal que orquestra todo o processo de automação."""
    logging.info("=== INICIANDO O PROCESSO DE AUTOMAÇÃO ===")
    
    if not DRIVE_LINK or not NOME_ARQUIVO:
        logging.error("Variáveis de ambiente DRIVE_LINK ou NOME_ARQUIVO_VENDAS não encontradas no .env")
        return

    try:
        abrir_navegador_e_acessar(DRIVE_LINK)
        baixar_planilha()
        
        # O Quarto vai usar o CAMINHO_ARQUIVO, então ele deve existir neste momento
        if not CAMINHO_ARQUIVO.exists():
            logging.error(f"O arquivo não foi encontrado em {CAMINHO_ARQUIVO}. Abortando.")
            return

        gerar_relatorio()
        limpar_arquivos_temporarios(CAMINHO_ARQUIVO)
        
        logging.info("=== PROCESSO DE AUTOMAÇÃO FINALIZADO COM SUCESSO ===")
        
    except subprocess.CalledProcessError as e:
        logging.error(f"Erro ao gerar o relatório com Quarto: {e}")
    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado durante a automação: {e}")

if __name__ == "__main__":
    main()
