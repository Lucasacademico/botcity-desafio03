from botcity.web import WebBot, Browser, By
import requests
from botcity.maestro import *
from webdriver_manager.chrome import ChromeDriverManager
from botcity.plugins.http import BotHttpPlugin
from datetime import datetime
import sys
import os
import planilha.planilha as planilha

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def inserir_usuario(usuario):
    url = 'http://127.0.0.1:5000/usuario'  # Ajuste o endpoint para usuários
    headers = {'Content-Type': 'application/json'}
    dados = {
        "nome": usuario['NOME'],
        "login": usuario['LOGIN'],
        "senha": usuario['SENHA'],
        "email": usuario['EMAIL']
    }

    try:
        resposta = requests.post(url=url, headers=headers, json=dados)
        resposta.raise_for_status()  # Verifica se houve erro
        retorno = resposta.json()  # Retorna a resposta em JSON
        print(f"Usuário inserido: {retorno}")
    except requests.exceptions.HTTPError as err:
        print(f"Erro HTTP ao inserir usuário: {err}")
    except Exception as err:
        print(f"Erro ao inserir usuário: {err}")


def main():
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()
    bot.headless = False
    bot.browser = Browser.CHROME
    bot.driver_path = ChromeDriverManager().install()

    # Opens the BotCity website.
    #bot.maximize_window()
    #bot.browse("www.ifam.edu.br")

    # Implement here your logic...
    print('inicio do processamento')

    # Leitura da planilha de usuários
    df_usuario = planilha.ler_excel('D:\\projetos-botcity\\Atividade 03 - Bot-Produto\\botproduto\\planilha\\RelacaoUsuario.xlsx', 'Usuarios')
    for index, usuario in df_usuario.iterrows():
        inserir_usuario(usuario)

    planilha.exibir_dados_excel(df_usuario)

    bot.wait(3000)
    bot.stop_browser()
   
    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


def not_found(label):
    print(f"Element not found: {label}")

if __name__ == '__main__':
    main()