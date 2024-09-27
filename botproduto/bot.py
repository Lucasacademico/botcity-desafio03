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

# Definir as funções
def executar_api():
    http=BotHttpPlugin('https://economia.awesomeapi.com.br/last/USD-BRL')
    return http.get_as_json()

def inserir_produto(produto, valor_dolar):
    valor_atualizado = float(valor_dolar) * float(produto['PRECO_REAL'])
    
    url = 'http://127.0.0.1:5000/produto'
    headers = {'Content-Type': 'application/json'}
    dados = {
        "descricao" : produto['DESCRICAO'],
        "unidade"   : produto['UNIDADE'],
        "quantidade" : produto['QUANTIDADE'],
        "preco_real" : produto['PRECO_REAL'],
        "preco_dolar": valor_atualizado,
    }

    try:
        resposta = requests.post(url=url,headers=headers, json=dados)
        resposta.raise_for_status()  # Verifica se houve algum erro na requisição
        # Retorna a resposta em formato JSON
        retorno = resposta.json()
        #retorno['']  
    except requests.exceptions.HTTPError as err:
        print(f"Erro HTTP: {err}")
    except Exception as err:
        print(f"Erro: {err}")



def main():
    
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME
    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()

    # Opens the BotCity website.
    #bot.maximize_window()
    #bot.browse("www.ifam.edu.br")

    # Implement here your logic...
    print('inicio do processamento')

    retornoJSON = executar_api()
    valor_dolar=retornoJSON['USDBRL']['high']
    # Processar todos os elementos da planilha
    print('Exibir planilha')
    df = planilha.ler_excel('D:\\projetos-botcity\\Atividade 03 - Bot-Produto\\botproduto\\planilha\\RelacaoProduto.xlsx', 'Plan1')
    for index, produto in df.iterrows():
        inserir_produto(produto,valor_dolar)
   
    planilha.exibir_dados_excel(df)

    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
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