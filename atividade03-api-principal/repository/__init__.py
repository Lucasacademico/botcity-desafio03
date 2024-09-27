from repository import usuario
from repository import produto
from repository import database
#import mysql.connector

# Definindo uma variável global
versao = "1.0.0"

# Definindo uma função de inicialização
def inicializar():
    print("Pacote meu_pacote inicializado")

# Executando a função de inicialização ao importar o pacote
inicializar()

__all__ = ["repository"]
