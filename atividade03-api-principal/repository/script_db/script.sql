-- Banco de Dados MySQL
-- Apagar o banco de dados
drop database banco;
-- Criar o banco de dados
create database banco;
-- Atribuir os privilégios de acesso aos objetos do banco
-- para o usuário root
GRANT ALL PRIVILEGES ON banco.* TO 'root' @'localhost';
-- Acesar o banco de dados: banco
USE banco;
-- Criar a tabela: usuario
CREATE TABLE usuario(
    id int AUTO_INCREMENT,
    nome varchar(50) NOT NULL,
    login varchar(20) NOT NULL,
    senha varchar(20) NOT NULL,
    email varchar(50) NOT NULL,
    PRIMARY KEY (id)
);

CREATE TABLE produto(
    id int AUTO_INCREMENT,
    descricao varchar(50) NOT NULL,
    unidade varchar(5) NOT NULL,
    quantidade DECIMAL(10,2) NOT NULL,
    preco_real DECIMAL(10,2) NOT NULL,
    preco_dolar DECIMAL(10,2) NOT NULL,
    PRIMARY KEY (id)
);

INSERT INTO produto(descricao,unidade,quantidade,preco_real,preco_dolar)
VALUES('Acai','Litro',10,20.50,0.0);

INSERT INTO produto(descricao,unidade,quantidade,preco_real,preco_dolar)
VALUES('Tucuma','kg',20,15.30,0.0);

INSERT INTO produto(descricao,unidade,quantidade,preco_real,preco_dolar)
VALUES('Tapioca','Und',10,5.60,0.0);

-- Inserir registros na tabela Produto
