import pyodbc

class Arquivos():
    def __init__(self, id: int, nome_arquivo: str, criacao:str):
        self.id = id
        self.nome_arquivo = nome_arquivo
        self.criacao = criacao 

    def get_id(self):
        return self.id
    def get_nome_arquivo(self):
        return self.nome_arquivo
    def get_criacao(self):
        return self.criacao
    
if __name__ == '__main__':
    arq1 = Arquivos(1,'arquivo_teste.txt', '20023-01-31')

    print(arq1.id, arq1.nome_arquivo, arq1.criacao)