import glob
import os

dwg = glob.glob(input('Escolha a extensão que deseja filtrar: '))  # Filtra apenas os arquivos com a extensão inserida e salva numa lista
# with open() as data:  # Abre o arquivo sobs.txt contendo parte da string a ser comparada
for root, dirs, files in os.walk('.'):
    for file in files:
        if file.endswith(".txt"):
            with open(file) as data:
                datalines = (line.rstrip('\r\n') for line in data)  # Separa cada sob como uma nova linha no arquivo txt
                for line in datalines:  # Faz uma varretura linha a linha no arquivo txt
                    if not any(line in s for s in dwg):  # Verifica se há alguma string no arquivo txt que não está contida na
                        # variável dwg (lista)
                        print(line + ' não encontrado.')  # Imprime a mensagem de arquivo não encontrado, caso positivo.
input()