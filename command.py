import os
from datetime import date

import pandas as pd


def criar_pasta(pasta):
    conf = input("Deseja criar a pasta?(S/N): ")
    if conf == 'S':
        os.chdir("C:/Users/gabri/Desktop")
        os.mkdir(pasta)
        print("Pasta Criada com Sucesso!")
    else:
        print("Operação Cancelada")


def criar_docx(arq):
    conf = input("Deseja criar o arquivo?(S/N): ")
    if conf == 'S':
        with open('C:/Users/gabri/OneDrive/Fatec/' + arq + '.docx', "w") as file:
            file.write('Testando')
        print("Arquivo Criado com Sucesso!")
    else:
        print("Operação Cancelada")


def criar_projeto(project):
    conf = input('Deseja criar o projeto?(S/N):')
    if conf == 'S':
        os.chdir("C:/Users/gabri/Desktop")
        os.mkdir(project)

        os.chdir("C:/Users/gabri/Desktop/" + project)
        os.mkdir('html_files')
        os.mkdir('css_files')
        os.mkdir('img_files')
        os.mkdir('js_files')
        with open('C:/Users/gabri/Desktop/' + project + '/html_files/index.html', 'w') as html:
            html.write('<html>')
            html.write('<head>')
            html.write('<title>Document</title>')
            html.write('</head>')
            html.write('<body>')
            html.write('<p>Projeto Teste<p>')
            html.write('</body>')
            html.write('</html')

        with open('C:/Users/gabri/Desktop/' + project + '/css_files/index.css', 'w') as css:
            css.write('/* Arquivo CSS */')

        with open('C:/Users/gabri/Desktop/' + project + '/js_files/index.js', 'w') as js:
            js.write('// Arquivo JS')

        print('Projeto Criado com Sucesso!')
    else:
        print('Operação Cancelada!')


def relatorio_bolsa(ind):
    url = 'http://bvmf.bmfbovespa.com.br/indices/ResumoCarteira' \
          'Teorica.aspx?Indice={}&idioma=pt-br'.format(ind.upper())
    dadosB3 = pd.read_html(url, decimal=',', thousands='.', index_col="Código")[0][:-1]
    data = date.today()
    name = str(data) + '-' + ind
    print(dadosB3)
    dadosB3.to_excel(r'C:\Users\gabri\Desktop\CarteiraTeorica\{}.xlsx'.format(name))
    print('Arquivo gerado com sucesso')


print('=======================================')
print('         Auxiliar do Gabriel           ')
print('=======================================')
print('1 - Criar Pasta no Desktop             ')
print('2 - Criar docx Facul                   ')
print('3 - Projeto de Site                    ')
print('4 - Carteira Teórica                   ')
print('=======================================')
op = int(input('Escolha uma opção: '))

if op == 1:
    nome = input('Informe o nome da pasta: ')
    criar_pasta(nome)
elif op == 2:
    nome = input('Informe o nome do arquivo: ')
    criar_docx(nome)
elif op == 3:
    nome = input('Informe o nome do projeto: ')
    criar_projeto(nome)
elif op == 4:
    indice = input('Informe o indice: ')
    relatorio_bolsa(indice)
else:
    print('Opção Inválida!')
