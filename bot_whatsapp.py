#importando biblioteca que lerá a planilha
import openpyxl
# importando o formatador da mensagem para maneira desejada
from urllib.parse import quote
#importando a biblioteca que irá abrir o browser no link  desejado
import webbrowser
#importando o sleep para ter tempo entre as operções para que não ocorra erros 
from time import sleep
#importando biblioteca que vai reconhecer o botão de enviar mensagem baseado na imagem do botão de enviar
import pyautogui

#Lendo a planilha e guardando dentro de "workbook"
workbook = openpyxl.load_workbook('Planilha.xlsx')

#Indicando a pagina da planilha que será iniciado
planilha_contato = workbook['Planilha1']

#Ciclo de repetição de todas as linhas da pagina da planilha selecionada
for linha in planilha_contato.iter_rows(min_row=2):
    #Carregando valores das linhas para nome, telefone e data
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value

    #mensagem personalizada usando nome e data informadas na planilha
    mensagem = f'Oi {nome} este é um teste de um bot para whatsapp, não precisa responder a essa mensagem, que foi enviada na data {data.strftime('%d/%m/%Y')}.'
    try:
        #O link pegará o telefone da planilha e assim mandará mensagem para o usuario selecionado
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        #Abrindo o browser no link e no numero carregado da planilha
        webbrowser.open(link_mensagem_whatsapp)
        sleep(35)
        #mapeando o centro da imagem
        botao_enviar = pyautogui.locateCenterOnScreen('botão_enviar.PNG')
        sleep(25)
        #Selecionando o eixo X e Y  onde deve ocorrer o click na imagem
        pyautogui.click(botao_enviar[0],botao_enviar[1])
        sleep(20)
        pyautogui.hotkey('ctrl','w')
        sleep(20)
    #tratamento de erro     
    except:
        # mensagem no terminal com nome da pessoa que não deu certo
        print(f'Não foi possivel enviar mensagem para {nome}')
        #escreve um arquivo csv com nome e telefone de todos que deram errado
        with open('erros.csv','a', newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone} ')


