import time
import pyautogui
import pyscreeze
import pygetwindow as gw
import datetime
from datetime import date
import csv
import ctypes
import subprocess
import os

# OBS: Para esse programa funcionar, se a variável realizarLogin estiver como True (O que vai estar se for pra lançar como rotina automática), o Protheus precisa estar fechado
# ou na janela de informar nome e senha de usuário. Se a variável realizarLogin estiver como False, então precisa estar em cima de um CTE não classificado pra funcionar
# OBS: Se for para rodar como tarefa agendada do Windows, colocar os arquivos .png e .txt necessários na pasta System32

# Use "python -m PyInstaller classificar.py" no terminal para gerar o arquivo executável

def dataAutomatica(flag): # Usada na determinação da data de vencimento automática, com base na data de hoje
    hoje = date.today() 
    dia = int(hoje.strftime("%d"))
    diaSemana = date.weekday(hoje) # Retorna um número: Segunda = 0, Terça = 1, Quarta = 2, Quinta = 3, Sexta - 4, Sábado = 5, Domingo = 6
    if flag == 1: # Caso TMB e CDLOG
        if dia < 16: # Do dia 1 até 15, vencimento é dia 30
            mes = int(hoje.strftime("%m"))
            ano = int(hoje.strftime("%Y"))
            return str("30/"+str(mes).zfill(2)+"/"+str(ano))
        else: # Do dia 16 até 31, o vencimento é dia 15 do mês seguinte
            mes = int(hoje.strftime("%m")) + 1
            if mes > 12: # Se passou de dezembro, volta pra Janeiro e muda o ano
                ano = int(hoje.strftime("%Y")) + 1
                mes = 1
            else:
                ano = int(hoje.strftime("%Y"))
            return str("15/"+str(mes).zfill(2)+"/"+str(ano))
    elif flag == 2: # Caso Trans Gaúcho
        if dia < 11: # Do dia 1 até 10, o vencimento é dia 20
            mes = int(hoje.strftime("%m"))
            ano = int(hoje.strftime("%Y"))
            return str("20/"+str(mes).zfill(2)+"/"+str(ano))
        elif dia < 21: # Do dia 11 até 20, o vencimento é dia 30
            mes = int(hoje.strftime("%m"))
            ano = int(hoje.strftime("%Y"))
            return str("30/"+str(mes).zfill(2)+"/"+str(ano))
        else: # Do dia 21 até 31, o vencimento é dia 10 do mês seguinte
            mes = int(hoje.strftime("%m")) + 1
            if mes > 12: # Se passou de dezembro, volta pra Janeiro e muda o ano
                ano = int(hoje.strftime("%Y")) + 1
                mes = 1
            else:
                ano = int(hoje.strftime("%Y"))
            return str("10/"+str(mes).zfill(2)+"/"+str(ano))
    elif flag == 3: # Caso "Outros"
        if diaSemana < 2: # Se for Segunda ou Terça, o vencimento será a próxima Sexta-feira.
            hoje = (hoje+datetime.timedelta(days=(4-diaSemana)))
        else: # Se for Quarta até Domingo, o vencimento será a próxima Quarta-feira. 
            hoje = (hoje+datetime.timedelta(days=(9-diaSemana)))
        return str(hoje.strftime("%d/%m/%Y"))

def pegarNumeroCTEs():
    # Ler quantidade de linhas de CTEs da planilha (-1, porque a primeira é cabeçalho)
    # Se tiver a planilha do número de CTEs não classificados, irá pegar o número de CTEs da planilha, caso contrário, irá pegar 25 como padrão
    try:
        reader = csv.reader(open("Contabilidade - CTES NÃO CLASSIFICADAS.csv"))
        numeroCTE = int(len(list(reader)) - 1)
    except:
        numeroCTE = 25
    print("Número de CTEs a ser lançado:",numeroCTE)
    return int(numeroCTE)

# Retorna se o Capslock tá ligado ou não
def capslock_status():
    return True if ctypes.WinDLL("User32.dll").GetKeyState(0x14) else False

# Liga o Capslock se estiver desligado
def liga_capslock():
  if capslock_status() == False:
      pyautogui.press('capslock')

# Desliga o Capslock se estiver ligado
def desliga_capslock():
  if capslock_status() == True:
      pyautogui.press('capslock')

# Função para localizar linhas na tela
def localizar_linhas():
    region = (0, 260, 110, 350)
    linhas = [
        (diretorioAtual + "\\linha1.png", 1),
        (diretorioAtual + "\\linha2.png", 2),
        (diretorioAtual + "\\linha3.png", 3),
        (diretorioAtual + "\\linha4.png", 4),
        (diretorioAtual + "\\linha5.png", 5),
        (diretorioAtual + "\\linha6.png", 6), # Única não testada
        (diretorioAtual + "\\linha7.png", 7),
        (diretorioAtual + "\\linha8.png", 8),
        (diretorioAtual + "\\linha9.png", 9),
        (diretorioAtual + "\\linha10.png", 10),
        (diretorioAtual + "\\linha11.png", 11),
        (diretorioAtual + "\\linha12.png", 12),
        (diretorioAtual + "\\linha13.png", 13),
        (diretorioAtual + "\\linha14.png", 14),
        (diretorioAtual + "\\linha15.png", 15),
        (diretorioAtual + "\\linha17.png", 17),
        (diretorioAtual + "\\linha21.png", 21)
    ]

    for imagem, numero in linhas:
        try:
            pyautogui.locateOnScreen(imagem, region=region)
            print(f'Tem {numero} linhas')
            return numero
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            print(f'Não tem {numero} linhas')

    return 0  # Se nenhuma linha foi encontrada

# Pega o dia útil anterior
def dia_util_anterior(data):
    data -= datetime.timedelta(days=1)
    while data.weekday() > 4: # Mon-Fri are 0-4
        data -= datetime.timedelta(days=1)
    return data

def rodarArquivoBatch(caminhoArquivo): # Também serve para abrir a planilha do Excel
    try:
        filepath = (caminhoArquivo) # Abre o Protheus, rodando o executável na pasta smartclient_64

        p = subprocess.Popen(filepath, shell=True)
        #p = subprocess.Popen(filepath, shell=True, stdout = subprocess.PIPE)
        #stdout, stderr = p.communicate()
    except:
        print("Erro de arquivo Batch")

    #if (p.returncode) != 0:
        #print(p.returncode)
        #exit() # is 0 if success

def registrarLog(mensagem):
    file_log = open(diretorioAtual + "\\mensagem_final_classificar.txt", "a") # Usado para manter um log de mensagens finais, pra ver qual erro pode ter ocorrido durante a rotina automática
    file_log.write(mensagem+'\n')
    file_log.close()

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------

pyautogui.useImageNotFoundException()
msg = "------------------ Rotina de Classificação automática iniciada ------------------"
# OBS: Esse programa foi feito usando o Protheus em janela maximizada e com tela única de resolução 1920x1080
# Failsafe: Em caso de emergência, jogue o mouse para o canto da tela para forçar o programa a parar (tente mais de uma vez se não parar)
print(msg)
diretorioAtual = str(os.getcwd())
print("Diretório atual:"+diretorioAtual)
im = pyautogui.screenshot(diretorioAtual + '\\fornecedor_atual.png',region=(1,1,500,500)) # Para garantir que o primeiro CTE NUNCA vai cair no caso de "mesmo fornecedor"
im = pyautogui.screenshot(diretorioAtual + '\\estado_atual.png',region=(1,1,300,300))
im = pyautogui.screenshot(diretorioAtual + "\\numeroCTE_atual.png", region=(1,1,300,300))
#pyautogui.hotkey('alt', 'tab')

'''
time.sleep(5)   # Usada para mostrar a posição atual do mouse
x, y = pyautogui.position()
print(x,y)
'''
# OBS: Importante! Manter o Capslock ligado antes de rodar o programa (idealmente)
# e já tem que estar em cima de um CTE (não pode estar aberto no Classificar) se rotinaAutomatica = False
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------
# Dados/Parâmetros:
# Configurações iniciais para a rotina automática:
file = open(diretorioAtual + "\\Configurações_Iniciais.txt", "r") # Abrindo o arquivo Configurações_Iniciais.txt para ler as seguintes variáveis

usuario = (str(file.readline())).rstrip()
senha = (str(file.readline())).rstrip()
caminhoProtheus = (str(file.readline())).rstrip() # Guarda o diretório onde está localizado o Protheus (Padrão = "C:\\smartclient_64\\smartclient.exe")
numeroPageDown = int(file.readline()) # Número de vezes que vai apertar Page Down antes de começar a classificar (Padrão = 0)
if ((str(file.readline())).rstrip() == "Sim"): # Lê o arquivo TXT, se for sim, seta o valor como True
    rotinaAutomatica = True # Esse código está rodando automática ou é um usuário rodando? (False se manual, True se automática, default=True)
else:
    rotinaAutomatica = False
# Se a rotina está rodando sozinha, o programa deve sempre realizar o Login no Protheus
if rotinaAutomatica:
    realizarLogin = True # Vai realizar o Login no Protheus ou vai começar a lançar o CTE que estiver selecionado? (True se sim, False se não, Default = True)
else:
    if ((str(file.readline())).rstrip() == "Sim"): # Lê o arquivo TXT, se for sim, seta o valor como True
        realizarLogin = True
    else:
        realizarLogin = False 
    
#realizarLogin = False # Temporário, apenas serve como teste durante o desenvolvimento
#rotinaAutomatica = False # Colocar essa variável como False para fazer com que o programa comece a classificar no CTE que está atualmente selecionado

file.close() # Fechando o arquivo TXT

if rotinaAutomatica:
    manual = False # Se a rotina está rodando sozinha, os dados devem ser sempre puxados automaticamente também
else:
    if (pyautogui.confirm(text='Bem vindo, deseja informa os dados padrão manualmente?', title='Manual?', buttons=['Sim', 'Não'])) == 'Sim':
        manual = True # Dados/Parâmetros informados aqui no código ou manual, quando rodar o programa? (False se não, True se sim)
    else:
        manual = False

if manual: # Se for pra informar manualmente
    pyautogui.alert(text='Por favor, informe os dados padrões dos CTEs', title='Início', button='OK')
    # Esses serão os dados usados quando o fornecedor do CTE for Outros
    # Data de vencimento e número de CTEs são inalterados, independentemente do fornecedor
    try:
        numeroCTEs = int(pyautogui.prompt(text='Informe quantos CTEs serão classificados: ', title='Número de CTEs' , default='1'))
    except:
        pyautogui.alert('Erro! Número de CTEs inválido!') #'This displays some text with an OK button.'
        exit()  #Encerra o programa se não encontrar a janela do Protheus   
    dataVencimento = pyautogui.prompt(text='Informe qual é a Data de Vencimento: ', title='Data de Vencimento' , default='DD/MM/YYYY')
    dataVencimentoPadrao = dataVencimento #As variáveis "Padrão" guardam o valor default para caso não seja um dos três fornecedores principais
    data15_30 = pyautogui.prompt(text='Informe qual é a Data de Vencimento para TMB e CDLog (15 ou 30): ', title='Data de Vencimento 15 ou 30' , default='DD/MM/YYYY')
    data10_20_30 = pyautogui.prompt(text='Informe qual é a Data de Vencimento para Trans Gaúcho (10, 20 ou 30): ', title='Data de Vencimento 10, 20 ou 30' , default='DD/MM/YYYY')
    if (pyautogui.confirm(text='Informe se tem ICMS:', title='ICMS?', buttons=['Sim', 'Não'])) == 'Sim':
        temICMS = True
    else:
        temICMS = False
    temICMSPadrao = temICMS
    if (pyautogui.confirm(text='Informe se é preciso alterar a alíquota de ICMS:', title='Alíquota de ICMS?', buttons=['Sim', 'Não'])) == 'Sim':
        aliquotaDiferente = True
    else:
        aliquotaDiferente = False
    aliquotaDiferentePadrao = aliquotaDiferente
    if aliquotaDiferente:
        aliquota = str(pyautogui.prompt(text='Informe qual é a Alíquota de ICMS: ', title='Alíquota de ICMS' , default='12'))
        aliquotaPadrao = aliquota
    if (pyautogui.confirm(text='Informe se é preciso alterar a situação tributária:', title='Situação Tributária?', buttons=['Sim', 'Não'])) == 'Sim':
        tesDiferente = True
    else:
        tesDiferente = False
    tesDiferentePadrao = tesDiferente
    if tesDiferente:
        tes = str(pyautogui.prompt(text='Informe qual é a situação tributária: ', title='Situação Tributária' , default='040'))
        tesPadrao = tes
    usuario = pyautogui.prompt(text='Informe qual é o usuário a ser usado: ', title='Usuário' , default='User')
    senha = pyautogui.prompt(text='Informe qual é a senha do usuário: ', title='Senha' , default='Password')
else: # Se for pra ser automático
    # Dados para login no Protheus
    usuario = "Nobre03"     # Conta do usuário
    senha = "NObres.2021."  # Senha do usuário
    # Dados para lançamento
    numeroCTEs = 100 # Número de loops que o programa vai realizar (ou seja, quantos CTEs serão classificados se der certo - Padrão: 25)
    dataVencimento = dataAutomatica(3) # Data a ser informada no vencimento, puxada automaticamente com base na data de hoje
    data15_30 = dataAutomatica(1) # Data a ser usada como vencimento para os CTes da TMB e da CDLog
    data10_20_30 = dataAutomatica(2) # Data a ser usada como vencimento para os CTes da Trans Gaúcho
    dataVencimentoPadrao = dataVencimento #As variáveis "Padrão" guardam o valor default para caso não seja um dos três fornecedores principais
    temICMS = True # Informe aqui se os CTEs tem ICMS ou não, (True se sim, False se não, Padrão = True)
    temICMSPadrao = temICMS
    aliquotaDiferente = False # Informe aqui se os CTEs precisam ter a alíquota alterada ou não (False se não, True se sim, Padrão = False)
    aliquotaDiferentePadrao = aliquotaDiferente
    aliquota = '12' # Se precisa alterar a alíquota, informe aqui qual é a alíquota correta
    aliquotaPadrao = aliquota
    tesDiferente = False # Informe aqui se precisa colocar a situação tributária manualmente nos CTEs (False se não, True se sim, Padrão = False, é raro precisar disso)
    tesDiferentePadrao = tesDiferente
    tes = '040' # Se precisa alterar a situação tributária, informe aqui qual é a situação tributária correta
    tesPadrao = tes
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------

# Lista de erros que podem acontecer nos CTEs para atrapalhar o programa: Base de cálculo errada, valor de ICMS errado, valor total errado, sem saldo liberado, 
# alíquota de icms errada, valor unitário errado, mais linhas do que o previsto até agora, ter antecipação, SEFAZ pode estar fora do ar, situação tributária errada,
# Sistema ou a Sefaz podem estar mais lentos que o normal, o sistema pode não ter puxado o centro de custo do CTE, o programa pode ter sido rápido demais em algum ponto,
# Pode entrar em Loop "infinito" no último CTE (teoricamente resolvido)

try:
    protheusWindow = gw.getWindowsWithTitle('TOTVS Manufatura')[0]
except:
    try:
        if realizarLogin: # Se realizarLogin for True, tenta abrir o Protheus sozinho
            print('Protheus não está aberto!') # Ainda mantenho o Print para ter o registro do erro
            rodarArquivoBatch(caminhoProtheus) # Abrindo o Protheus (Diretório = "C:\smartclient_64\smartclient.exe")
            time.sleep(10) # (Arquivo .bat = str(os.getcwd() + ("\\Abre_Protheus.bat")) )
            try:
                protheusWindow = gw.getWindowsWithTitle('Initial Par')[0]
            except:
                protheusWindow = gw.getWindowsWithTitle('metros Iniciais')[0]
            protheusWindow.activate()
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(12)
        else:
            raise Exception("Não está programado para abrir o Protheus")
    except:
        print('Erro! Protheus não foi aberto!') # Ainda mantenho o Print para ter o registro do erro
        if rotinaAutomatica:
            registrarLog('Erro! Protheus não foi aberto!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
        else:  
            pyautogui.alert('Erro! Protheus não foi aberto!') #'This displays some text with an OK button.'
        exit()  #Encerra o programa se não encontrar a janela do FR

try:
    protheusWindow = gw.getWindowsWithTitle('TOTVS Manufatura')[0]
except:
    print('Erro! Protheus não encontrado!') # Ainda mantenho o Print para ter o registro do erro
    if rotinaAutomatica:
        registrarLog('Erro! Protheus não encontrado!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
    else:
        pyautogui.alert('Erro! Protheus não encontrado!') #'This displays some text with an OK button.'
    exit()  #Encerra o programa se não encontrar a janela do FR
protheusWindow.activate() # Abre a janela do Protheus
protheusWindow.maximize() # Maximiza o Protheus

time.sleep(1)
# ////////// ---------------------------------------------------------- Início da Lógica de Login ---------------------------------------------------------- //////////

if realizarLogin:
    time.sleep(1)
    try:
        location = pyautogui.locateOnScreen(diretorioAtual + "\\tela_login.png",region=(802,269,296,150)) # Verifica se o Protheus está na tela de informar usuário e senha
        #print(location)
        # Se não achar significa que não está na tela de Login
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        try:
            time.sleep(20) # Espera mais vinte segundos e testa novamente
            location = pyautogui.locateOnScreen(diretorioAtual + "\\tela_login.png",region=(802,269,296,150)) # Verifica se o Protheus está na tela de informar usuário e senha
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            if rotinaAutomatica:
                registrarLog('Erro! Não está na tela de Login!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                protheusWindow.close()
                time.sleep(2)
                pyautogui.press('enter')
            else:
                pyautogui.alert('Erro! Não está na tela de Login!') #'This displays some text with an OK button.'
            print('Erro! Não está na tela de Login!') # Ainda mantenho o Print para ter o registro do erro
            exit()  #Encerra o programa se não encontrar a janela do FR

    # Digita usuário e senha para logar no Protheus
    desliga_capslock()
    pyautogui.click(938, 493) # Normalmente, quando abre o Protheus a barra de usuário já está selecionada e então é só digitar, mas estou fazendo clicar nela só pra garantir
    pyautogui.write(usuario)
    time.sleep(1) # Espera um pouco o sistema responder
    pyautogui.press('tab')
    pyautogui.write(senha)
    time.sleep(3) # Espera um pouco o sistema responder
    pyautogui.press('enter')
    pyautogui.press('capslock') # Reativando o Capslock

    print("Verificando se o Login deu certo...")
    time.sleep(5) # Espera um pouco o sistema responder
    try:
        pyautogui.locateOnScreen(diretorioAtual + "\\tela_login_OK.png", region=(820,730,302,84)) # Verifica se o login deu certo e não está mais na tela de informar usuário e senha
        # Se achar significa que não deu erro e foi pra tela seguinte
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        try:
            time.sleep(15) # Espera mais quinze segundos para garantir que vai logar
            pyautogui.locateOnScreen(diretorioAtual + "\\tela_login_OK.png", region=(820,730,302,84))
        except:
            if rotinaAutomatica:
                registrarLog('Erro! Login inválido!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                protheusWindow.close()
                time.sleep(2)
                pyautogui.press('enter')
            else:
                pyautogui.alert('Erro! Login inválido!') #'This displays some text with an OK button.'
            print('Erro! Login inválido!') # Ainda mantenho o Print para ter o registro do erro
            exit()  #Encerra o programa se não encontrar a janela do FR
    else:
        print("Login realizado no Protheus com o usuário", usuario)

    # Entra no ambiente certo para classificar CTEs
    for j in range(1,3):
        pyautogui.press('tab') # Vai pra próxima barra
    pyautogui.write("2") # Digita o número do Ambiente do Compras
    time.sleep(1) # Espera um pouco o sistema responder
    for j in range(1,4):
        pyautogui.press('tab') # Aperta TAB três vezes para ir pro botão Entrar
    pyautogui.press('enter') # Entra no sistema principal

    # Confere se abriu certo o Protheus
    time.sleep(10)
    try:
        pyautogui.locateOnScreen(diretorioAtual + "\\tela_inicial_Protheus.png",region=(0,95,183,167))
        # Se não achar significa que não está na tela de Documento Entrada
    except:
        try:
            time.sleep(10) # Espera mais dez segundos
            pyautogui.locateOnScreen(diretorioAtual + "\\tela_inicial_Protheus.png",region=(0,95,183,167))
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            if rotinaAutomatica:
                registrarLog('Erro! Não está na tela do Compras!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                protheusWindow.close()
                time.sleep(2)
                pyautogui.press('left')
                time.sleep(1)
                pyautogui.press('enter')
            else:
                pyautogui.alert('Erro! Não está na tela do Compras!') #'This displays some text with an OK button.'
            print('Erro! Não está na tela do Compras!') # Ainda mantenho o Print para ter o registro do erro
            exit()  #Encerra o programa se não encontrar a janela do FR
    else:
        print("Acessando Documento Entrada...")

    time.sleep(1)
    pyautogui.click(67, 461) # Clica em Atualizações
    time.sleep(1) # Espera um pouco o sistema responder
    pyautogui.click(64, 640) # Clica em Movimentos
    time.sleep(1) # Espera um pouco o sistema responder
    pyautogui.click(105, 699) # Clica em Documento Entrada
    time.sleep(1)
    try:
        pyautogui.locateOnScreen(diretorioAtual + "\\AbrindoDocEntrada.png") # Se não tiver dado certo pra entrar em Documento de Entrada, tenta mais uma vez
    except:
        time.sleep(1)
        pyautogui.click(67, 461) # Clica em Atualizações
        time.sleep(1) # Espera um pouco o sistema responder
        pyautogui.click(64, 640) # Clica em Movimentos
        time.sleep(1) # Espera um pouco o sistema responder
        pyautogui.click(105, 699) # Clica em Documento Entrada
        time.sleep(1)

    # Se for dia primeiro do mês, tem que voltar pro dia útil anterior por causa do fechamento
    data_atual = date.today()
    dia_hoje = int(data_atual.strftime("%d"))
    if dia_hoje == 1:
        time.sleep(1) # Espera um pouco o sistema responder
        #dia_anterior = str(dia_util_anterior(data_atual).strftime("%d"))
        #mes_anterior = str(dia_util_anterior(data_atual).strftime("%m"))
        #ano_anterior = str(dia_util_anterior(data_atual).strftime("%Y"))
        #print(dia_anterior+mes_anterior+ano_anterior)
        pyautogui.write(str(dia_util_anterior(data_atual).strftime("%d/%m/%Y")))
        time.sleep(1) # Espera um pouco o sistema responder
        for j in range(1,4):
            pyautogui.press('enter') # Aperta Enter cinco vezes para ir pro botão Entrar
            time.sleep(1) # Espera um pouco só para garantir que não vai acontecer nenhum erro por ter sido rápido demais
    else:
        for j in range(1,5):
            pyautogui.press('enter') # Aperta Enter quatro vezes para ir pro botão Entrar
            time.sleep(1)

    # Filtrar para pegar apenas os CTEs pendentes de classificação
    time.sleep(20) # Espera sistema entrar em Documento Entrada (isso demora)
    try:
        pyautogui.locateOnScreen(diretorioAtual + "\\tela_CTE.png", region=(0,85,217,53))
        # Se não achar significa que não está na tela de Documento Entrada
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        try:
            time.sleep(60) # Espera 1 minuto pra garantir que vai entrar no sistema
            pyautogui.locateOnScreen(diretorioAtual + "\\tela_CTE.png", region=(0,85,217,53))
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            if rotinaAutomatica:
                registrarLog('Erro! Não está na tela de Documento Entrada!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                protheusWindow.close()
                time.sleep(2)
                pyautogui.press('left')
                time.sleep(1)
                pyautogui.press('enter')
            else:
                pyautogui.alert('Erro! Não está na tela de Documento Entrada!') #'This displays some text with an OK button.'
            print('Erro! Não está na tela de Documento Entrada!') # Ainda mantenho o Print para ter o registro do erro
            exit()  #Encerra o programa se não encontrar a janela do FR
        else:
            print("Filtrando CTEs...")
    else:
        print("Filtrando CTEs...")

    pyautogui.click(1896, 154) # Clica em "Filtrar"
    time.sleep(3)
    try:
        # Filtro usado: Espec.Docum. Igual a 'CTE' e Status Igual a '%F1_STATUS4%'
        location = pyautogui.locateOnScreen(diretorioAtual + "\\filtro_cte.png", region=(629,276,657,572)) # Achar a localização do filtro certo
        pyautogui.click(pyautogui.center(location))
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        if rotinaAutomatica:
            registrarLog('Erro! O filtro não foi encontrado!') # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
            protheusWindow.close()
            time.sleep(2)
            pyautogui.press('left')
            time.sleep(1)
            pyautogui.press('enter')
        else:
            pyautogui.alert('Erro! O filtro não foi encontrado!') #'This displays some text with an OK button.'
        print('Erro! O filtro não foi encontrado!') # Ainda mantenho o Print parCa ter o registro do erro
        exit()  #Encerra o programa se não encontrar a janela do FR

    time.sleep(2)
    pyautogui.click(1171, 780) # Clica em "Aplicar Filtros Selecionados"
    time.sleep(1)

    for j in range(1,3):
            pyautogui.press('enter') # Aperta Enter duas vezes para ajustar o filtro e e finalmente filtrar os CTEs

    time.sleep(10) # Espera Filtrar
    print("CTEs filtrados")

# ////////// ------------------------------------------------------ Início da Lógica de Classificação ------------------------------------------------------ ////////// 
ctesLancados = 0 # Número de CTEs que foram classificados efetivamente
liga_capslock() # Ligando o capslock se estiver desligado
for j in range(1,numeroPageDown+1): # Aperta um certo número de vezes Page Down antes de começar a classificar os CTEs
    time.sleep(1)
    pyautogui.press('pagedown')

#Início do Loop principal
for i in range(numeroCTEs):
    try:
        time.sleep(1) # Espera mais um segundo pra evitar um bug que o "C" não é detectado porque foi digitado logo após o sistema liberar pra classificar outro CTE
        print('Verificando se está tudo certo...')
        #pyautogui.locateOnScreen("C:\\Users\\Contab\\Documents\\Lightshot\\OK2.png") #Absolute Directory Path
        pyautogui.locateOnScreen(diretorioAtual + "\\OK2.png") # Verifica se o Protheus está liberado para abrir um CTE
        #pyautogui.doubleClick(pyautogui.center(location))
        print('Abrindo CTE:',i+1)
        pyautogui.write('c')
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        time.sleep(5) #Espera mais cinco segundos pra ver se vai abrir
        try:
            pyautogui.locateOnScreen(diretorioAtual + "\\OK2.png")
            print('Abrindo CTE:',i+1)
            pyautogui.write('c')
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            try:
                time.sleep(5) #Espera mais cinco segundos pra ver se vai abrir
                pyautogui.locateOnScreen(diretorioAtual + "\\OK2.png")
                print('Abrindo CTE:',i+1)
                pyautogui.write('c')
            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                try:
                    time.sleep(60) #Espera mais um minuto pra ver se vai abrir
                    pyautogui.locateOnScreen(diretorioAtual + "\\OK2.png")
                    print('Abrindo CTE:',i+1)
                    pyautogui.write('c')
                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                    print('Erro (Não voltou pra tela original)! Foram lançados',ctesLancados,'CTEs!') # Ainda deixo o Print para ter o registro da quantidade de CTEs lançados
                    mensagemFinal = 'Erro (Não retornou pra tela original)! Foram lançados '+ str(ctesLancados) +' CTEs!' # Só pra indicar que o programa chegou ao fim
                    if rotinaAutomatica:
                        registrarLog(mensagemFinal) # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                        protheusWindow.close()
                        time.sleep(2)
                        pyautogui.press('left')
                        time.sleep(1)
                        pyautogui.press('enter')
                    else:
                        pyautogui.alert(mensagemFinal) #'This displays some text with an OK button.'
                    exit()  #Encerra o programa se algo der errado quando for tentar salvar o CTE

    time.sleep(4) # Esperando o CTE abrir (Padrão = 5)

    #CTE já deve estar aberto nesse ponto
    try:
        print('Verificando se o CTE foi aberto...') # Ver se o CTE está aberto (Se não tiver, é porque o sistema provavelmente está mais lento que o normal)
        pyautogui.locateOnScreen(diretorioAtual + "\\classificarAberto.png") # Isso demora mais ou menos 1 segundo
        print('CTe aberto')
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        time.sleep(5) #Padrão: (5), esperando mais um pouco pro sistema abrir o CTE
        try:
            print('Verificando pela segunda vez se o CTE foi aberto...') # Testando de novo pra ver se o CTE tá aberto agora depois de esperar 5 segundos
            pyautogui.locateOnScreen(diretorioAtual + "\\classificarAberto.png") # Isso demora mais ou menos 1 segundo
            print('CTe aberto')
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            try:
                time.sleep(30) #Padrão: (30), esperando mais um monte pro sistema abrir o CTE
                print('Verificando pela terceira vez se o CTE foi aberto...') # Testando de novo pra ver se o CTE tá aberto agora depois de esperar 5 segundos
                pyautogui.locateOnScreen(diretorioAtual + "\\classificarAberto.png") # Isso demora mais ou menos 1 segundo
                print('CTe aberto')
            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                print('Erro! O CTE não está aberto! Foram lançados',ctesLancados,'CTEs!') # Ainda mantenho o Print para ter o registro da quantidade de CTEs lançados
                mensagemFinal = 'Erro (CTE não foi aberto)! Foram lançados '+ str(ctesLancados) +' CTEs!' # Só pra indicar que o programa chegou ao fim
                if rotinaAutomatica:
                    registrarLog(mensagemFinal) # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                    protheusWindow.close()
                    time.sleep(2)
                    pyautogui.press('left')
                    time.sleep(1)
                    pyautogui.press('enter')
                else:
                    pyautogui.alert(mensagemFinal) #'This displays some text with an OK button.'
                exit()  #Encerra o programa se algo der errado quando for tentar abrir o CTE

    outrasLinhas = - 17 # O quanto o mouse vai para baixo para clicar nas outras linhas. Se for só uma linha, o valor dessa variável vai ser 0
    numeroLinhas = 0 # Se não achar nenhuma imagem de linha, o valor da variável continua zero

    #Identificador do número de linhas
    numeroLinhas = localizar_linhas()

    if numeroLinhas == 0:
        print('Erro de linhas!')
        mensagemFinal = 'Erro (Número de linhas não catalogada ainda)! Foram lançados '+ str(ctesLancados) +' CTEs!' # Só pra indicar que o programa chegou ao fim
        if rotinaAutomatica:
            registrarLog(mensagemFinal) # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
            protheusWindow.close()
            time.sleep(2)
            pyautogui.press('left')
            time.sleep(1)
            pyautogui.press('enter')
        else:
            pyautogui.alert(mensagemFinal) #'This displays some text with an OK button.'
        exit()  #Encerra o programa se algo der errado quando for tentar abrir o CTE

    for j in range(1,3):
        pyautogui.press('enter') # Vai pra primeira célula da primeira linha após abrir o CTE

    #-------------------------------------------------------------------- Loop das Linhas -------------------------------------------------------------------

    for i in range(0,numeroLinhas): #Loop das Linhas
        outrasLinhas += 17 # 17 é a diferença entre cada linha, então 17 é somado no início de cada loop de linha

        #pyautogui.doubleClick(1089, 308+outrasLinhas) # Clica em tipo de Operação
        if outrasLinhas != 0:
            pyautogui.press('down')
        for j in range(1,10):
            pyautogui.press('right') # Vai pra Tipo de Operação

        pyautogui.write('fr') # Informa que o tipo de operação é FR (Frete Sobre Venda Mercadoria)
        try: # Tentativa de detectar se o Capslock tá ativado ou não
            #location pegava a localização de uma imagem específica na tela, agora mudei pra usar as coordenadas exatas
            #location = pyautogui.locateOnScreen("C:\\Users\\Contab\\Documents\\Lightshot\\TpOpe.png")
            #pyautogui.doubleClick(pyautogui.center(location))
            time.sleep(1)
            pyautogui.locateOnScreen(diretorioAtual + "\\capslock.png", region=(1056, 697, 160, 79)) #Se localizar essa imagem, indica que o capslock tá desativado pois FR deu erro
            pyautogui.press('enter')
            print('FR - Ligando Capslock ... Por favor, já deixar o Capslock ativado antes de rodar o programa')
            pyautogui.press('capslock')
            pyautogui.write('fr')
            time.sleep(1)
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            print('FR')

        # Identificador de Fornecedor
        flagOutros = False  # Indica que o fornecedor caiu no caso "Outros" e deve testar mais se é de Simples Nacional ou não
        if outrasLinhas == 0: # Só vai ver o fornecedor na primeira linha pois isso precisa ocorrer depois de colocar FR e a TES aparecer
            try:
                print('Vendo fornecedor') # region = (610,225, 100, 30) - Região limitada onde as imagens dos fornecedores são procuradas
                pyautogui.locateOnScreen(diretorioAtual + "\\fornecedor_atual.png",region=(610,225, 100, 30))
                pyautogui.locateOnScreen(diretorioAtual + "\\estado_atual.png",region=(1712,235,60,26))
                print('Mesmo fornecedor')
            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                print('Fornecedor novo') #Se não encontrar o mesmo CNPJ e o mesmo estado, é porque mudou o fornecedor
                im = pyautogui.screenshot(diretorioAtual + '\\fornecedor_atual.png',region=(610,225, 100, 30)) # Atualiza a print do Fornecedor
                im = pyautogui.screenshot(diretorioAtual + '\\estado_atual.png',region=(1712,235,60,26)) # Atualiza a print do Estado
                im = pyautogui.screenshot(diretorioAtual + "\\numeroCTE_atual.png", region=(1353,190,95,32)) # Atualiza a print do número de CTE porque logicamente é um CTE novo
                try:
                    pyautogui.locateOnScreen(diretorioAtual + "\\fornecedor_TMB.png",region=(600,215, 120, 50))
                    print('Fornecedor - Transportadora Minas Brasil')
                    temICMS = True  
                    aliquotaDiferente = False
                    tesDiferente = False
                    dataVencimento = data15_30
                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                    try:
                        pyautogui.locateOnScreen(diretorioAtual + "\\fornecedor_CDLOG.png",region=(600,215, 120, 50))
                        print('Fornecedor - CDLOG')
                        temICMS = True
                        aliquotaDiferente = True
                        aliquota = '17'
                        tesDiferente = False
                        dataVencimento = data15_30
                    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                        try:
                            pyautogui.locateOnScreen(diretorioAtual + "\\fornecedor_TG.png",region=(600,215, 120, 50))
                            print('Fornecedor - Trans Gaúcho')
                            temICMS = False
                            aliquotaDiferente = False
                            tesDiferente = False
                            dataVencimento = data10_20_30
                        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                            try:
                                pyautogui.locateOnScreen(diretorioAtual + "\\fornecedor_ZANOTELLI.png",region=(600,215, 120, 50))
                                pyautogui.locateOnScreen(diretorioAtual + "\\SC.png",region=(1702,225,80,46))
                                print('Fornecedor - Zanotelli Transporte e Logística (SC)') 
                                temICMS = True
                                aliquotaDiferente = True
                                aliquota = '17'
                                tesDiferente = False
                                dataVencimento = dataVencimentoPadrao
                            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                try:
                                    pyautogui.locateOnScreen(diretorioAtual + "\\PR.png",region=(1702,225,80,46))
                                    print('Fornecedor - Outros (PR)') # CTEs do Paraná tem isenção de ICMS
                                    temICMS = False
                                    aliquotaDiferente = False
                                    tesDiferente = False
                                    dataVencimento = dataVencimentoPadrao
                                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                    try:
                                        time.sleep(1) # Espera aparecer a TES
                                        pyautogui.locateOnScreen(diretorioAtual + "\\TES122SIMPLES2.png",region=(1121,230,110,90))
                                        print('Fornecedor - Simples Nacional 122') # Se tiver a TES 122, é do Simples Nacional
                                        temICMS = False # Não tá detectando essa imagem
                                        aliquotaDiferente = False
                                        tesDiferente = False
                                        dataVencimento = dataVencimentoPadrao
                                    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                        try:
                                            pyautogui.locateOnScreen(diretorioAtual + "\\TES292SIMPLES.png",region=(1121,230,110,90))
                                            print('Fornecedor - Simples Nacional 292') # Se tiver a TES 292, é do Simples Nacional
                                            temICMS = False # Não tá detectando essa imagem
                                            aliquotaDiferente = False
                                            tesDiferente = False
                                            dataVencimento = dataVencimentoPadrao
                                        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                            try:
                                                pyautogui.locateOnScreen(diretorioAtual + "\\TES156SIMPLES.png",region=(1121,230,110,90))
                                                print('Fornecedor - Simples Nacional 156') # Se tiver a TES 156, é do Simples Nacional
                                                temICMS = False # Não tá detectando essa imagem
                                                aliquotaDiferente = False
                                                tesDiferente = False
                                                dataVencimento = dataVencimentoPadrao
                                            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                                #print('Fornecedor - Outros')
                                                flagOutros = True
                                                temICMS = temICMSPadrao
                                                aliquotaDiferente = aliquotaDiferentePadrao
                                                tesDiferente = tesDiferentePadrao
                                                if aliquotaDiferente:
                                                    aliquota = aliquotaPadrao
                                                if tesDiferente:
                                                    tes = tesPadrao
                                                dataVencimento = dataVencimentoPadrao
                    if flagOutros: # Testes extra pra confirmar se é "Outros" ou se é Simples Nacional
                        try:
                            pyautogui.locateOnScreen(diretorioAtual + "\\TES157SIMPLES.png",region=(1121,230,110,90))
                            print('Fornecedor - Simples Nacional 157') # Se tiver a TES 157, é do Simples Nacional
                            temICMS = False # Não tá detectando essa imagem
                            aliquotaDiferente = False
                            tesDiferente = False
                            dataVencimento = dataVencimentoPadrao
                            flagOutros = False
                        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                            try:
                                pyautogui.locateOnScreen(diretorioAtual + "\\TES053SIMPLES.png",region=(1121,230,110,90))
                                print('Fornecedor - Simples Nacional 053') # Se tiver a TES 053, é do Simples Nacional
                                temICMS = False # Não tá detectando essa imagem
                                aliquotaDiferente = False
                                tesDiferente = False
                                dataVencimento = dataVencimentoPadrao
                                flagOutros = False
                            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                try:
                                    pyautogui.locateOnScreen(diretorioAtual + "\\TES120SIMPLES.png",region=(1121,230,110,90))
                                    print('Fornecedor - Simples Nacional 120') # Se tiver a TES 120, é do Simples Nacional
                                    temICMS = False # Não tá detectando essa imagem
                                    aliquotaDiferente = False
                                    tesDiferente = False
                                    dataVencimento = dataVencimentoPadrao
                                    flagOutros = False
                                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                    try:
                                        pyautogui.locateOnScreen(diretorioAtual + "\\TES164SIMPLES.png",region=(1121,230,110,90))
                                        print('Fornecedor - Simples Nacional 164') # Se tiver a TES 164, é do Simples Nacional
                                        temICMS = False # Não tá detectando essa imagem
                                        aliquotaDiferente = False
                                        tesDiferente = False
                                        dataVencimento = dataVencimentoPadrao
                                        flagOutros = False
                                    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                        try:
                                            pyautogui.locateOnScreen(diretorioAtual + "\\TES155SIMPLES.png",region=(1121,230,110,90))
                                            print('Fornecedor - Simples Nacional 155') # Se tiver a TES 155, é do Simples Nacional
                                            temICMS = False # Não tá detectando essa imagem
                                            aliquotaDiferente = False
                                            tesDiferente = False
                                            dataVencimento = dataVencimentoPadrao
                                            flagOutros = False
                                        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                            try:
                                                pyautogui.locateOnScreen(diretorioAtual + "\\IcmsZerado.png",region=(1269,274,86,42))
                                                print('Fornecedor - Simples Nacional?') # Última esperança de pegar o caso de Simples Nacional
                                                temICMS = True
                                                aliquotaDiferente = True
                                                aliquota = '12'
                                                tesDiferente = False
                                                dataVencimento = dataVencimentoPadrao
                                                flagOutros = False
                                            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                                try:
                                                    pyautogui.locateOnScreen(diretorioAtual + "\\MG.png",region=(1702,225,80,46))
                                                    print('Fornecedor - Outros (MG)') # CTEs de MG puxam alíquota 18 automaticamente
                                                    temICMS = True
                                                    aliquotaDiferente = True
                                                    aliquota = '12'
                                                    tesDiferente = False
                                                    dataVencimento = dataVencimentoPadrao
                                                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                                                    print('Fornecedor - Outros') # Agora está confirmado que são outros o fornecedor
                                                    temICMS = temICMSPadrao
                                                    aliquotaDiferente = aliquotaDiferentePadrao
                                                    tesDiferente = tesDiferentePadrao
                                                    if aliquotaDiferente:
                                                        aliquota = aliquotaPadrao
                                                    if tesDiferente:
                                                        tes = tesPadrao
                                                    dataVencimento = dataVencimentoPadrao
            else: # Se o fornecedor for o mesmo, testa pra ver se não é o mesmo CTE (Se for, provavelmente entrou em Loop infinito)
                try:
                    # Basicamente, só testa se o fornecedor e o estado forem os mesmos
                    pyautogui.locateOnScreen(diretorioAtual + "\\numeroCTE_atual.png", region=(1353,190,95,32)) # Se achar, abriu o mesmo CTE novamente
                    print('Entrou em Loop Infinito!')
                    mensagemFinal = 'Erro (Possível Loop Infinito)! Foram lançados '+ str(ctesLancados) +' CTEs!' # Só pra indicar que o programa chegou ao fim
                    if rotinaAutomatica:
                        registrarLog(mensagemFinal) # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
                        protheusWindow.close()
                        time.sleep(2)
                        pyautogui.press('left')
                        time.sleep(1)
                        pyautogui.press('enter')
                    else:
                        pyautogui.alert(mensagemFinal) #'This displays some text with an OK button.'
                    exit()  #Encerra o programa se algo der errado quando for tentar abrir o CTE
                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                    im = pyautogui.screenshot(diretorioAtual + "\\numeroCTE_atual.png", region=(1353,190,95,32)) # Atualiza a print do número de CTE porque é um CTE novo

        #Primeiro arruma TES, depois ICMS e por último a alíquota se as respectivas variáveis estão como True
        if tesDiferente: # Parte que arruma a situação tributária (TES) se necessário
            print('Arrumando a Situação Tributária')
            #pyautogui.doubleClick(347, 829) # Movendo a tela pra onde fica a célula da Situação Tributária
            #pyautogui.doubleClick(1196, 313+outrasLinhas)
            for j in range(1,40):
                pyautogui.press('right') # Pressiona a tecla da direita trinta e nove vezes
            pyautogui.press('enter')
            pyautogui.write(tes)
            time.sleep(1)
            #pyautogui.moveTo(347, 829) # Voltando a tela pra parte inicial
            #pyautogui.dragTo(1, 829)
            for j in range(1,41):
                pyautogui.press('left') # Pressiona a tecla da esquerda quarenta vezes pra voltar pra frente de Tipo de Operação
        # Fim do if da TES

        if temICMS: # Parte que arruma o ICMS se tiver
            print('Copiando valor total')
            #pyautogui.doubleClick(904, 313+outrasLinhas) # Clicando direto invés de usar as teclas
            for j in range(1,4):
                pyautogui.press('left') # Pressiona a tecla da esquerda três vezes
            pyautogui.press('enter')
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.hotkey('ctrl', 'v')

            time.sleep(1)

            # Inserindo o valor de ICMS
            print('Corrigindo base de cálculo de ICMS')
            #pyautogui.click(1304, 311+outrasLinhas) # Clicando direto invés de usar as teclas
            for j in range(1,4):
                pyautogui.press('right') # Pressiona a tecla da direita três vezes
            pyautogui.press('enter')
            pyautogui.hotkey('ctrl', 'v')
        # Fim do if do ICMS 

        if aliquotaDiferente: # Parte que arruma a alíquota de ICMS se preciso
            for j in range(1,8):
                pyautogui.press('right') # Pressiona a tecla da direita sete vezes
            pyautogui.write(aliquota)
            pyautogui.press('enter')   
        # Fim do if da alíquota


        # Voltando pro início das linhas
        #for j in range(1,21):
            #pyautogui.press('left') # Pressiona a tecla da esquerda vinte vezes pra voltar pro início
        for j in range(1,11):
            pyautogui.press('left') # Pressiona a tecla da esquerda dez vezes pra voltar pro início
        if temICMS:
            for j in range(1,3):
                pyautogui.press('left') # Pressiona a tecla da esquerda mais duas vezes pra voltar pro início
            if aliquotaDiferente:
                for j in range(1,9):
                    pyautogui.press('left') # Pressiona a tecla da esquerda mais oito vezes pra voltar pro início
        time.sleep(1)
    #Fim do loop For das linhas

    #------------------------------------------------------------- Fim do Loop das Linhas ----------------------------------------------------------------------------

    # Abrindo a aba de Duplicatas para poder colocar o vencimento
    print('Indo pra Duplicatas')
    pyautogui.doubleClick(854, 867)

    time.sleep(1)

    # Colocando a data de vencimento e salvando o CTE
    print('Vencimento')
    pyautogui.doubleClick(279, 927)
    pyautogui.write(dataVencimento)
    time.sleep(1)
    print('Salvando...')
    pyautogui.hotkey('ctrl', 's')
    time.sleep(7) # Padrão: (7), esperando aparecer a mensagem de OK 
    print('Fechando...')
    try:
        location = pyautogui.locateOnScreen(diretorioAtual + "\\botao_fechar_salvar.png",region=(1026,517,283,250)) # Tenta achar o botão pra fechar o aviso que salvou e clicar nele
        #pyautogui.locateOnScreen(diretorioAtual + "\\aviso_OK.png")
        pyautogui.click(pyautogui.center(location))
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        pyautogui.press('enter') # Se não achar, aperta enter
        #pyautogui.click(1165, 639) # Clica no botão de Salvar
        # Normalmente, o botão pra fechar o aviso de OK já vem destacado e é só apertar Enter, mas existe um bug no Protheus que faz com que isso não ocorra raramente
        # Com esse bug, aperta Enter não faz nada e faz com que o programa trave pois Ctrl+W não consegue fechar o CTE se algum aviso estiver aberto
        time.sleep(1)
        try:
            location = pyautogui.locateOnScreen(diretorioAtual + "\\botao_fechar_salvar.png",region=(1026,517,283,250)) # Tenta achar o botão pra fechar o aviso que salvou e clicar nele
            #pyautogui.locateOnScreen(diretorioAtual + "\\aviso_OK.png")
            pyautogui.click(pyautogui.center(location)) # Tenta aperta o botão uma segunda vez pra evitar bugs
            time.sleep(1)
        except:
            pass
    else:
        time.sleep(1)

    try:
        print('Verificando se o CTE foi lançado...') # Ver se o CTE ainda está aberto (Se tiver, é devido que houve algum erro e não deu pra lançar o CTE)
        pyautogui.locateOnScreen(diretorioAtual + "\\classificarAberto.png")
        print('Erro!')
        #exit()  # Encerra o programa se algo der errado quando for tentar salvar o CTE
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(9) #Padrão: (9), esperando o sistema liberar pra ir pro próximo CTE
        pyautogui.press('down')
    except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
        print('Ok!')
        ctesLancados += 1 # Confirma aqui que o CTE foi lançado
        time.sleep(9) #Padrão: (9), esperando o sistema liberar pra ir pro próximo CTE

#Fim do Loop principal
#------------------------------------------------------------- Final da Lógica de Classificação ----------------------------------------------------------------------------

print('Fim! Foram lançados',ctesLancados,'CTEs!') # Ainda mantenho o Print para ter o registro da quantidade de CTEs lançados
mensagemFinal = 'Fim! Foram lançados '+ str(ctesLancados) +' CTEs!' # Só pra indicar que o programa chegou ao fim
if rotinaAutomatica:
    registrarLog(mensagemFinal) # Registra o erro que deu em um arquivo.TXT se for a rotina automática que estiver rodando
    protheusWindow.close()
    time.sleep(2)
    pyautogui.press('left')
    time.sleep(1)
    pyautogui.press('enter')
else:
    pyautogui.alert(mensagemFinal) #'This displays some text with an OK button.'