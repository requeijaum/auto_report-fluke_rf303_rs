
#! /usr/bin/env python3


#comentar esse if para testar Python3 no console
#descomentar isso aqui quando construir o .exe com o pyInstaller


#método antigo
#if getattr(sys, 'frozen', False):
#    os.chdir(sys._MEIPASS)


#método novo
#def resource_path(relative_path):
#    """ Get absolute path to resource, works for dev and for PyInstaller """
#    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
#    return os.path.join(base_path, relative_path)
#

#coding: utf-8
#Criado por Rafael Requião em Novembro/2017
#Usar Python3 ... hora de boas práticas



###############################################################################################

#configurações

import os, sys, io, time, datetime, serial, serial.tools.list_ports, openpyxl, string

from datetime import date

#import PIL
#from PIL import Image


#variaveis globais
global 	choice
global 	Liberar

global 	wb
global 	sheet
global	sheet_list
global	s

#imagem pra planilha
#global img
#img = None



#resolver NotDefined's

porta		= ""
ser		= ""

operador 	= ""
data		= ""
hora		= None

equip_tag	= ""
equip_pat 	= ""
equip_serie	= ""

modelo		= ""
fabricante	= ""

#choice		= ""

Liberar		= False
Liberar1    = False
Liberar2	= False


ident		= ""
testado		= False

#desbugar salvamento da planilha entre alterações de worksheet
wb 			= None
#wb.active	= None

sheet		= None
sheet_list	= None
s			= None

pergunta_passa	= None

teste_saida 	= ""
teste_fuga		= ""

connected       = []

cadastrado      = False


#implementar numero de salvamentos de planilha
salvou = 0


#menus aqui

def menu_principal() :


	# coisa nova aqui - estava no final da função
	# serve pra controlar a execução de testes
	# a cada visita aqui no menu_principal() ele verifica teste_comm()

	#testado = False


	#debug do retorno após teste_auto:
	flush_in()	
	choice = ""

	#dar bem-vindo e mostrar menu
	#salvar opcoes em arquivo .txt?	

	print (30 * "-" , "MENU" , 30 * "-")
	print("\n\n* Relatório do Analisador de Bisturi Fluke RF 303 RS v0.1 *\n\n")

	#verificar_vars()

	print("   0 - Carregar planilha de checklist")
	print("   1 - Configurar porta serial")
	print("   2 - Cadastro de operador e equipamento")
	print("   3 - Teste de Potência de Saída")
	print("   4 - Teste de fuga de corrente em alta-frequência")
	print("   5 - Sair do programa")

	print("\n")

	#tirei do loop e coloquei dentro dessa funcao principal	
	verificar_vars()

	


def menu_configurar() :

	global porta
	global lista
	global connected

	#limpar coisas
	clear_screen()
	porta = ""

	try:
		
		#global porta
				
		print("\n\n  Portas seriais disponíveis:\n")
		lista = serial.tools.list_ports.comports()
		connected = []

		for element in lista:
			connected.append(element.device)
	
		#print(str(connected))
	
		for a in range(0 , len(lista)) :
			print("   Porta " + str(a) + " = " + str(lista[a]) )

	 

		pegar_porta()
	

		global ser
		ser = serial.Serial(porta , 2400, timeout=0) 
		''' baudrate fixo '''

		print("\n    Porta \"" + str(porta) + "\" selecionada\n\n")
	
		teste_comm()


	except:
		print("\n   *** Portas seriais não encontradas! *** \n")
		time.sleep(3)



def pegar_porta():

	global porta
	global lista
	global connected
	global ser

	try:

		porta = lista[int(entrar("\n\n  Digite a porta: "))]
		porta = str(porta).split()
		porta = str(porta[0])

	except:
		print("\n   Índice inválido... digite novamente.\n")
		time.sleep(1)


def menu_cadastro() :

	global cadastrado
	
	if cadastrado == True :
	
		ue = entrar("\n\n     Você deseja alterar os dados do operador e do equipamento? (s/n)")
		
		if ue == "s" or ue == "S" :
			time.sleep(1)
			clear_screen()
			cadastrar()
			
		
		if ue == "n" or ue == "N" :
			time.sleep(1)
			clear_screen()
			menu_principal()
			
		else :
			print("\n     ?\n")
			
	
	if cadastrado == False:

		cadastrar()
		


		
def cadastrar() :

	global cadastrado 
	
	clear_screen()
 
	global operador

	operador = entrar("\n\n   Digite o nome do operador: ")

	if(len(operador) > 5) :
			
		today = datetime.date.today()
		now   = datetime.time()
		

		global data

		data = str(today.strftime("%d/%m/%Y"))

		global hora

		hour = str(datetime.datetime.now)
		hour = str.split(hour)
		hora = str(hour[1])

		#verificar strings vazias pra operador

		print("\n\n     Operador \"" + operador + "\" irá executar o procedimento em: " + data + "\n")
		time.sleep(1)

		print("\n\n     Por favor, cadastre o equipamento.\n     Apenas o número de série é obrigatório.\n")
		time.sleep(1)


		global equip_tag
		global equip_pat
		global equip_serie

		equip_tag   = entrar("\n   Digite a TAG do equipamento: ")
		equip_pat   = entrar("\n   Digite o patrimônio do equipamento: ")
		equip_serie = entrar("\n   Digite o número de série do equipamento: ")

		#verificar strings vazias pra equipamento

		if (len(equip_serie) > 2) :
			print("\n\n   Equipamento cadastrado! \n     Tag  : \"" + equip_tag + "\"\n     Pat. : \"" + equip_pat + "\"\n     Série: \"" + equip_serie +"\"")
			
		elif (len(equip_serie) < 2):
			print("\n\n   Equipamento NÃO cadastrado!")
			equip_tag   = ""
			equip_pat   = ""
			equip_serie = ""

		

	elif(len(operador) < 5) :
		print("\n\n     Operador não cadastrado. Nome mínimo: 5 caracteres.")
		operador = ""
		cadastrado = False
		
	cadastrado = True
	
	time.sleep(3)
	menu_principal()
	
	#fim da função




def menu_teste_saida() :

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	global testado

	
	clear_screen()
	verificar_vars()
	clear_screen()
	
	planilha_salvar()
	clear_screen()

	#desbugar troca de modo
	enviar("EXITMODE")
	enviar("SETMODE:OUTPUT")
	print( receber() )

	clear_screen()

	print("\n\n*** TESTES DE POTÊNCIA DE SAÍDA ***")

	print("\n\nLista de testes: \n\n")

	print("   1 - POTÊNCIA COM CARGA DE 100 OHMS - SAÍDA BIPOLAR PRECISE")
	print("   2 - POTÊNCIA COM CARGA DE 100 OHMS - SAÍDA BIPOLAR STANDARD")
	print("   3 - POTÊNCIA COM CARGA DE 100 OHMS - SAÍDA BIPOLAR MACRO")
	print("   4 - POTÊNCIA COM CARGA DE 300 OHMS - SAÍDA MONOPOLAR LOW CUT")
	print("   5 - POTÊNCIA COM CARGA DE 300 OHMS - SAÍDA MONOPOLAR PURE CUT")
	print("   6 - POTÊNCIA COM CARGA DE 300 OHMS - SAÍDA MONOPOLAR BLEND CUT")
	print("   7 - POTÊNCIA COM CARGA DE 500 OHMS - SAÍDA COAGULAÇÃO DESICCATE 1")
	print("   8 - POTÊNCIA COM CARGA DE 500 OHMS - SAÍDA COAGULAÇÃO FULGURATE")
	print("   9 - POTÊNCIA COM CARGA DE 500 OHMS - SAÍDA COAGULAÇÃO SPRAY")

	print("\n\n   Digite \"0\" para retornar ao menu principal\n   Digite \"auto\" para iniciar sequência automática dos testes acima")

	#pedaço importante de código pra funcionar com a planilha


	try:
		#global wb.active
		wb.active = 1	
		print("\n   DEBUG: wb.active = " + str(wb.active))
		#sheet = sheet_list[1]
		#s = sheet.decode("utf8").encode("ascii", "replace")

		#s = sheet

		#print("     DEBUG: sheet = " + str(sheet))
		#print("sheet_list: \n" + str(sheet_list))

		#print("\n" + "s = sheet = " + s + "\n")

		#s = s.decode('some_encoding').encode('ascii', 'replace')

		#tentar resolver erro: "NameError: name 'wb' is not defined"

	except:
		print("\n     Erro: Valor de \"wb\" não declarado anteriormente. \n           Precisa carregar uma planilha antes.")

	#pedaço terminou ;)

	global teste_saida
	
	#RESOLVER BUG do retorno de teste auto
	teste_saida = None

	teste_saida = entrar("\n\n   Selecione o teste desejado: ")

	if   teste_saida == "1" : 
		teste_saida_01()

	elif teste_saida == "2" : 
		teste_saida_02()

	elif teste_saida == "3" : 
		teste_saida_03()

	elif teste_saida == "4" : 
		teste_saida_04()

	elif teste_saida == "5" : 
		teste_saida_05()

	elif teste_saida == "6" : 
		teste_saida_06()

	elif teste_saida == "7" : 
		teste_saida_07()

	elif teste_saida == "8" : 
		teste_saida_08()

	elif teste_saida == "9" : 
		teste_saida_09()

	elif teste_saida == "auto" or teste_saida == "AUTO" : 
		teste_saida_auto()

	elif teste_saida == "0" : 
		testado = False
		time.sleep(1)		
		menu_principal()

		#DEBUG:
		print("teste_saida = " + str(teste_saida))
		print("choice      = " + str(choice))


	else: 
		entrar("Opção não-reconhecida. Tente novamente.")




#Sequências de teste de potência

def teste_saida_01():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s


	print("\n# POTENCIA COM CARGA DE 100 OHMS - SAIDA BIPOLAR PRECISE\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:100')
	
	print("\n## Valor nominal: 10 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E16'] = c
	wb.active['C16'] = p

	print("\n## Valor nominal: 30 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E17'] = c
	wb.active['C17'] = p

	print("\n## Valor nominal: 70 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E18'] = c
	wb.active['C18'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()



def teste_saida_02():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 100 OHMS - SAIDA BIPOLAR STANDARD\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:100')
	print("\n## Valor nominal: 10 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E23'] = c
	wb.active['C23'] = p

	print("\n## Valor nominal: 30 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E24'] = c
	wb.active['C24'] = p

	print("\n## Valor nominal: 70 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E25'] = c
	wb.active['C25'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()



def teste_saida_03():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 100 OHMS - SAIDA BIPOLAR MACRO\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:100')
	print("\n## Valor nominal: 10 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E30'] = c
	wb.active['C30'] = p

	print("\n## Valor nominal: 30 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E31'] = c
	wb.active['C31'] = p

	print("\n## Valor nominal: 70 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E32'] = c
	wb.active['C32'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()



def teste_saida_04():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 300 OHMS - SAIDA MONOPOLAR LOW CUT\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:300')
	print("\n## Valor nominal: 10 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E37'] = c
	wb.active['C37'] = p

	print("\n## Valor nominal: 75 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E38'] = c
	wb.active['C38'] = p

	print("\n## Valor nominal: 300 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E39'] = c
	wb.active['C39'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()


def teste_saida_05():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 300 OHMS - SAIDA MONOPOLAR PURE CUT\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:300')
	print("\n## Valor nominal: 10 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E44'] = c
	wb.active['C44'] = p

	print("\n## Valor nominal: 75 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E45'] = c
	wb.active['C45'] = p

	print("\n## Valor nominal: 300 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E46'] = c
	wb.active['C46'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()


def teste_saida_06():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 300 OHMS - SAIDA MONOPOLAR BLEND CUT\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:300')
	print("\n## Valor nominal: 10 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E51'] = c
	wb.active['C51'] = p

	print("\n## Valor nominal: 75 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E52'] = c
	wb.active['C52'] = p

	print("\n## Valor nominal: 200 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E53'] = c
	wb.active['C53'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()


def teste_saida_07():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 500 OHMS - SAIDA DE COAGULAÇÃO DISSECATE 1 \n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:500')
	print("\n## Valor nominal: 1 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E58'] = c
	wb.active['C58'] = p

	print("\n## Valor nominal: 30 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E59'] = c
	wb.active['C59'] = p

	print("\n## Valor nominal: 120 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E60'] = c
	wb.active['C60'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()


def teste_saida_08():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 500 OHMS - SAIDA DE COAGULAÇÃO FULGURATE\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:500')
	print("\n## Valor nominal: 1 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E65'] = c
	wb.active['C65'] = p

	print("\n## Valor nominal: 30 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E66'] = c
	wb.active['C66'] = p

	print("\n## Valor nominal: 120 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E67'] = c
	wb.active['C67'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()



def teste_saida_09():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# POTENCIA COM CARGA DE 500 OHMS - SAIDA DE COAGULAÇÃO SPRAY\n")
	enviar('SETMODE:OUTPUT')
	enviar('SETLOAD:500')
	print("\n## Valor nominal: 1 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E72'] = c
	wb.active['C72'] = p

	print("\n## Valor nominal: 30 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E73'] = c
	wb.active['C73'] = p

	print("\n## Valor nominal: 120 W\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	medir(a="potencia")
	wb.active['E74'] = c
	wb.active['C74'] = p


	global teste_saida

	if (teste_saida == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_saida()




def teste_saida_auto():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	#DEBUG do retorno de menu apos teste "auto": choice... vamos zerar?

	global pergunta_passa
	global teste_saida
	global choice
	
	#mesmo parando... ele continua o teste... preciso travar a ordem
	clear_screen()

	print("\n\n\n    Sequência de testes automática   \n\n\n")
	#DEBUG:
	print("pergunta_passa = " + str(pergunta_passa))

	teste_saida_01()

	#DEBUG:
	print("pergunta_passa = " + str(pergunta_passa))

	if (pergunta_passa == True):
		teste_saida_02()	

		if (pergunta_passa == True):
			teste_saida_03()

			if (pergunta_passa == True):
				teste_saida_04()

				if (pergunta_passa == True):
					teste_saida_05()

					if (pergunta_passa == True):	
						teste_saida_06()

						if (pergunta_passa == True):
							teste_saida_07()

							if (pergunta_passa == True):
								teste_saida_08()
							
								if (pergunta_passa == True):
									teste_saida_09()
									
	#acho q n está salvando
	planilha_salvar()

	#parece que o programa n sabe sair direito e vai para o menu principal...

	#elif, if ou else?
	if (pergunta_passa == False):
		
		#isso buga?
		teste_saida = None

		choice     = None
		testado    = True

		menu_teste_saida()




# Menu de testes de fuga e IEC 60601


def menu_teste_fuga() :

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	global testado
	
	clear_screen()
	verificar_vars()
	clear_screen()
	
	planilha_salvar()
	clear_screen()
	
	#desbugar troca de modo
	enviar("EXITMODE")
	enviar("SETMODE:OUTPUT")
	enviar("SETLOAD:200")
	enviar("EXITMODE")
	enviar("SETMODE:RFLKG")
	enviar("RDLOAD")
	
	print( receber() )

	clear_screen()

	print("\n\n*** TESTE DE FUGA DE ALTA FREQUÊNCIA ***")
	print("Carga de 200 ohms, terra do chasi, potência máxima")


	
	print("\n\nLista de testes: \n\n")

	print("   1 - FUGA DE RETORNO PACIENTE")
	print("   2 - FUGA ATIVA")
	print("   3 - FUGA BIPOLAR (PONTA DIREITA)")
	print("   4 - FUGA BIPOLAR (PONTA ESQUERDA)")
	#print("   5 - TESTE DE SEGURANÇA CONFORME IEC 60601")

	print("\n\n   Digite \"0\" para retornar ao menu principal\n   Digite \"auto\" para iniciar sequência automática dos testes acima\n\n\n\n")


	#pedaço importante de código pra funcionar com a planilha

	global sheet
	global sheet_list
	global s

	try:
	#global wb.active
		wb.active = 2	
		print("\n   DEBUG: wb.active = " + str(wb.active))
		#sheet = sheet_list[2]
		#s = sheet
		#print("     DEBUG: sheet = " + str(sheet))
		#print("sheet_list: \n" + str(sheet_list))


	except:
		print("\n     Erro: Valor de \"wb\" não declarado anteriormente. \n           Precisa carregar uma planilha antes.")

	#pedaço terminou ;)

	global teste_fuga

	#RESOLVER BUG do retorno de teste auto
	teste_fuga = None

	teste_fuga = entrar("\n\n   Selecione o teste desejado: ")

	if  teste_fuga == "1" : 
		teste_fuga_01()

	elif teste_fuga == "2" : 
		teste_fuga_02()

	elif teste_fuga == "3" : 
		teste_fuga_03()

	elif teste_fuga == "4" : 
		teste_fuga_04()

	#elif teste_fuga == "5" : 
	#	teste_fuga_05()

	elif teste_fuga == "auto" or teste_fuga == "AUTO" : 
		teste_fuga_auto()

	elif teste_fuga == "0" : 
		testado = False
		time.sleep(1)		
		menu_principal()

		#DEBUG:
		print("teste_fuga = " + str(teste_fuga))
		print("choice      = " + str(choice))

	else: 
		entrar("Opção não-reconhecida. Tente novamente.")



#Sequências de teste de fuga de corrente em Alta-freq


def teste_fuga_01():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# FUGA DO RETORNO DO PACIENTE")
	print("# 200 ohms ; Terra do Chassi ; Potência máx. (Corte e Coag.)\n")

	enviar('SETMODE:RFLKG')
	enviar('SETLOAD:200')
	enviar("RDLOAD")



	print("\n## Alterar para: LOW CUT\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C15'] = c


	print("\n## Alterar para: PURE CUT\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C16'] = c


	print("\n## Alterar para: BLEND CUT\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C17'] = c


	print("\n## Alterar para: SAÍDA DE COAGULAÇÃO DESICCATE 1\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C18'] = c


	print("\n## Alterar para: SAÍDA DE COAGULAÇÃO FULGURATE\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C19'] = c


	print("\n## Alterar para: SAÍDA DE COAGULAÇÃO SPRAY\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C20'] = c


	global teste_fuga

	if (teste_fuga == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_fuga()


def teste_fuga_02():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s


	print("\n# FUGA ATIVA")
	print("# 200 ohms ; Terra do Chassi ; Potência máx. (Corte e Coag.)\n")

	enviar('SETMODE:RFLKG')
	enviar('SETLOAD:200')
	enviar("RDLOAD")

	print("\n## Alterar para: LOW CUT\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C25'] = c


	print("\n## Alterar para: PURE CUT\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C26'] = c


	print("\n## Alterar para: BLEND CUT\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C27'] = c


	print("\n## Alterar para: SAÍDA DE COAGULAÇÃO DESICCATE 1\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C28'] = c


	print("\n## Alterar para: SAÍDA DE COAGULAÇÃO FULGURATE\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C29'] = c


	print("\n## Alterar para: SAÍDA DE COAGULAÇÃO SPRAY\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C30'] = c


	global teste_fuga

	if (teste_fuga == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_fuga()




def teste_fuga_03():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s


	print("\n# FUGA BIPOLAR (PONTA DIREITA)")
	print("# 200 ohms ; Terra do Chassi ; Potência máx. (Corte e Coag.)\n")

	enviar('SETMODE:RFLKG')
	enviar('SETLOAD:200')
	enviar("RDLOAD")

	print("\n## SAÍDA BIPOLAR PRECISE\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C35'] = c


	print("\n## SAÍDA BIPOLAR STANDARD\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C36'] = c


	print("\n## SAÍDA BIPOLAR MACRO\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C37'] = c



	global teste_fuga

	if (teste_fuga == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_fuga()




def teste_fuga_04():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	print("\n# FUGA BIPOLAR (PONTA ESQUERDA)")
	print("# 200 ohms ; Terra do Chassi ; Potência máx. (Corte e Coag.)\n")

	enviar('SETMODE:RFLKG')
	enviar('SETLOAD:200')
	enviar("RDLOAD")

	print("\n## SAÍDA BIPOLAR PRECISE\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C42'] = c


	print("\n## SAÍDA BIPOLAR STANDARD\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C43'] = c


	print("\n## SAÍDA BIPOLAR MACRO\n")
	entrar("   Pressione \"Enter\" para continuar...\n")

	medir(a="corrente")
	wb.active['C44'] = c



	global teste_fuga

	if (teste_fuga == "auto") :
		pergunta_sair()

	else :
		planilha_salvar()
		menu_teste_fuga()







def teste_fuga_auto():

	global 	wb
	#global wb.active
	global 	sheet
	global	sheet_list
	global	s

	#DEBUG do choice...

	global teste_fuga
	global choice
	global pergunta_passa
	
	#mesmo parando... ele continua o teste... preciso travar a ordem
	clear_screen()

	print("\n\n\n    Sequência de testes automática   \n\n\n")
	#DEBUG:
	print("pergunta_passa = " + str(pergunta_passa))

	teste_fuga_01()

	#DEBUG:
	print("pergunta_passa = " + str(pergunta_passa))

	if (pergunta_passa == True):
		teste_fuga_02()	

		if (pergunta_passa == True):
			teste_fuga_03()

			if (pergunta_passa == True):
				teste_fuga_04()

				#if (pergunta_passa == True):
					#teste_fuga_05()

	
	#acho q n está salvando
	planilha_salvar()
	
	#elif, if ou else?
	if (pergunta_passa == False):

		#isso buga?
		teste_fuga = None

		choice     = None
		testado    = True
		menu_teste_fuga()
		




#funcoes aqui


from platform import system as system_name # Returns the system/OS name
from os import system as system_call       # Execute a shell command


def clear_screen():
	os.system('cls' if os.name=='nt' else 'clear')
	
	
	
	
def flush_in():

	#bug introduzido a esta função no quando mudei de:
	#Python 3.4 para Python 3.6.3
	
	#https://docs.python.org/3/library/termios.html	
	'''
	Set the tty attributes for file descriptor fd from the attributes, which is a list like the one returned by tcgetattr(). The when argument determines when the attributes are changed: 
	TCSANOW to change immediately, TCSADRAIN to change after transmitting all queued output, or TCSAFLUSH to change after transmitting all queued output and discarding all queued input.
	'''
	
	#https://linux.die.net/man/3/tcflush
	

	try:
		import msvcrt
		while msvcrt.kbhit():
			msvcrt.getch()
			
	except ImportError:
		import sys, termios, tty 			#inclui modulo "tty"
		#termios.tcflush(sys.stdin, termios.TCIOFLUSH) 	#TCIOFLUSH tá bugando?
		sys.stdin.flush()


def entrar(texto):
	
	#if os.name=='nt':
	#	flush_in() 
	#
	#else:
	#	sys.stdin.flush()
		


	flush_in() 

	return input(texto)
	




def verificar_vars():

	global Liberar
	global wb
	global data
	global hora
	global pergunta_passa

	#DEBUG aqui
	print("     bool Liberar = " + str(Liberar) + ", Liberar1 = " + str(Liberar1) + ", Liberar2 = " + str(Liberar2))
	print("     bool pergunta_passa = " + str(pergunta_passa))
	print("     wb = " + str(wb))

	#verificar preenchimento das variaveis de configuracao
	if (len(porta) == 0) :
		print("\n   *** Porta serial não configurada! ***")

	elif (len(porta) > 3) :
		print("\n   *** Porta serial: \"" + porta + "\" ***\n")
	
	if (((len(operador) == 0) or (len(operador) < 5) ) & (len(equip_serie) == 0)) : 
		print("\n   *** Operador e/ou equipamento não cadastrados! ***\n")
	
	elif ((len(operador) > 5) & (len(equip_serie) > 2)) :  
		#mudei len(serie) de 5 pra 2
		#print("\n")
		print("   *** Operador    -->  \""+ operador + "\"")
		#print("   *** Equipamento -->  TAG  : \"" + equip_tag + "\"" )
		#print( 23 * " " , "Pat. : \"" + equip_pat + "\"")
		#print( 23 * " " , "Série: \"" + equip_serie + "\" \n")
	
		print("   *** Equipamento -->  TAG  : \"" + equip_tag + "\"" + "   Pat. : \"" + equip_pat + "\"")
		print( 23 * " " , "Série: \"" + equip_serie + "\" \n")
	

	else :
		print("    \"" + operador + "\" em \"" + data + "\" \n")
		

	if ( (len(str(wb)) == 0) or (wb == None) or (wb == "") ) :
		print("   *** PLANILHA NÃO CARREGADA *** \n\n")
		
	elif (len(str(wb)) > 3) :
		print("   *** PLANILHA CARREGADA *** \n\n")
		
	




def teste_comm():

	global ident
	global wb
	global Liberar
	global Liberar1
	global Liberar2


	#configurar liberação por planilha carregada

	if(len(str(wb)) == 0) :
		Liberar2 = False

	elif (len(str(wb)) > 3 ) :
		Liberar2 = True

	print("DEBUG: Liberar2 = " + str(Liberar2))
	time.sleep(1)

	ident = ""

	#if ( (choice == 3) or (choice == 4) ):

	if( len(porta) > 0 ):

		print("\n\nTestando comunicação...")
		enviar("GOTOLOCAL")		
		enviar("GOTOREMOTE")
		print(receber())
		enviar("EXITMODE")
		receber()
		enviar("IDENT")
		ident = receber()

		if "RF303" not in ident:
			print("\n   *** RF303 não reconhecido! ***\n")
			Liberar1 = False

		elif "RF303" in ident :
			#print("Identificação: " + ident)
			Liberar1 = True

		#Debug:
		print("Ident.  : " + ident)


	#verifica travas de execução de teste

	if ((Liberar1 == True) and (Liberar2 == True)) :
		Liberar = True

	elif ((Liberar1 == False) or (Liberar2 == False) ):
		Liberar = False

		
	#DEBUG
	print("Liberar1: " + str(Liberar1))
	print("Liberar2: " + str(Liberar2))
	print("Liberar : " + str(Liberar))

	time.sleep(2)


def testado_comm() :

	global testado

	if (testado == False) :
		teste_comm()
		testado = True

	if (testado == True)  :
		print("\n     DEBUG: Comunicação já foi testada...")





def enviar(texto_enviado):
	ser.flush()
	ser.write(bytearray(texto_enviado + "\r", "ascii"))
	time.sleep(1/4) #mudei de 1 pra 0.5, hehe

	#DEBUG:
	print("> " , texto_enviado)
	
	#Python 3.6.4 do Xubuntu 17.04 deve ter introduzido bug em print() após enviar()
	return 0


def receber():
	
	texto_retornar = b"texto_retornar"
	
	texto_recebido = bytearray(ser.read(size=128)) # ajustar size e encoding
	
	time.sleep(1/4)	
	
	lista_recebido = texto_recebido.splitlines()

	#Debug:
	#print("   lista_recebido: " + str(lista_recebido))

	#lista_recebido: [bytearray(b'*'), bytearray(b'*'), bytearray(b'MODEL RF303 VERSION  1.10')]

	for a in range(0, len(lista_recebido)):
	
		if ( len(lista_recebido[a])  > 1) : texto_retornar = lista_recebido[a]
	
	#DESBUGAR com a linha abaixo
	texto_retornar = texto_retornar.decode("ascii")
	#mas parece que bugou... comentar de novo.
	
	ser.flush()
	texto_recebido = ""
	return texto_retornar
	
	

def erro_teste():

	global Liberar
	global Liberar1
	global Liberar2

	global operador
	global equip_serie

	clear_screen()
	print("\n\n\n\n     ERRO!\n     O teste selecionado não será realizado")
	
	if (Liberar1 == False) :
		print("\n     *** RF303 não reconhecido ***")


	if ((len(operador) == 0) & (len(equip_serie) == 0)) : 
		print("\n     *** Operador e equipamentos não cadastrados! ***\n")

	if (len(str(wb)) == 0) :
		print("\n     *** PLANILHA NÃO CARREGADA ***\n\n ")
		


	time.sleep(5)
	menu_principal()


	
	


#Implementar média de medidas de corrente e potência - varia bastante entre uma medida e outra


def medir(a) :

	#limpar buffer - não previ o acender de display e retorno de algum valor durante SETMODE/SETLOAD
	print("DEBUG: receber() = " + receber() + "\n")

	try: 
		if a == "corrente":

			global c
			c    = 0
			sumc = 0
				
			for a in range(1, 4):
				#se "corrente"
				#print("Medida #0"+ str(a))			
				enviar("RDCURRENT")
				time.sleep(1/2) #debug
				RDCURRENT = receber()
				corrente = RDCURRENT
				#tratar bytearray, converter pra número inteiro
				corrente = corrente.split()
				corrente = corrente[0]
				#corrente = corrente.decode("ascii")
				#print("   corrente = " + corrente)
				c = int(corrente)
				print("   Medida #0"+ str(a) + "c = " + str(c))
				
				sumc = sumc + c
				time.sleep(1/2) #debug
				#print("   sumc = " + str(sumc))
				#verificar endianess
				if a == 3 : break			
			
			#feito 3 vezes... somar valor e depois dividir
			c = (sumc / 3)
			c = round(c)
			#tirei media :p
			print("\n--> Corrente média = " + str(c))
			print("\n")	
			
			#concatenar texto
			s1   = " "
			seq1 = ( str(c), " (mA)")
			c = s1.join( seq1 )
			
			return c
				
		
		if a == "potencia":

			global p
			p    = 0
			sump = 0
				
			for a in range(1, 4):
				#se "potencia"
				#print("Medida #0"+ str(a))	
				enviar("RDPOWER")
				
				RDPOWER = receber()
				time.sleep(1/2) #debug
				potencia = RDPOWER
				#tratar bytearray, converter pra número inteiro
				potencia = potencia.split()
				potencia = potencia[0]

				#print("potencia = " + potencia)
				p = float(potencia)
				print("   Medida #0"+ str(a) + "p = " + str(p))
				
				sump = sump + p
				time.sleep(1/2) #debug
				#print("   sump = " + str(sump))
				#verificar endianess
				if a == 3 : break	
			
			#feito 3 vezes... somar valor e depois dividir
			p = (sump / 3)
			p = round(p)
			#tirei media :p
			print("\n--> Potência média = " + str(p))
			print("\n")	

			
			#concatenar texto
			s2   = " "
			seq2 = ( str(p), " (W)")
			p = s2.join( seq2 )
			
			
			return p
			
	except:
		print("DEBUG: medir() não retornou valor numérico", "c = ", c , "; p = ", p)
		#tentar de novo?
		

def planilha_carregar():

	global testado

	clear_screen()

	global  wb
	#global wb.active
	global  sheet
	global  s
	global  sheet_list

	#chamar planilha

	from openpyxl import Workbook

	try:
		wb = openpyxl.load_workbook('checklist_bisturi.xlsx'	)
		print("\nCarregou planilha \"checklist_bisturi.xlsx\"...\n")
	
	except:
		print("\n     Erro: Planilha \"checklist_bisturi.xlsx\" não foi encontrada.")


	try:
		sheet_list = wb.get_sheet_names()
		sheet_list = str.format(str(sheet_list))
		print("sheet_list: \n" + str(sheet_list))

	except:
		print("\n     Erro: Valor de \"wb\" não declarado anteriormente. \n           Precisa carregar uma planilha antes.")


	time.sleep(2)








def planilha_salvar():

	#pedaço importante de código pra funcionar com a planilha

	global salvou
	
	global wb
	#global wb.active	
	global sheet
	global s

	global operador
	global data
	global hora

	global equip_tag
	global equip_pat
	global equip_serie

	global modelo
	global fabricante


	#global analisador_tag
	#global analisador_pat
	#global analisador_serie

	from openpyxl import writer


	try:
		wb.active = 0
		#sheet = sheet_list[0]
		#s = sheet.decode("utf8").encode("ascii", "replace")
		s = sheet

		#print("     DEBUG: sheet = " + str(sheet))
		#print("sheet_list: \n" + str(sheet_list))

		#print("\n" + "s = sheet = " + s + "\n")

		#s = s.decode('some_encoding').encode('ascii', 'replace')

		#tentar resolver erro: "NameError: name 'wb' is not defined"

	except:
		print("\n     Erro: Valor de \"wb\" não declarado anteriormente. \n           Precisa carregar uma planilha antes.")

	#pedaço terminou ;)



	wb.active['C27'] = data
	
	#ta bugado  a hora
	wb.active['C28'] = hora 
	
	wb.active['D35'] = operador
	

	wb.active['C22'] = equip_tag
	wb.active['C23'] = equip_pat
	wb.active['C24'] = equip_serie

	
	#planilha_perfumaria()


	#método novo de salvar arquivo
	#CRIAR UM TRY E EXCEPT PARA VERIFICAR SE DÁ PRA ESCREVER ARQUIVO
	
	
	try:
	
		global salvou
		
		#entendi como usar caminho absoluto...
		if (os.name == 'nt'):
			local = os.path.abspath("") + "\\"
			
		else:
			local = os.path.abspath("") + "/"
			
		extensao = (".xlsx")	
		nome_arquivo = operador + " " + str(salvou) + " " + equip_serie + extensao
		
		
		#verificar conversão de objeto 'Worksheet' para dados binários
		#tentar salvar em workbook virtual pra uma variavel ou algo assim
		'''
		>>> dir(openpyxl.writer.excel.save_virtual_workbook)

		['__annotations__', '__call__', '__class__', '__closure__', '__code__', 
		'__defaults__', '__delattr__', '__dict__', '__dir__', '__doc__', '__eq__', 
		'__format__', '__ge__', '__get__', '__getattribute__', '__globals__', 
		'__gt__', '__hash__', '__init__', '__init_subclass__', '__kwdefaults__', 
		'__le__', '__lt__', '__module__', '__name__', '__ne__', '__new__', 
		'__qualname__', '__reduce__', '__reduce_ex__', '__repr__', '__setattr__', 
		'__sizeof__', '__str__', '__subclasshook__']

		'''	
			
		#hex(id(wb))   #me dá o endereço do objeto "wb" tipo 'Workbook'... preciso converter pra bytes
		
		dados = 		writer.excel.save_virtual_workbook(wb)
		
		arquivo = 		open(nome_arquivo, mode="wb")
		
		arquivo.write(dados)
		arquivo.close()
			
		
		#metodo antigo de salvar arquivo
		#wb.save(filename = 'checklist_bisturi_alterado.xlsx' )
		
		#colocar no final
		
		#implementar alguma forma de salvar a planilha sem falhas por engano de substituição
		#salvou = salvou +1
		#desisto
		
		print("\n\n     Planilha salva com sucesso!  \n\n")
		
		
	except:
		print("\n\n\n     FALHA AO SALVAR O ARQUIVO.")
		print("\n\n\n     VERIFIQUE AS PERMISSÕES DO DIRETÓRIO ATUAL.")
		
	time.sleep(3)

	




#def planilha_perfumaria() :
	
	#colocar bordas na área externa das 4 páginas


	#colocar imagem das marcas da liga e do HMG

	'''
	Traceback (most recent call last):
	  File "/usr/local/lib/python3.4/dist-packages/openpyxl/drawing/image.py", line 29, in _import_image
	    import Image as PILImage
	ImportError: No module named 'Image'

	During handling of the above exception, another exception occurred:

	Traceback (most recent call last):
	  File "/usr/local/lib/python3.4/dist-packages/openpyxl/drawing/image.py", line 31, in _import_image
	    from PIL import Image as PILImage
	ImportError: No module named 'PIL'

	During handling of the above exception, another exception occurred:

	Traceback (most recent call last):
	  File "./r12.py", line 1750, in <module>
	    menu_teste_saida()
	  File "./r12.py", line 261, in menu_teste_saida
	    planilha_salvar()
	  File "./r12.py", line 1614, in planilha_salvar
	    planilha_perfumaria()
	  File "./r12.py", line 1640, in planilha_perfumaria
	    img = Image('marca.png')
	  File "/usr/local/lib/python3.4/dist-packages/openpyxl/drawing/image.py", line 53, in __init__
	    image = _import_image(img)
	  File "/usr/local/lib/python3.4/dist-packages/openpyxl/drawing/image.py", line 33, in _import_image
	    raise ImportError('You must install PIL to fetch image objects')
	ImportError: You must install PIL to fetch image objects

	from openpyxl.drawing.image import Image
	'''
	'''
	global img

	for i in range(0,4):
		wb.active = i
		img = Image.open('marca.png')
		img.anchor(wb.active.cell('B5'))
		wb.active.add_image(img , 'B5')
	'''

	










def sair() :

	clear_screen()
	print("\n\n sair \n\n")
	quit()


def pergunta_sair() :

	global wb
	global pergunta
	global pergunta_passa
	global choice

	pergunta = entrar("\n     Deseja continuar o teste? (S/N)  --> ")

	if(pergunta == "s" or pergunta == "S" or pergunta == "sim" or pergunta == "SIM" ) :
		pergunta_passa = True
		print("\n   Continuando")
		time.sleep(1)			


	elif(pergunta == "n" or pergunta == "N" or pergunta == "nao" or pergunta == "NAO" ) :

		pergunta_passa = False

		#invocar planilha_salvar() aqui ou a cada teste?
	
		#como retornar pro menu anterior?
		if (choice == "3") :
			#planilha_salvar()
			print("\n\n   Retornando ao menu principal")
			time.sleep(1)	
			testado = True
			menu_teste_saida()

		elif (choice == "4") :
			#planilha_salvar()
			print("   Retornando ao menu principal")
			time.sleep(1)	
			testado = True
			menu_teste_fuga()


	else : 
		print("?")
		pergunta_sair()

	#tentar passar valor em retorno... tem algum bug
	return pergunta_passa




















#programa principal aqui em loop

loop    = True

while loop :

	clear_screen()
	menu_principal()

	#DEBUG do retorno de teste_auto
	#flush_in()
	#choice = ""

	choice = entrar("\n\n  Selecione a opção desejada: ")

	if choice == "0" : 
		planilha_carregar()

	elif choice == "1" : 
		menu_configurar()

	elif choice == "2" : 
		menu_cadastro()

	elif choice == "3" : 
		if (Liberar == True and (len(str(wb)) > 3) and (len(operador) > 1 ) ): 
			menu_teste_saida()

		elif (Liberar == False or (len(str(wb)) == 0) or (len(operador) == 0 )) :
			erro_teste()


	elif choice == "4" : 
		if (Liberar == True and (len(operador) > 1 ) ): 
			menu_teste_fuga()

		elif (Liberar == False or (len(operador) == 0 )) :
			erro_teste()


	elif choice == "5" : 
		sair()
		loop = False

	else: 
		entrar("Opção não-reconhecida. Tente novamente.")




