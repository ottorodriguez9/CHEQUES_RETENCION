
import pyautogui
#import win32gui, win32con
from docx import *
import os.path

"""Automatizador de pagos con retencion"""


captura=Document()
safeoff=pyautogui.FAILSAFE=True

positioner=pyautogui

clicker=pyautogui

writer=pyautogui

print('#'*110,'\n')

#verifica que existe un archivo guardado con datos de una sesion anterior y le indica la usuario que si desea usar esa info o ingresar nuevos datos
#para futura mejora
#a�adir opcion de editar datos anteriores, ej. indicar los campos que existen, cuales campos a editar y cuales campos a eliminar

while True:
    if not os.path.exists('sesion_anterior.txt'):
        print('Bienvenido al sistema automatizacion de cheques con retencion, favor de colocar los datos solitados debajo\n'.center(1))

        print('#'*110,'\n')

        lista_de_campos=[
                        'Beneficiario','Monto neto a pagar','Concepto L1',
                        'Concepto L2','Concepto L3','No. De cuenta ITBIS',
                        'Monto de ITBIS','No. de Cuenta ISR','Monto ISR',
                        'No. de cuenta por pagar','Monto total facturado',
                        'Codigo del suplidor','Balance de cuenta'
                        ]

        Colector_de_datos=[input(f"{_+1}-Coloque {lista_de_campos[_]}:\n ").upper() for _ in range(len(lista_de_campos))]
        print(Colector_de_datos)
        with open('sesion_anterior.txt','w') as file_handler:
            datos= [file_handler.write(f'{Colector_de_datos[a]}\n') for a in range(len(Colector_de_datos))]
        print('#'*60,'\n')

        print('Por favor verique que los datos esten correctos!!!\n')

        print(*([[f'{a}-{b[0]}-{b[1]}'] for a,b in enumerate(zip(lista_de_campos,Colector_de_datos))]),sep='\n')

        print('#'*60,'\n')

        validar=input('Confirmo que los datos esten correctos, de ser asi presione "s", seguido de la tecla "Enter" para continuar, de lo contrario pulse "n" para corregir:\n ').lower()

        while True:
            if validar == 'n':
                campo_a_modificar=int(input('Indique la cantidad de campos a modificar (solo coloque numeros ej. 1,2..10):\n'))
                for i in range(campo_a_modificar):
                    edicion_de_campo=int(input('Ingrese numero del campo a modificar:\n '))
                    Colector_de_datos[edicion_de_campo]=input('Ingrese dato dato a ser corregido:\n ').upper()
                print('Favor de verificar nueva vez, que los datos esten correctos.\n ',*Colector_de_datos,sep='\n')
                confirmacion=input('Si esta todo correcto digite la letra "C", para continuar o "N" para modificar los campos nuevamente:\n ').upper()
                if confirmacion == "N":
                    continue
                else:
                        break
            break
        break
    else:
        print('!!Hemos detectado un archivo con datos de una secion anterior!!!\n')
        
        opcion = input("Desea restaurar los datos tal como estan o desea editar los datos?\n Presione 'R' para restaurar o,\n 'E' para editar los campos").lower()
        
        if opcion=='r':
            with open ('sesion_anterior.txt') as file_handler:
                Colector_de_datos=[a.strip('\n\r') for a in file_handler]
                print (*Colector_de_datos,sep='\n')
                break

	    #elif opcion=='e':
	        #with open ('sesion_anterior.txt') as file_handler:
                #  Colector_de_datos=[a.strip('\n\r') for a in file_handler]
                # print (Colector_de_datos)
		
        else:
            os.remove('sesion_anterior.txt')
            
        
        #print('Sesion terminada, bye!!!')
        #break





####Incluir archivo TXT, para copiar datos guardados en la ultima secion
coordenas = [
        (276,34),(286,103),(323,103),(384,162),
        (170,351),(318,523),(699,568),(153,33),
        (159,103),(288,126),(335,140)
        ]

#Elegir solicitud de banco

def abrir_solicitud():
    for i in range(3):

        positioner.move(coordenas[i][0],coordenas[i][1])
        clicker.click(coordenas[i][0],coordenas[i][1])
        
pyautogui.PAUSE = 1.0

#Colocar datos en solicitud

def datos_solicitud():
    "Esta funcion coloca los datos en las lineas de la solicitud"
    pyautogui.PAUSE = 10.0        
    pyautogui.PAUSE = 1.0
    positioner.move(coordenas[3][0],coordenas[3][1])
    pyautogui.press(['enter','enter'])
    writer.write(Colector_de_datos[0])
    pyautogui.press('enter')
    writer.write(Colector_de_datos[1])
    #pyautogui.press('enter')
    for i in range(2,5):
        pyautogui.press('enter')
        writer.write(Colector_de_datos[i])
    #Digitar conceptos lineas de cheques
##    for i in range(2,5):
##        pyautogui.press('enter')
##        writer.write(Colector_de_datos[i])
    
def digitar_cuentas():
    
    """Esta funcion recoge los datos de las cuentas
    y montos y los coloca en cada casilla"""
    
    positioner.move(coordenas[4][0],coordenas[4][1])
    clicker.click(coordenas[4][0],coordenas[4][1])
    loop=5
    while loop<8:
        #seccion para escribir el numero de cuenta contable
        for i in range(loop,loop+1):
            if len(Colector_de_datos[i])==0:
                continue
            else:
                
                writer.write(Colector_de_datos[i])
                pyautogui.press(['enter','enter'])
                
        for h in range(loop+1,loop+2):
            if len(Colector_de_datos[h])==0:
                continue
            else:
                writer.write(Colector_de_datos[h])
                pyautogui.press('enter',presses=2)
                pyautogui.press('enter')
        loop+=2
    writer.write(Colector_de_datos[9])
    pyautogui.press('enter')
    writer.write(Colector_de_datos[10])
    for i in range(5,7):
        positioner.move(coordenas[i][0],coordenas[i][1])
        clicker.click(coordenas[i][0],coordenas[i][1])
    clicker.click(609,521)
    
def generar_cheque():
    
    """Coloca los datos en el modulo de cheques"""
    
    for i in range(7,9):
        positioner.move(coordenas[i][0],coordenas[i][1])
        clicker.click(coordenas[i][0],coordenas[i][1])
    pyautogui.PAUSE = 10.0        
    pyautogui.PAUSE = 1.0
    writer.write(Colector_de_datos[11])
    pyautogui.press(['enter','enter','enter'])
    positioner.move(coordenas[10][0],coordenas[10][1])
    clicker.click(coordenas[10][0],coordenas[10][1])
    pyautogui.press(['enter','enter'])
    writer.write(Colector_de_datos[10])
    clicker.click(237,441)
    #pyautogui.press('up')
    for i in range(2,5):
        writer.write(Colector_de_datos[i])
        pyautogui.press('enter')
    
    clicker.click(316,522)
    if float(Colector_de_datos[12])< float(Colector_de_datos[1]):
        pyautogui.press('enter')
    #pyautogui.press('enter')
    clicker.click(768,566)
    clicker.click(604,539) 



def editar_cheque():
    
    """ 
        Esta funcion edita los datos
        del cheque para finalizar el
        proceso   
                            
    """
    clicker.click(274,30)
    positioner.move(305,126)
    clicker.click(305,126)
    #pyautogui.press('f2')
    pyautogui.PAUSE = 10.0
    positioner.move(252,218)
    clicker.click(252,218)
    pyautogui.PAUSE = 0.5
    writer.write(Colector_de_datos[1])
    clicker.click(647,360)
    writer.write('0.00')
    pyautogui.press(['enter','enter'])
    pyautogui.press(['enter','enter'])
    writer.write(Colector_de_datos[1])
    pyautogui.press(['enter','enter'])
    pyautogui.press('enter')

    if Colector_de_datos[5]:

    	writer.write(Colector_de_datos[5])
    	pyautogui.press(['enter','enter'])
    
    	writer.write(Colector_de_datos[6])
    	pyautogui.press(['enter','enter'])
    	pyautogui.press('enter')

    if Colector_de_datos[7]:	

    	writer.write(Colector_de_datos[7])
    	pyautogui.press(['enter','enter'])
    	writer.write(Colector_de_datos[8])

    pyautogui.press('prtsc')
    captura.save('captura.docx')
    os.system('start captura.docx')
    os.system('show captura.docx')
    clicker.click()
    pyautogui.hotkey('ctrl','v')
    
##hide = win32gui.GetForegroundWindow()
##win32gui.ShowWindow(hide , win32con.SW_HIDE)
abrir_solicitud()
datos_solicitud()
digitar_cuentas()
generar_cheque()
editar_cheque()
##win32gui.ShowWindow(hide , win32con.SW_SHOW)

