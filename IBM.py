# -*- coding: utf-8 -*-
"""
@author: JEMU

MODULO AUXILIAR DONDE SE MUESTRAN UNOS POCOS PROGRAMAS SENCILLOS Y SIMPLIFICADOS.
UTILIZAN LA LIBRERIA win32com YA QUE LA VERSION DE IBM PCCOM DE LA EMPRESA ES LA 6.0 DE 32Bits. 
LO QUE ES MUY ESPECIFICA PARA REALIZAR AUTOMATIZACIONES DE OBJETOS ENTRE MAQUINAS DE 64bits e IBM de 32bits.
"""
import time
import logging
from win32com import client # pip install pywin32 REQUIRED LIBRARY ASSUMED TO BE INSTALLED
import autoit # pip install pyautoit REQUIRED LIBRARY ASSUMED TO BE INSTALLED
from PyQt5.QtWidgets import  QMessageBox
# Configure logging
logging.basicConfig(level=logging.INFO)

def _launch_pcomm_session(pcomm_exe_path, ws_file_path):
    lauch = autoit.run(f'"{pcomm_exe_path}" "{ws_file_path}"')
    time.sleep(5)  # Wait for the PCOMM window to open
    return lauch
def _show_warning(message):
    QMessageBox.warning(None, 'Warning', message)
def _connect_to_sesion():
    try:
        # Create session object
        autECLSession_obj = client.Dispatch("PCOMM.autECLSession")
    
        # Refresh session list and try to connect to an existing session
        autECLSession_obj.RefreshSessions()
        sessions = autECLSession_obj.ListAllSessionNames()
        logging.info(f"Available sessions: {sessions}")
        # Test Sesion usually ask the user to Select one from a dialog window. 
        session_name = 'A'  
        if session_name in sessions:
            autECLSession_obj.SetConnectionByName(session_name)
        else:
            #Test directories usually ask the user to open then from a dialog window.
            pcomm_exe_path = "C:\\user\\pcsws.exe"
            ws_file_path = "C:\\user\\private\\sesion-a.ws"
            # Launch a new session if none are found
            _launch_pcomm_session(pcomm_exe_path, ws_file_path)
            autECLSession_obj.SetConnectionByName(session_name)
        return autECLSession_obj
        
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        

def ABRIR_DEUDA(value):
    #Connect to PCOMM
    Sesion_1 = _connect_to_sesion()
    try:
        #Look if the user is in the right module to execute the program if not ask the user to go to the right module and re-run the program 
        if Sesion_1.autECLPS.GetTextRect(1,1,10,80).find("SC0008") !=-1 and Sesion_1.autECLPS.GetTextRect(1,1,10,80).find("D0012") !=-1:
            #Wait times for the PCOMM to load
            Sesion_1.autECLOIA.WaitForAppAvailable() 
            Sesion_1.autECLOIA.WaitForInputReady()
            #Enter the cliente code in the system
            Sesion_1.autECLPS.SendKeys(value, 12, 34)
            Sesion_1.autECLOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[enter]")
            Sesion_1.autECLOIA.WaitForAppAvailable()
            Sesion_1.autECLOIA.WaitForInputReady()
            #Select the option where the program will change the Client Status
            Sesion_1.autECLPS.SendKeys("1", 17, 23)
            Sesion_1.autECLOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[enter]")       
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[pf11]")        
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("2")        
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[enter]")       
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            #Add a comment in the client page dating the day of re-open for payment
            Sesion_1.autECLPS.SendKeys("[pf8]")        
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("CUENTA ABIERTA POR PAGO")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[enter]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            #Exist the client back to the client access Screen
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            #Check if a pop-up window appears if the client cannot buy every product.
            Sesion_1.autECLPS.GetTex(4, 29,80)            
            if Sesion_1.autECLPS.GetTex(4, 29,80).find("Uni") !=-1:            
                Sesion_1.autECLIOIA.WaitForAppAvailable()
                Sesion_1.autECLIOIA.WaitForInputReady()
                Sesion_1.autECLPS.SendKeys("[pf12]")  
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
        else:
            _show_warning("No está en el módulo correcto.\nVaya a la opción 1217.\nYvuelva a ejecutar el programa")
    except Exception as e:
        logging.error(f"An error occurred in Abrir_Deuda: {e}")


def CERRAR_DEUDA(value):
    #Connect to PCOMM
    Sesion_1 = _connect_to_sesion()
    try:
        #Look if the user is in the right module to execute the program if not ask the user to go to the right module and re-run the program 
        if Sesion_1.autECLPS.GetTextRect(1,1,10,80).find("SC0008") !=-1 and Sesion_1.autECLPS.GetTextRect(1,1,10,80).find("D0012") !=-1:
            #Wait times for the PCOMM to load
            Sesion_1.autECLOIA.WaitForAppAvailable() 
            Sesion_1.autECLOIA.WaitForInputReady()
            #Enter the cliente code in the system
            Sesion_1.autECLPS.SendKeys(value, 12, 34)
            Sesion_1.autECLOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[enter]")
            Sesion_1.autECLOIA.WaitForAppAvailable()
            Sesion_1.autECLOIA.WaitForInputReady()
            #Select the option where the program will change the Client Status
            Sesion_1.autECLPS.SendKeys("1", 17, 23)
            Sesion_1.autECLOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[enter]")       
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[pf11]")        
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("2")        
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[enter]")       
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            #Add a comment on the client page dating the day of close for debt
            Sesion_1.autECLPS.SendKeys("[pf8]")        
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("CUENTA CERRADA POR DEUDA")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[enter]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            #Exist the client back to the client access Screen
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()        
            #Check if a pop-up window appears if the client cannot buy every product.            
            if Sesion_1.autECLPS.GetTex(4, 1,80).find("Uni") !=-1:            
                Sesion_1.autECLIOIA.WaitForAppAvailable()
                Sesion_1.autECLIOIA.WaitForInputReady()
                Sesion_1.autECLPS.SendKeys("[pf12]")  
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
            Sesion_1.autECLPS.SendKeys("[pf12]")
            Sesion_1.autECLIOIA.WaitForAppAvailable()
            Sesion_1.autECLIOIA.WaitForInputReady()
        else:
            _show_warning("No está en el módulo correcto.\nVaya a la opción 1217.\nYvuelva a ejecutar el programa")
    except Exception as e:
        logging.error(f"An error occurred in Abrir_Deuda: {e}")

