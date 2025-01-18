# -*- coding: utf-8 -*-
"""
@author: JEMU

MODULO AUXILIAR PARA LOS PROGRAMAS QUE SE BASAN EN AUTOMATIZAR OBJETOS
QUE SE EMPEZARON A DESARROLLAR EN VISUAL BASIC POR LAS INTERACCIONES QUE AHI ENTRE PROGRAMAS
Y SE HA MODIFICADO A PYTHON PARA UNIFICARLOS TODOS EN UNA UI  PARA EL USUARIO.

*UNICO PROGRAMA YA QUE SAP TIENE MUCHOS DATOS SENSIBLES QUE HAY QUE OCULTAR.

"""

import sys
import os
import pandas as pd
import win32com.client
from PyQt5.QtWidgets import QApplication, QFileDialog, QInputDialog, QMessageBox

def _Connect_to_SAP():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        app = SapGuiAuto.GetScriptingEngine
    except:
        app = None

    if app is not None:
        try:
            connection = app.Children(0)
        except:
            connection = None

        if connection is not None:
            try:
                session = connection.Children(0)
            except:
                session = None

    return app, connection, session

def _show_warning(message):
    popup = QApplication.instance() if QApplication.instance() else QApplication(sys.argv)
    QMessageBox.warning(None, 'Warning', message)
    popup.exec_()

def _get_file_path(dialog_title):
    popup = QApplication.instance() if QApplication.instance() else QApplication(sys.argv)
    file_dialog = QFileDialog()
    file_dialog.setWindowTitle(dialog_title)
    file_path, _ = file_dialog.getOpenFileName(file_dialog, dialog_title, "", "Excel Files (*.xlsx;*.xls)")
    return file_path

def _get_user_input(dialog_title):
    popup = QApplication.instance() if QApplication.instance() else QApplication(sys.argv)
    value, ok = QInputDialog.getDouble(None, dialog_title, dialog_title)
    if ok:
        return value
    else:
        return None

def _show_confirmation(message):
    popup = QApplication.instance() if QApplication.instance() else QApplication(sys.argv)
    reply = QMessageBox.question(None, 'Confirmation', message, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
    return reply == QMessageBox.Yes

def Preparar_Pago():
    # Connect to SAP
    app, connection, session = _Connect_to_SAP()
    #Ask User to open files
    fichero = _get_file_path("Abre la relación de pago")
    plantilla = _get_file_path("Abre la Plantilla Call Transaction") #Excel template to use as an Batch input.

    if not fichero or not plantilla:
        _show_warning("No se ha seleccionado fichero. Se cancela el proceso")
        return

    # Open Excel files
    wbR = pd.ExcelFile(fichero)
    wsR = wbR.parse(wbR.sheet_names[0])
    wbP = pd.ExcelFile(plantilla)
    wsP = wbP.parse(wbP.sheet_names[0])

    # Perform operations on Excel files
    wsR.columns = wsR.columns.str.strip().str.replace('/', '.').str.replace('-', '')

    importe_real = _get_user_input("Introduce el total del Pagaré")
    if importe_real is None:
        _show_warning("No se ha introducido el total del Pagaré. Se cancela el proceso")
        return

    importe_relacion = float(wsR.iloc[7, 1])

    if importe_real == importe_relacion:
        final_row = wsR.iloc[:, 1].last_valid_index() + 1
        inicial_row = 10
        FPI = 10
        FPF = wsP.iloc[:, 3].last_valid_index() + 1 if wsP.iloc[:, 3].notna().any() else 10

        if FPI < FPF:
            wsP.iloc[9:FPF, 3] = ''

        wsR.iloc[:, 1] = wsR.iloc[:, 1].astype(str).str.replace('/', '.').str.replace('-', '')
        vencimiento = wsR.iloc[7, 3]
        asignacion = vencimiento[-4:] + vencimiento[3:5] + vencimiento[:2]

        date = pd.Timestamp.today().strftime('%d.%m.%Y')
        NPag = wsR.iloc[7, 0]
        cargos, abonos, facturas = 0, 0, 0

        dic_fusis = {}

        for i in range(inicial_row, final_row):
            valorC = wsR.iloc[i, 3]
            if len(wsR.iloc[i, 1]) == 7:
                if wsR.iloc[i, 1][0] == '4':
                    wsP.at[FPI, 3] = "X" + wsR.iloc[i, 1]
                    FPI += 1
                    facturas += valorC
                elif wsR.iloc[i, 1][0] in ['5', '6', '7']:
                    wsP.at[FPI, 3] = "V" + wsR.iloc[i, 1]
                    FPI += 1
                    facturas += valorC
            elif wsR.iloc[i, 1][0] == 'C' and valorC < 0:
                cargos += valorC
            elif wsR.iloc[i, 1][0] in ['C', 'A'] and valorC > 0:
                abonos += valorC
            else:
                dic_fusis[i] = i

        wsP.at[1, 4] = date
        wsP.at[1, 6] = date
        wbP.save()
        wbP.close()

        cliente = "123abc" #Fake client for security reasons
        session.findById("wnd[0]/tbar[0]/okcd").text = "Batch_input"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/radP_CALLT").Select
        session.findById("wnd[0]/usr/ctxtP_FILE").text = plantilla
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
        importe_sap = session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-NETTO").DisplayedText
        importe_sap = float(importe_sap.replace(',', ''))
        facturas = float(str(facturas).replace(',', ''))
        session.findById("wnd[0]/tbar[1]/btn[14]").press()

        if importe_sap == facturas:
            _Aplicar_Pago(app, connection, session,importe_real,cliente,asignacion,
                         cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero)
        else:
            diff = importe_sap - facturas
            _show_warning(f"Hay diferencias entre las partidas y la relación de facturas. {diff}")
            confirmacion = QInputDialog.getText(None, "Confirmación", "¿Quiere ajustar por redondeo? (yes/no): ")[0]
            if confirmacion.lower() == 'yes':
                _Aplicar_Dif(diff,app, connection, session,importe_real,cliente,asignacion,
                             cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero)
            else:
                _show_warning("Se Cancela el Proceso")
                return
    else:
        _show_warning("El importe no cuadra. Se cancela el proceso")
        return


def _Aplicar_Pago(app, connection, session,importe_real,cliente,asignacion,
             cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero):
    session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "09"
    # Cliente = QInputDialog.getText(None, "Introduce el codigo del CLIENTE")[0]
    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
    # CME = QInputDialog.getText(None, "Introduce la variable de CME")[0]
    session.findById("wnd[0]/usr/ctxtRF05A-NEWUM").text = "W"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importe_real
    session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
    session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"PAG. #nombreCliente {NPag} VTO. {vencimiento}"
    session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[14]").press()
    cargos = cargos * -1

    if abonos != 0:
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "16"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = abonos
        session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
        session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"TOTAL ABONOS {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[14]").press()

    if cargos != 0:
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "06"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = cargos
        session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
        session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"TOTAL CARGOS {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[14]").press()

    nfusis = len(dic_fusis)
    matriz = list(dic_fusis.values())

    for i in range(nfusis):
        fusi_linea = matriz[i]
        v_cargo = wsR.iloc[fusi_linea, 3]
        nom_cargo = wsR.iloc[fusi_linea, 1]
        if v_cargo > 0:
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "16"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = v_cargo
            session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
            session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"CARGO {nom_cargo} COSTES OPERATIVOS"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[14]").press()
        elif v_cargo < 0:
            v_cargo = v_cargo * -1
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "06"
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = v_cargo
            session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
            session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"CARGO {nom_cargo} COSTES OPERATIVOS"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[14]").press()

    ini = int(float(format(session.findById("wnd[0]/usr/txtRF05A-ANZAZ").DisplayedText, 'g')))
    session.findById("wnd[0]/mbar/menu[0]/menu[3]").Select()  # Simular
    fin = int(float(format(session.findById("wnd[0]/usr/txtRF05A-ANZAZ").DisplayedText, 'g')))

    for j in range(ini + 1, fin):
        session.findById("wnd[0]/usr/txtRF05A-ANZAZ").SetFocus()
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[1]/usr/txt*BSEG-BUZEI").text = j
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"PAG. #nombreCliente {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[14]").press()

    session.findById("wnd[0]/mbar/menu[2]/menu[6]").Select()  # Ultimo documento
    session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"PAG. #nombreCliente {NPag} VTO. {vencimiento}"
    session.findById("wnd[0]/tbar[1]/btn[14]").press()

    confirmation = _show_confirmation("Comprueba los apuntes. ¿Quieres aplicar el pago?")
    if not confirmation:
        _show_warning("Se cancela el proceso sin aplicar el pago")
        return

    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "fb03"
    session.findById("wnd[0]").sendVKey(0)
    nombre = session.findById("wnd[0]/usr/txtRF05L-BELNR").text
    session.findById("wnd[0]/tbar[0]/btn[15]").press()

    ruta = os.path.dirname(fichero)
    wbR.to_excel(f"{ruta}\\{nombre} #nombreCliente {importe_real}.xlsx", index=False)
    _show_warning(f"Se ha aplicado el pago con nº asieno {nombre}. Y se ha guardado el archivo")

def _Aplicar_Dif(diff,app, connection, session,importe_real,cliente,asignacion,
             cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero):
    if diff < -1:
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "06"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
        session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
        session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"Dif. {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[14]").press() 
        _Aplicar_Pago(app, connection, session,importe_real,cliente,asignacion,
            cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero) 
    elif diff > 1:
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "16"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
        session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#DatoSensible"
        session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"Dif. {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        _Aplicar_Pago(app, connection, session,importe_real,cliente,asignacion,
             cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero)
    elif diff < 0:
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "40"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "#DatoSensible"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"Dif. {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[1]/usr/ctxtCOBL-GSBER").text = "#DatoSensible"
        session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").text = "#DatoSensible"
        session.findById("wnd[1]").sendVKey(0)
        _Aplicar_Pago(app, connection, session,importe_real,cliente,asignacion,
             cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero)    
    elif diff > 0:
        diff = diff * -1
        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "50"
        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "#DatoSensible"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = f"Dif. {NPag} VTO. {vencimiento}"
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[1]/usr/ctxtCOBL-GSBER").text = "#DatoSensible"
        session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").text = "#DatoSensible"
        session.findById("wnd[1]").sendVKey(0)
        _Aplicar_Pago(app, connection, session,importe_real,cliente,asignacion,
                     cargos, abonos, facturas,NPag,vencimiento,dic_fusis,wsR,wbR,fichero)
