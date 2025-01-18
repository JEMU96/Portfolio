# Portfolio
This is a portfolio to showcase some small programs, in order to demonstrate the basics of my capabilities.

GUI.py is the main module, the one that runs and starts the program.
  class ShakingButton -> A small animation for the buttons in the UI; a simple move left and right when the cursor enters the buttons.
  class ProgramSelector -> UI where the user can select the module to be imported dynamically that has the name with which other programs must interact (SAP, WEB, IBM). It also has a progress bar that will be updated whenever the program selection takes a step.
  function dynamic_import_module -> A simple function to import a module by name.
  function show_warning -> A simple function to display a warning message. (Unused in this module)
  function Excel_open_read -> This function asks the user to select an Excel file via a pop-up window. Using the pandas library, the program reads the Excel file and passes the rows as arguments for the selected program.

IBM.py: Whenever selected in the UI, this module will be imported and show the user the two non-private functions to choose from in order to run the program.
  function _launch_pcomm_session -> This function is used to launch IBM PCCOM session if needed.
  function _connect_to_session -> This function tries to connect to session 'A' of the IBM PCCOM if the user has it open.
  function ABIR_DEUDA -> This function interacts with the IBM PCCOM in order to change the status of a client from not billable to billable.
  function CERRAR_DEUDA -> This function interacts with the IBM PCCOM in order to change the status of a client from billable to not billable.

WEB.py: Whenever selected in the UI, this module will be imported and show the user the two non-private functions to choose from in order to run the program.
  function _Get_login_credentials -> A simple UI pop-up to log in if the user was not already logged in. The user can cancel the login form from the pop-up, do it manually, and continue with the program.
  function _logged -> A simple function to enter the user credentials from _Get_login_credentials into the web.
  function _login_loop -> A simple function to ensure the user is logged in to continue with the program.
  function _Input_Text -> A simple UI pop-up to ask the user for any input.
  function _Conn_open_driver -> A simple function to connect to the Chrome driver. ASSUMED USER HAS CHROME INSTALLED.
  functions #Navigate the web -> A list of functions to navigate and interact with the elements of the web.
  function ABIR_DEUDA -> This function interacts with the IBM PCCOM in order to change the status of a client from not billable to billable.
  function CERRAR_DEUDA -> This function interacts with the IBM PCCOM in order to change the status of a client from billable to not billable.

SAP.py: Whenever selected in the UI, this module will be imported and show the user the non-private functions to choose from in order to run the program. SAP has a lot of sensitive data, so I only present one of them, making sure there is no personal information.
  function _Connect_to_SAP -> Connects to the scripting engine of SAP "SapGuiAuto".
  function Preparar_Pago -> Asks the user to open two Excel files, operates on them, and after confirming some data in SAP, calls _Aplicar_Pago or _Aplicar_Dif.
  function _Aplicar_Pago -> Enters the final data into SAP and saves the file with the name of the "accounting entry" to save a copy.
  function _Aplicar_Dif -> Enters the differential data depending on the amount on different accounts and continues with _Aplicar_Pago.

Picture Credits:
  -Green Thumb up (continue_icon.png): https://www.cleanpng.com/png-thumb-signal-computer-icons-emoji-clip-art-thumbs-736308/
  -Red Thumb down (cancel_icon.png): https://www.cleanpng.com/png-thumb-signal-emoticon-clip-art-thumbs-down-clipart-187757/
  -Black Thumb up(window_icon.png): https://es.wikipedia.org/wiki/Archivo:Thumb_up_icon_2.svg
