'Script Creado por Zamir Pineda
Sub Limpiar_SAP()

'Inicio del módulo SAP

    Dim text As String
 
 'Comandos en consola de Windows CMD para iniciar SAP automáticamente
    Shell "TASKKILL /IM saplogon.exe /F"
    Application.Wait (Now + TimeValue("00:00:03"))
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
    Application.Wait (Now + TimeValue("00:00:05"))

    Call Iniciar_Sap
    Application.Wait (Now + TimeValue("00:00:05"))

End Sub
