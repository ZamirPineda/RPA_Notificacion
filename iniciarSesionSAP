Sub Iniciar_Sap()
    'Limpia mensajes de inicio de sesión en SAP
    
    Dim Application
    
    
    On Error Resume Next
    
    'Iniciar sesión automáitcamente
    Set SapGui = GetObject("SAPGUI")
    Set Appl = SapGui.GetScriptingEngine
    Set Connection = Appl.openconnection("1. Grupo Nutresa_ERP_PRD", True) ' Modulo de sap al que se logea
    Set session = Conecction.Children(0)
    
    'Limpieza de errores y mensajes
    session.findById("wnd[0]").maximize
    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    

End Sub
