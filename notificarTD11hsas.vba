'Script Creado por Zamir Pineda
Sub Notificacion_TD_11_Hsas()

'Declarar variables
Dim Application
Dim Connection
Dim Filas As Long
Dim i As Long
Filas = ThisWorkbook.Sheets("BaseHambu").Range("E2").CurrentRegion.Rows.Count
For i = 2 To Filas

'Conectar con SAP
If Not IsObject(Application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Application.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

'Limpiar SAP
session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]").resizeWorkingPane 198, 32, False
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
'Notificación TD 11
session.findById("wnd[0]/tbar[0]/okcd").text = "COR6N"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-AUFNR").text = Sheets("BaseHambu").Cells(i, 5).Value
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").text = "11"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").SetFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").caretPosition = 2
session.findById("wnd[0]").sendVKey 11
Next i

MsgBox "Datos enviados correctamente"

End Sub
