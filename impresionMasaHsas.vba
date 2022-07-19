'Script Creado por Zamir Pineda
Sub Impresion_En_Masa_HSAS()

'Declarar variables
Dim Application
Dim Connection
Dim Filas As Long
Dim i As Long
Filas = ThisWorkbook.Sheets("BaseHambu").Range("X2:Y2").CurrentRegion.Rows.Count
For i = 2 To Filas

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

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "ZPP_POM_2446_1"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtZEPP_2446_1-AUFNR").text = Sheets("BaseHambu").Cells(i, 24).Value
session.findById("wnd[0]/usr/txtZEPP_2446_1-AUFNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtZEPP_2446_1-NUM_IMPRESIONES").text = Sheets("BaseHambu").Cells(i, 25).Value
session.findById("wnd[0]/usr/txtZEPP_2446_1-CONSC_MUESTRAS").text = "1"
session.findById("wnd[0]/usr/txtZEPP_2446_1-CONSC_MUESTRAS").SetFocus
session.findById("wnd[0]/usr/txtZEPP_2446_1-CONSC_MUESTRAS").caretPosition = 1
session.findById("wnd[0]/usr/btnBTN_IMP_ETIQUETA").press
session.findById("wnd[1]/usr/ctxtSSFPP-TDDEST").text = "ZAC5711035I"
session.findById("wnd[1]/usr/ctxtSSFPP-TDDEST").caretPosition = 11
session.findById("wnd[1]/tbar[0]/btn[86]").press

Next i


End Sub
