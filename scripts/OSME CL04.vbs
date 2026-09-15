If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIQ09"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 1500

session.findById("wnd[0]/usr/txtSERNR-LOW").Text = "{{SERIAL}}"
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2000

On Error Resume Next
session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06").Select
If Err.Number <> 0 Then
    WScript.Echo ""
    WScript.Quit
End If
On Error Goto 0
WScript.Sleep 1000

On Error Resume Next
session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\06/ssubSUB_DATA:SAPLITO0:0122/subSUB_0122B:SAPLITO0:1221/btn%_AUTOTEXT002").press
If Err.Number <> 0 Then
    WScript.Echo ""
    WScript.Quit
End If
On Error Goto 0
WScript.Sleep 1000

Set tree = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell")
Set coll = tree.GetAllNodeKeys()
ordem = ""
For x = 1 To coll.Length - 1
    key = coll.ElementAt(x)
    num = tree.GetItemText(key, "2")
    if Left(num, 4) = "1000" And Len(num) = 12 Then
        ordem = num
        Exit For
    End If
Next
WScript.Echo ordem
