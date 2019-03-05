Attribute VB_Name = "ModConnection"
Public Function GetConnection2()
On Error GoTo kate
Dim Dbase1 As Connection
Set Dbase1 = New Connection
Dbase1.CursorLocation = adUseClient
Dbase1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ProjectStaff.mdb.mdb;"
Exit Function
kate:
MsgBox Err.Description
End Function

Public Function GetConnection1()
On Error GoTo kate
Dim Dbase2 As Connection
Set Dbase2 = New Connection
Dbase2.CursorLocation = adUseClient
Dbase2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ProjectStudents.mdb;"
Exit Function
kate:
MsgBox Err.Description
End Function

Public Function GetConnection3()
On Error GoTo kate
Dim Dbase3 As Connection
Set Dbase3 = New Connection
Dbase3.CursorLocation = adUseClient
Dbase3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ProLecturers.mdb;"
Exit Function
kate:
MsgBox Err.Description
End Function


