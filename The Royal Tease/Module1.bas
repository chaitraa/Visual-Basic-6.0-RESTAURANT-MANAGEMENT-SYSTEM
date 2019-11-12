Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public tRS As New ADODB.Recordset

Public I, J, K As Long



Function CheckNum(KeyNum)
If KeyNum = 8 Then CheckNum = KeyNum: Exit Function
If KeyNum < 46 Or KeyNum > 57 Then
CheckNum = 0
MsgBox ("Please Enter Numbers Only")
Else
CheckNum = KeyNum
End If
End Function


Function DateFormat(vdate1)
DateFormat = Format(vdate1, "dd/MMM/yyyy")
End Function



