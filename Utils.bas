Attribute VB_Name = "Utils"

Option Explicit

'author: palinnilap@protonmail.com

Public Function GetHeaders(ws As Worksheet) As Range

    Set GetHeaders = _
        Range( _
            ws.Range("A1"), _
            ws.Range("DZ1").End(xlToLeft) _
        )
        
End Function

Public Sub RaisesErrorIfZero(num As Long)

    If num = 0 Then
        Err.Raise 9999, , "RaisesErrorIfZero cannot receive 0 as the parameter"
    End If
    
End Sub