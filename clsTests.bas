Option Explicit

Private Type TTesting

    name As String
    iSuccess As Long
    iFail As Long
    testsCount As Long

End Type

Private this As TTesting

'---------------------------------
'           Properties
'---------------------------------

Public Property Get iSuccess() As Long: iSuccess = this.iSuccess: End Property
Public Property Get iFail() As Long: iFail = this.iFail: End Property

'---------------------------------
'           Setup
'---------------------------------

Private Sub Class_Initialize()
    
    this.testsCount = 1
    this.iSuccess = 0
    this.iFail = 0
    
End Sub

Public Function Create(name As String) As clsTests
    
    this.name = name
    Debug.Print "#" & this.name
    
    Set Create = Me
    
End Function


'---------------------------------
'           EndReport
'---------------------------------

Public Function GetEndReport(Optional testColl As Collection) As String
    
    'Add up all tests
    Dim t As clsTests
    If Not testColl Is Nothing Then
        
        For Each t In testColl
            Call AddTestResults(t)
        Next t
        
    End If
    
    'Create string
    Dim msg As String
        msg = "--------------------------------"
        msg = msg & vbNewLine & vbTab & this.name
        msg = msg & vbNewLine & "--------------------------------"
        msg = msg & vbNewLine & "TEST OBJS:  " & this.testsCount
        msg = msg & vbNewLine & "TESTS RUN:  " & this.iSuccess + this.iFail
        msg = msg & vbNewLine & "SUCCESSES:  " & this.iSuccess
        msg = msg & vbNewLine & "FAILURES :  " & this.iFail
        msg = msg & vbNewLine & "--------------------------------"

    GetEndReport = msg
    
End Function

Private Sub AddTestResults(t As clsTests)

    this.testsCount = this.testsCount + 1
    this.iSuccess = this.iSuccess + t.iSuccess
    this.iFail = this.iFail + t.iFail

End Sub


'---------------------------------
'           Asserts
'---------------------------------

Public Sub AssertTrue(val As Boolean)
    If val Then Success ("is true") Else Fail ("is false")
End Sub

Public Sub AssertFalse(val As Boolean)
    If Not val Then Success ("is false") Else Fail ("is true")
End Sub

Public Sub AssertEqual(val, val2)
    If val = val2 Then Success (val & " equals " & val2) Else Fail (val & " <> " & val2)
End Sub

Public Sub AssertNotEqual(val, val2)
    If val <> val2 Then Success (val & " does not equal " & val2) Else Fail (val & " = " & val2)
End Sub

Public Sub AssertIsNothing(obj)
    If obj Is Nothing Then Success ("obj is nothing") Else Fail ("obj is something")
End Sub

Public Sub AssertNotIsNothing(obj)
    If Not obj Is Nothing Then Success ("obj is not nothing") Else Fail ("obj is nothing")
End Sub

Public Sub Success(Optional postmessage As String)

    Debug.Print "       ....SUCCESS    | " & postmessage
    this.iSuccess = iSuccess + 1

End Sub

Public Sub Fail(Optional postmessage As String)

    Debug.Print "       ....FAILED     | " & postmessage
    this.iFail = iFail + 1
    
End Sub


