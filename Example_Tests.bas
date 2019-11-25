Attribute VB_Name = "tests_Utils"

Option Explicit

'a module level variable makes the mock easier to work with
Private wb As Workbook

Public Function Main_Utils() As clsTests

    'PURPOSE: run all tests in this module in one click 
    '         and report results into debug window

    'setup test class
    Dim test As New clsTests
    Set test = test.create("Main_Utils")
    Set Main_Utils = test

    'create a collection of all tests to run
    Dim testsColl As New Collection

    testsColl.Add test_GetHeaders
    testsColl.Add test_RaisesErrorIfZero
    
    Debug.Print test.GetEndReport(testsColl)

End Function


Public Function test_GetHeaders() as clsTests
    
    'setup test class
    Dim test As New clsTests
    Set test = test.create("test_GetHeaders")
    Set test_GetHeaders = test
    
    Dim ws As Worksheet
    Dim r As Range

    Set ws = getfakeworksheet
    Set r = Utils.GetHeaders(ws)
    
    test.AssertEqual r.count, 3
    
    KillWb
    
End Function

Public Function test_RaisesErrorIfZero() As clsTests

    'setup test class
    Dim test As New clsTests
    Set test = test.Create("test_RaisesErrorIfZero")
    Set test_RaisesErrorIfZero = test
    
    'Test
    On Error GoTo errH

    Call Utils.RaisesErrorIfZero(0)
    
    'This part of the code will not execute if the error was raised
    test.Fail ("did not raise error successfully")
    Exit Function
    
errH:
    On Error GoTo -1
    test.Success ("error raised successfully")

End Function

'--------------------------
'	Mock Functions
'--------------------------

Private Function getfakeworksheet()

    Dim ws As Worksheet
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    
    ws.Range("A1").Value2 = "TestCol1"
    ws.Range("A1").Value2 = ""   'test if this works when some headers are blank
    ws.Range("c1").Value2 = "TestCol With Spaces"
    
    ws.Range("A2").Value2 = "TestVal1"
    ws.Range("B2").Value2 = "TestVal2"
    ws.Range("c2").Value2 = ""
    
    Set getfakeworksheet = ws
    
End Function

Private Sub KillWb()
    
    wb.Close False
    
End Sub

