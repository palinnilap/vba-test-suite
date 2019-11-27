Attribute VB_Name = "tests_Utils"
Option Explicit

Private wb As Workbook

Public Function Main_Utils() As clsTests

    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("Main_Utils")
    Set Main_Utils = test
    Dim testsColl As New Collection
    
    testsColl.Add test_Quicksort2d(test2d)
    testsColl.Add test_GetRecordDict
    testsColl.Add test_CloneDict
    testsColl.Add test_GetHeaders
    testsColl.Add test_RowToDict
    testsColl.Add test_GetFirstColRg
    
    Debug.Print test.GetEndReport(testsColl)

End Function

Private Function getfakeworksheet()

    Dim ws As Worksheet
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    
    ws.Range("A1").Value2 = "TestCol1"
    ws.Range("c1").Value2 = "TestCol With Spaces"
    
    ws.Range("A2").Value2 = "TestVal1"
    ws.Range("B2").Value2 = "TestVal2"
    ws.Range("c2").Value2 = ""
    
    Set getfakeworksheet = ws
    
End Function

Private Sub KillWb()
    
    wb.Close False
    
End Sub

Public Function test_GetFirstColRg()
    
    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_GetFirstColRg")
    Set test_GetFirstColRg = test
    
    Dim ws As Worksheet
    
    Set ws = getfakeworksheet
    
    Dim r As Range
    Set r = Utils.GetFirstColRg(ws)
    
    test.AssertEqual r.count, 2
    
    KillWb
    
End Function

Public Function test_RowToDict()
    
    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_RowToDict")
    Set test_RowToDict = test
    
    Dim ws As Worksheet
    
    Set ws = getfakeworksheet
    
    Dim r As Range
    Set r = Utils.GetHeaders(ws)
    
    Dim dict As Object
    Set dict = RowToDict(r, 2)
    
    test.AssertEqual dict("TestCol1"), "TestVal1"
    test.AssertEqual dict(""), "TestVal2"
    test.AssertEqual dict("TestCol With Spaces"), ""
    
    KillWb
    
End Function

Public Function test_GetHeaders()
    
    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_GetHeaders")
    Set test_GetHeaders = test
    
    Dim ws As Worksheet
    
    Set ws = getfakeworksheet
    
    Dim r As Range
    Set r = Utils.GetHeaders(ws)
    
    test.AssertEqual r.count, 3
    
    KillWb
    
End Function

Public Function test_CloneDict()
    
    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_CloneDict")
    Set test_CloneDict = test
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "ID", 1
    dict.Add "City", "Akron"
    dict.Add "C2_ID", Null
    dict.Add "C3_ID", ""
    
    test.AssertEqual 1, dict("ID")
    test.AssertTrue IsNull(dict("C2_ID"))
    test.AssertEqual "", dict("C3_ID")
    
    Dim newDict As Object
    Set newDict = Utils.CloneDict(dict)
    
    test.AssertEqual 1, newDict("ID")
    test.AssertTrue IsNull(newDict("C2_ID"))
    test.AssertEqual "", newDict("C3_ID")
    
End Function

Public Function test2d()

    Dim arr(0 To 2, 0 To 1) As Variant
    
    arr(0, 0) = 1
    arr(0, 1) = "Zulu"
    arr(1, 0) = 2
    arr(1, 1) = "Alpha"
    arr(2, 0) = 3
    arr(2, 1) = "Charley"
    
    test2d = arr
    
End Function

Public Function test_Quicksort2d(arr)
    
    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_Quicksort2d")
    Set test_Quicksort2d = test
    
    Dim result As Variant
    
    result = QuickSort2D(arr, LBound(arr, 1), UBound(arr, 1))
    
    Call test.AssertEqual(result(0, 0), 2)
    Call test.AssertEqual(result(0, 1), "Alpha")
    Call test.AssertEqual(result(2, 0), 1)
    Call test.AssertEqual(result(2, 1), "Zulu")
    
End Function

Public Function test_DictToArr()

    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_DictToArr")
    Set test_DictToArr = test
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add 1, "Zulu"
    dict.Add 2, "Alpha"
    dict.Add 3, "Charley"
    
    test.AssertEqual dict(1), "Zulu"
    
    Dim result As Variant
    
    result = DictToArr(dict)
    
    Call test.AssertEqual(result(0, 0), 1)
    Call test.AssertEqual(result(0, 1), "Zulu")
    Call test.AssertEqual(result(2, 0), 3)
    Call test.AssertEqual(result(2, 1), "Charley")
    
    Call test_Quicksort2d(result)
    
End Function

Public Function test_clsErrHandlerEmail()

    Call CreateGlobalLogger
    Dim test As clsTests: Set test = mFactories.CreateTest("test_clsErrHandlerEmail")
    Set test_clsErrHandlerEmail = test
    
    Call CreateGlobalLogger
    On Error GoTo errH
    Err.Raise 112
    
    Exit Function
    
errH:
    logger.Critical ("test_clsErrHandlerEmail")
    
End Function


