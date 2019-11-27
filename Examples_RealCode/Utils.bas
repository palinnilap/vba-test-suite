Attribute VB_Name = "Utils"
Option Explicit

'Version 23 - getheaders and getrowdict

Public Function QuickSort2D(arr2D As Variant, first As Long, last As Long)
  
  'PURPOSE: Takes a two column array [zero-indexed] and sorts by the second column
  
  Dim centVal As Variant
  Dim vTemp(0, 0 To 1) As Variant
  
  Dim tempLow As Long
  Dim tempHi As Long
  tempLow = first
  tempHi = last
  
  centVal = LCase(arr2D((first + last) \ 2, 1))
  Do While tempLow <= tempHi
    
    Do While LCase(arr2D(tempLow, 1)) < centVal And tempLow < last
      tempLow = tempLow + 1
    Loop
    
    Do While centVal < LCase(arr2D(tempHi, 1)) And tempHi > first
      tempHi = tempHi - 1
    Loop
    
    If tempLow <= tempHi Then
    
        ' Swap values
        vTemp(0, 0) = arr2D(tempLow, 0)
        vTemp(0, 1) = arr2D(tempLow, 1)
        
        arr2D(tempLow, 0) = arr2D(tempHi, 0)
        arr2D(tempLow, 1) = arr2D(tempHi, 1)
        
        arr2D(tempHi, 0) = vTemp(0, 0)
        arr2D(tempHi, 1) = vTemp(0, 1)
        
        ' Move to next positions
        tempLow = tempLow + 1
        tempHi = tempHi - 1
      
    End If
    
  Loop
  
  If first < tempHi Then QuickSort2D arr2D, first, tempHi
  If tempLow < last Then QuickSort2D arr2D, tempLow, last

  QuickSort2D = arr2D
  
End Function

'------------------------------
'      Misc Functions
'------------------------------

Public Function GetUniqueSaveName(desiredFileName As String)

    'PURPOSE: if needed, will append an integer to make savename unique

    Dim count As Integer
    Dim newName As String
    Dim offset As Long
    
    newName = desiredFileName
    count = 0
    
    'if name already exists, will add numbers until a unique name is made
    Do While WorkbookExist(newName) Or IsWorkBookOpen(newName)
        
        If count = 0 Then
            offset = 0
        Else
            offset = Len(CStr(count))
        End If
        
        count = count + 1
        
        Dim extStart As Long
        extStart = InStrRev(newName, ".")
        
        newName = Left(newName, extStart - 1 - offset) & count & Mid(newName, extStart)
        
        If count > 100 Then
            Err.Raise 1111, , "GetUniqueSaveName tried to find a unique" _
                & " name for the following, but 100 versions already existed: " _
                & desiredFileName
        End If
        
    Loop

    GetUniqueSaveName = newName
    
End Function

Public Function IntelliTab(str As String, length As Long) As String

    str = str & "                              "
    str = str & "                              "
    str = str & "                              "
    str = str & "                              "
    
    IntelliTab = Left(str, length)

End Function

Function ConfirmMsgBox( _
                Optional msg As String = "Run Macro?", _
                Optional boxTitle As String = "Confirm Action" _
                ) As Boolean
    
    Dim choice As VbMsgBoxResult
    choice = MsgBox(msg, vbOKCancel + vbQuestion, boxTitle)
    
    If choice = vbOK Then
        ConfirmMsgBox = True
    Else
        ConfirmMsgBox = False
    End If
    
End Function

Public Function ValidateFolderAddressFormat(folderAddress As String, Optional termBackSlash As Boolean = True) As String

    Dim backslash As Boolean
    If Right(folderAddress, 1) = "\" Then
        backslash = True
    End If

    'modify string as needed
    If termBackSlash And Not backslash Then
        ValidateFolderAddressFormat = folderAddress & "\"
        
    ElseIf Not termBackSlash And backslash Then
        ValidateFolderAddressFormat = Left(folderAddress, Len(folderAddress) - 1)
        
    Else
        ValidateFolderAddressFormat = folderAddress
            
    End If

End Function

Public Function JustFilename(f)
    'takes fullname and returns only file name (with ext)
    JustFilename = Mid(f, InStrRev(f, "\") + 1)

End Function

Public Function CloneDict(d As Object) As Object

  'https://stackoverflow.com/questions/3022182/how-do-i-clone-a-dictionary-object
  
  Dim newDict
  Set newDict = CreateObject("Scripting.Dictionary")

  Dim key As Variant
  For Each key In d.keys
    newDict.Add key, d(key)
  Next
  
  newDict.CompareMode = d.CompareMode

  Set CloneDict = newDict
  
End Function

'-----------------------------
'       Worksheet Tools
'-----------------------------

Public Function FindCell(searcharea As Range, searchPhrase, Optional exact As Boolean = False) As Range
    
    'case insensitive
    'finds any cell containing the search phrase
    'converts whatever is in the cell to a string
    'if found, returns cell
    'if not found, returns nothing
    'example code: "If Findcell(...) is nothing then "
    
    Call CreateGlobalLogger
    'Call logger.Dbug("Utils.FindCell Started, ws:" & searcharea.Parent.name & "; phrase: " & searchPhrase)
    
    Dim r As Range
    Dim rText As String
    searchPhrase = CStr(searchPhrase)
    
    For Each r In searcharea

        'Convert r.value2 to string
        rText = cText(r)

        'Exact Match
        If Not exact Then
        
            If LCase(rText) Like "*" & LCase(searchPhrase) & "*" Then
                Set FindCell = r
                Exit For
            End If
        
        ElseIf exact Then
            
            If rText = searchPhrase Then
                Set FindCell = r
                Exit For
            End If
            
        End If
        
    Next r

    'Call logger.Dbug("Utils.FindCell found something: " & (Not r Is Nothing))

End Function

Public Function FindAllCells( _
        searcharea As Range _
        , searchPhrase _
        , Optional exactMatch As Boolean = False _
        ) As Collection
    
    'case insensitive (unless exact)
    'finds any cell containing the search phrase (unless exact)
    'converts whatever is in the cell to a string
    'if found, returns a collection of ranges
    'if not found, returns nothing
    'example code: "If Findcell(...) is nothing then "
    
    Dim r As Range
    Dim rText As String
    Dim hits As New Collection
    
    searchPhrase = CStr(searchPhrase)
    
    For Each r In searcharea
        
        'convert to string
        rText = cText(r)

        'search exactMatch on/off
        If exactMatch Then
            
            If rText = searchPhrase Then
                hits.Add r
            End If
            
        ElseIf Not exactMatch Then
        
            If LCase(rText) Like "*" & LCase(searchPhrase) & "*" Then
                hits.Add r
            End If
        End If
        
    Next r
    
    'Return nothing if no hits
    If hits.count = 0 Then
        Set FindAllCells = Nothing
    Else
        Set FindAllCells = hits
    End If
    
End Function

Public Function cText(r As Range) As String

    Dim rText As String
    
    If r = Empty Then
        rText = ""
    ElseIf r.NumberFormat = "m/d/yyyy" Then
        rText = Format(CDate(r.Value2), "m/d/yyyy")
    ElseIf Not Application.WorksheetFunction.IsText(r) Then
        rText = CStr(r.Value2)
    Else
        rText = r.Value2
    End If

    cText = rText

End Function

Public Function GetHeaders(ws As Worksheet) As Range

    Set GetHeaders = _
        Range( _
            ws.Range("A1"), _
            ws.Range("DZ1").End(xlToLeft) _
        )
        
End Function

Public Function GetFirstColRg(ws As Worksheet) As Range

    Set GetFirstColRg = _
        Range( _
            ws.Range("A1"), _
            ws.Range("A1000000").End(xlUp) _
        )
        
End Function

Public Function RowToDict(headers As Range, rowNum As Long) As Object

    Dim dict As Object
    Dim r As Range
    Dim ws As Worksheet
    
    Set dict = CreateObject("Scripting.dictionary")
    Set ws = headers.Parent
    
    For Each r In headers
        dict.Add r.Value2, ws.Cells(rowNum, r.Column).Value2
    Next
    
    Set RowToDict = dict

End Function

'------------------------------
'      Error Subroutines
'------------------------------

Sub ContinueOrAbort(promptMsg As String, errNum As Long, errMsg As String)

    Dim choice As VbMsgBoxResult
    
    choice = MsgBox( _
        Prompt:=promptMsg _
        , Buttons:=vbYesNo + vbExclamation _
        , Title:="Error" _
        ) _

    If choice = vbNo Then
        Err.Raise errNum, , errMsg
    End If
        
End Sub

Public Sub PrintErrH()

    Dim t As String
    Dim modName As String
    
    modName = Application.VBE.ActiveCodePane.CodeModule.name
    
    t = vbTab & "On Error GoTo ErrH"
    t = t & vbNewLine & vbTab & "dim logger as clslogger: "
    t = t & "set logger = mFactories.CreateLogger(IMMEDIATE_LEVEL, TXTLOG_LEVEL)"
    t = t & vbNewLine & vbTab & "call logger.dbug(""" & modName & ". START"")"
    t = t & vbNewLine
    t = t & vbNewLine & "exitHere:"
    t = t & vbNewLine & vbTab & "call logger.dbug(""" & modName & ". END"")"
    t = t & vbNewLine & vbTab & "exit sub"
    t = t & vbNewLine & "errH:"
    t = t & vbNewLine & vbTab & "call logger.critical(""" & modName & "."")"
    t = t & vbNewLine & vbTab & "Resume exitHere"
    
    Debug.Print t

End Sub

'------------------------------
'      Arrays
'------------------------------

Public Function LoadArray(ws As Worksheet) As Variant

    LoadArray = GetRangeAllData(ws).Value2

    'in array rows are dimension 1 and columns are dimension 2
End Function

Public Function GetRangeAllData(ws As Worksheet) As Range

    If ws Is Nothing Then
        Err.Raise 6667, , "RangeAllData failed because ws was nothing"
    End If

    Set GetRangeAllData _
        = Range( _
            ws.Cells(1, ws.Columns.count).End(xlToLeft) _
            , ws.Cells(ws.Rows.count, 1).End(xlUp) _
        )
    
End Function

Public Sub UnloadArray(arr As Variant, intoWs As Worksheet)

    Dim i As Long
    Dim j As Long
    
    intoWs.Cells.Clear
    
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        For j = LBound(arr, 2) To UBound(arr, 2)
        
            intoWs.Cells(i, j).Value2 = arr(i, j)
        
        Next j
    
    Next i
    
End Sub


Public Sub SaveMacros()

    'PURPOSE: Exports a copy of each of this workbook's VBComponent code into separate .bas files
    'note 1: exports all as '.bas'. They still import as the correct type
    
    Dim i As Integer, path As String, fname As String

    path = "C:\Users\" & Environ("Username") & "\Desktop\CodeModules"

    Call FolderExist(path, autocreate:=True)

    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.count
            If .VBComponents(i).CodeModule.CountOfLines > 0 Then
                fname = path & "\"
                fname = fname & Format(Now(), "yyyy-mm-dd_hh.mm_")
                fname = fname & UCase(Left(ThisWorkbook.name, 6)) & "_"
                fname = fname & .VBComponents(i).CodeModule.name
                fname = fname & ".bas"
                .VBComponents(i).Export fname
            End If
        Next i
    End With

End Sub

Public Sub UpdateMacroLinks(ShapesToChange As Collection)
    'PURPOSE: Remove an external workbook reference from all shapes triggering macros

    Dim shp As Shape
    Dim MacroLink As String
    Dim SplitLink As Variant
    Dim NewLink As String
    
    'Loop through each shape in worksheet
      For Each shp In ShapesToChange
      
        'Grab current macro link (if available)
          MacroLink = shp.OnAction
        
        'Determine if shape was linking to a macro
          If MacroLink <> "" And InStr(MacroLink, "!") <> 0 Then
            'Split Macro Link at the exclaimation mark (store in Array)
              SplitLink = Split(MacroLink, "!")
            
            'Pull text occurring after exclaimation mark
              NewLink = SplitLink(1)
            
            'Remove any straggling apostrophes from workbook name
                If Right(NewLink, 1) = "'" Then
                  NewLink = Left(NewLink, Len(NewLink) - 1)
                End If
            
            'Apply New Link
              shp.OnAction = NewLink
          End If
      
      Next shp
      
    Exit Sub
    
End Sub

Public Sub LineCount()
    
    'PURPOSE: prints number of VBComponents and number of lines
    
    Dim i As Integer
    Dim x As Long
    Dim total As Long
    Dim tabs As String

    total = 0
    
    With ThisWorkbook.VBProject
        For i = 1 To .VBComponents.count
        
            x = .VBComponents(i).CodeModule.CountOfLines
           
                           
'                If Len(.VBComponents(i).name) < 10 Then
'                    tabs = vbTab & vbTab & vbTab & vbTab
'                ElseIf Len(.VBComponents(i).name) < 17 Then
'                    tabs = vbTab & vbTab
'                Else
'                    tabs = "  "
'                End If
                
            Debug.Print IntelliTab(.VBComponents(i).name, 25) & CStr(x)
            
            total = total + x
        Next i
    End With
    
    Debug.Print "-----------------"
    Debug.Print "# of Objects:       " & ThisWorkbook.VBProject.VBComponents.count
    Debug.Print "LINES GRAND TOTAL:  " & total
    
End Sub

'------------------------------
'      Exist Functions
'------------------------------

Public Function WorksheetExist(WorksheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExist = (.Sheets(WorksheetName).name = WorksheetName)
        On Error GoTo 0
    End With
End Function

Public Function IsWorkBookOpen(filename As String) As Boolean
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open filename For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: IsWorkBookOpen = False
    End Select
End Function

Public Function WorkbookExist(filename As String) As Boolean
    
    If Dir(filename) = "" Then
        WorkbookExist = False
    Else
        WorkbookExist = True
    End If
    
End Function

Public Function FolderExist(path As String, Optional autocreate As Boolean = False) As Boolean

    'PURPOSE: if folder doesn't exist, gives user choice to create folder or exit
    'when autocreate=true, does not prompt user

    Dim Folder As String
    Dim Answer As VbMsgBoxResult

    Folder = Dir(path, vbDirectory)
 
    If Folder = vbNullString And Not autocreate Then
        Answer = MsgBox("'" & path & "\' does not exist. Would you like to create it?", vbYesNo, "Create Path?")
        If Answer = vbYes Then
            VBA.FileSystem.MkDir (path)
        Else
            FolderExist = False
            Exit Function
        End If
        
    ElseIf Folder = vbNullString And autocreate Then
        VBA.FileSystem.MkDir (path)
        
    End If
    
    FolderExist = True

End Function




