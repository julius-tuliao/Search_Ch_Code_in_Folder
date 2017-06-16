Attribute VB_Name = "Search"
Option Explicit
'Sub SearchFolders()
''UpdatebyKutoolsforExcel20151202
'    Dim xFso As Object
'    Dim xFld As Object
'    Dim xStrSearch As String
'    Dim xStrPath As String
'    Dim xStrFile As String
'    Dim xOut As Worksheet
'    Dim xWb As Workbook
'    Dim xWk As Worksheet
'    Dim xRow As Long
'    Dim xFound As Range
'    Dim xStrAddress As String
'    Dim xFileDialog As FileDialog
'    Dim xUpdate As Boolean
'    Dim xCount As Long
'    On Error GoTo ErrHandler
'    Set xFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
'    xFileDialog.AllowMultiSelect = False
'    xFileDialog.Title = "Select a forlder"
'    If xFileDialog.Show = -1 Then
'        xStrPath = xFileDialog.SelectedItems(1)
'    End If
'    If xStrPath = "" Then Exit Sub
'    xStrSearch = "KTE"
'    xUpdate = Application.ScreenUpdating
'    Application.ScreenUpdating = False
'    Set xOut = Worksheets.Add
'    xRow = 1
'    With xOut
'        .Cells(xRow, 1) = "Workbook"
'        .Cells(xRow, 2) = "Worksheet"
'        .Cells(xRow, 3) = "Cell"
'        .Cells(xRow, 4) = "Text in Cell"
'        Set xFso = CreateObject("Scripting.FileSystemObject")
'        Set xFld = xFso.GetFolder(xStrPath)
'        xStrFile = Dir(xStrPath & "\*.xls*")
'        Do While xStrFile <> ""
'            Set xWb = Workbooks.Open(Filename:=xStrPath & "\" & xStrFile, UpdateLinks:=0, ReadOnly:=True, AddToMRU:=False)
'            For Each xWk In xWb.Worksheets
'                Set xFound = xWk.UsedRange.Find(xStrSearch)
'                If Not xFound Is Nothing Then
'                    xStrAddress = xFound.Address
'                End If
'                Do
'                    If xFound Is Nothing Then
'                        Exit Do
'                    Else
'                        xCount = xCount + 1
'                        xRow = xRow + 1
'                        .Cells(xRow, 1) = xWb.Name
'                        .Cells(xRow, 2) = xWk.Name
'                        .Cells(xRow, 3) = xFound.Address
'                        .Cells(xRow, 4) = xFound.Value
'                    End If
'                    Set xFound = xWk.Cells.FindNext(After:=xFound)
'                Loop While xStrAddress <> xFound.Address
'            Next
'            xWb.Close (False)
'            xStrFile = Dir
'        Loop
'        .Columns("A:D").EntireColumn.AutoFit
'    End With
'    MsgBox xCount & "cells have been found", , "Kutools for Excel"
'ExitHandler:
'    Set xOut = Nothing
'    Set xWk = Nothing
'    Set xWb = Nothing
'    Set xFld = Nothing
'    Set xFso = Nothing
'    Application.ScreenUpdating = xUpdate
'    Exit Sub
'ErrHandler:
'    MsgBox Err.Description, vbExclamation
'    Resume ExitHandler
'End Sub


Sub SearchWKBooks()
Dim WS As Worksheet
Dim myfolder As String
Dim Str As String
Dim c As Range
Dim a As Single
Dim sht As Worksheet
Dim value As String
Dim firstAddress As String
Dim i, cell As Long

Set WS = ActiveSheet
With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
    myfolder = .SelectedItems(1) & "\"
End With
Str = Application.InputBox(prompt:="Search string:", Title:="Search all workbooks in a folder", Type:=2)

If Str = "" Then Exit Sub
WS.Range("A1") = "Search string:"
WS.Range("B1") = Str
WS.Range("A2") = "Path:"
WS.Range("B2") = myfolder
WS.Range("A3") = "Workbook"
WS.Range("B3") = "Worksheet"
WS.Range("C3") = "Cell Address"
WS.Range("D3") = "Link"
a = 0
value = Dir(myfolder)
Do Until value = ""
    If value = "." Or value = ".." Then
    Else
        If Right(value, 3) = "xls" Or Right(value, 4) = "xlsx" Or Right(value, 4) = "xlsm" Then
            On Error Resume Next
            Workbooks.Open Filename:=myfolder & value, Password:="zzzzzzzzzzzz"
            If Err.Number > 0 Then
                WS.Range("A4").Offset(a, 0).value = value
                WS.Range("B4").Offset(a, 0).value = "Password protected"
                a = a + 1
            Else
                On Error GoTo 0
                For Each sht In ActiveWorkbook.Worksheets
                        Set c = sht.Cells.Find(Str, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                        If Not c Is Nothing Then
                            firstAddress = c.Address
                            Do
                                WS.Range("A4").Offset(a, 0).value = value
                                WS.Range("B4").Offset(a, 0).value = sht.Name
                                WS.Range("C4").Offset(a, 0).value = c.Address
                                WS.Hyperlinks.Add Anchor:=WS.Range("D4").Offset(a, 0), Address:=myfolder & value, SubAddress:= _
                                sht.Name & "!" & c.Address, TextToDisplay:="Link"
                                a = a + 1
                                Set c = sht.Cells.FindNext(c)
                            Loop While Not c Is Nothing And c.Address <> firstAddress
                        End If
                Next sht
            End If
            Workbooks(value).Close False
            On Error GoTo 0
        End If
    End If
    value = Dir
Loop
Cells.EntireColumn.AutoFit


End Sub
