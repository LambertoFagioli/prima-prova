Private Sub Workbook_Open()
    Call NascondiInterfaccia
    Worksheets(2).Activate
    chiudoOK = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not chiudoOK Then
        MsgBox "Per chiudere usa uno dei due tasti", vbCritical, TITOLO
        Cancel = True
    End If
End Sub

--------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    
    Dim sh As Worksheet
    Dim cella As Variant
    
    Application.EnableEvents = False
    Set sh = ThisWorkbook.Worksheets(2)
    Set cellRange = sh.Range("B4:G4")
    cella = Target.Address(False, False)
    
    If Not Intersect(Target, cellRange) Is Nothing Then
        If cella = "B4" Then
            If Not CheckDateRange(Target.Value) Then
                Target.Value = ""
                sh.Range(Target.Address).Select
                Application.EnableEvents = True
                Exit Sub
            End If
        Else
            If cella = "D4" Or cella = "E4" Then
                If Target.Value = "" Then
                    MsgBox "Una o più celle obbligatorie sono vuote", vbCritical
                    Target.Value = ""
                    sh.Range(Target.Address).Select
                    Application.EnableEvents = True
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Application.EnableEvents = True
End Sub
---------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    
    Dim sh As Worksheet
    Dim cella As Variant
    
    Application.EnableEvents = False
    Set sh = ThisWorkbook.Worksheets(2)
    Set cellRange = sh.Range("B4:G4")
    cella = Target.Address(False, False)
    
    If Not Intersect(Target, cellRange) Is Nothing Then
        If cella = "B4" Then
            If Not CheckDateRange(Target.Value) Then
                Target.Value = ""
                sh.Range(Target.Address).Select
                Application.EnableEvents = True
                Exit Sub
            End If
        Else
            If cella = "D4" Or cella = "E4" Or cella = "F4" Or cella = "G4" Then
                If Target.Value = "" Then
                    MsgBox "Una o più celle obbligatorie sono vuote", vbCritical
                    Target.Value = ""
                    sh.Range(Target.Address).Select
                    Application.EnableEvents = True
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Application.EnableEvents = True
End Sub
----------------------------------------
Sub RegistraOperazione()
    Dim sh As Worksheet
    Dim nCol As Long, i As Long
    Dim lo As ListObject
    Dim nextRow As ListRow
    Dim dbNome As String
    
    Set sh = ActiveSheet
    dbNome = sh.ListObjects(1).Name
    Set lo = sh.ListObjects(dbNome)
    Set nextRow = lo.ListRows.Add
    nCol = lo.ListColumns.Count
    
    ' Scrittura completa della riga
    For i = 1 To nCol - 1
        nextRow.Range.Cells(1, i).Value = sh.Cells(4, i + 1).Value
    Next i
    
    nextRow.Range.Cells(1, nCol).Value = Month(sh.Cells(4, 2).Value)
        
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add _
            Key:=lo.ListColumns("DATA").Range, _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With

    sh.Range("B4:G4").Value = ""
    sh.Range("B4").Select
    ThisWorkbook.Save

End Sub
--------------------------------------
Option Explicit
'
Public cellRange As Range
Public chiudoOK As Boolean
'

Public Sub NascondiInterfaccia()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
End Sub

Public Sub RipristinaInterfaccia()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
End Sub

Public Sub AnnullaChiude()
    ThisWorkbook.Saved = True
    Call RipristinaInterfaccia
    chiudoOK = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Quit
End Sub

Function CheckDateRange(dataIn) As Boolean
    Dim startDate As Date
    Dim endDate As Date
    Dim annoCorrente As String
    
    annoCorrente = Worksheets(5).Range("C3")
    startDate = "01/01/" & annoCorrente
    endDate = "31/12/" & annoCorrente

    If dataIn < startDate Or dataIn > endDate Then
        MsgBox "La data che hai digitato è sbagliata", vbCritical
        CheckDateRange = False
        Exit Function
    End If
        
    CheckDateRange = True
End Function

