Sub SearchTerm()
Const strTest = "SearchTerm"
Dim wsSource As Worksheet
Dim wsDest As Worksheet
Dim NoRows As Long
Dim DestNoRows As Long
Dim I As Long
Dim rngCells As Range
Dim rngFind As Range
    
    Set wsSource = ActiveSheet
    
    NoRows = wsSource.Range("A65536").End(xlUp).Row
    DestNoRows = 3
    Set wsDest = ActiveWorkbook.Worksheets.Add
    
    wsSource.Range("A3:I3").Copy
    wsDest.Activate
    Range("A1").Select
    wsDest.Paste
    wsSource.Activate
    wsSource.Range("A4:I4").Copy
    wsDest.Activate
    Range("A2").Select
    wsDest.Paste
    
    For I = 1 To NoRows
    
        Set rngCells = wsSource.Range("A" & I & ":I" & I)
        
        If Not (rngCells.Find(strTest) Is Nothing) Then
            rngCells.EntireRow.Copy wsDest.Range("A" & DestNoRows)
            
            DestNoRows = DestNoRows + 1
        End If
    Next I
    wsDest.Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    wsDest.Columns("B:B").EntireColumn.AutoFit
    
End Sub

