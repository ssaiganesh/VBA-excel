Sub SearchTerm()
Dim strTest As String
strTest = InputBox("Search for the word and create seperate sheet", "search and filter rows")
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
    wsDest.Cells.Select
    wsDest.Cells.EntireRow.AutoFit
    
    wsDest.Name = strTest & " " & wsSource.Name
End Sub
