Sub DeleteBlankRows()

    'Sub to delete blank rows
    'Will cycle through each row
    'Has a counter to avoid infinite loop
    'Can check as many columns per row as desired by changing NumberColumns
        
    FinalRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
    Count = 0
    
    ' Change this to reflect number of columns to check before deleting
    NumberColumns = 5
    
    For x = 1 To FinalRow
        If IsEmpty(Cells(x, 1)) Then
            rowEmpty = True
            For y = 1 To NumberColumns
                If Not (IsEmpty(Cells(x, y))) Then
                    rowEmpty = False
                End If
            Next y
            If rowEmpty Then
                Rows(x).Delete
                x = x - 1
                Count = Count + 1
            End If
        End If
        If Count > FinalRow Then
            Exit For
        End If
    
    Next x

End Sub

Sub NextRow()
    'Selects next full row

    x = ActiveCell.Row
    Rows(x + 1).Select

End Sub


Function CountCharsinString(FindThis As String, StringToBeTested As String) As Integer
    'Sub to check how many times a particular character occurs in a string
    'Useful to check how many CSVs are in a cell
    
    Character = FindThis
    workingString = StringToBeTested
    
    CountCharsinString = 0
    For x = 1 To Len(StringToBeTested)
        If Left(workingString, 1) = Character Then CountCharsinString = CountCharsinString + 1
        workingString = Right(workingString, Len(workingString) - 1)
        
    Next x
    
End Function

Function AddSheetsAtEnd(NewSheetName)

    Sheets.Add(After:=Sheets(Sheets.Count)).Name = NewSheetName
    
End Function


Sub HighlightAlternateRows()

    ActiveSheet.UsedRange

    FinalRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
    FinalCol = ActiveSheet.Cells.SpecialCells(xlLastCell).Column
    
    'Will start on row after title row, and highlight each evenly numbered row
    TitleRow = 1
    
    For allRows = TitleRow + 1 To FinalRow
        If allRows Mod 2 = 0 Then
            Range(Cells(allRows, 1), Cells(allRows, FinalCol)).Select
        
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            
        End If
    Next allRows

End Sub


