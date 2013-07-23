Sub Right()
' Keyboard Shortcut: Ctrl+Shift+R
    With Selection
        .HorizontalAlignment = xlRight
    End With
End Sub
Sub Left()
' Keyboard Shortcut: Ctrl+Shift+L
    With Selection
        .HorizontalAlignment = xlLeft
    End With
End Sub
Sub BoldRed()
' Keyboard Shortcut: Ctrl+Shift+T
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub
Sub Underline()
' Keyboard Shortcut: Ctrl+Shift+W
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Sub ClearFormatting()
' Keyboard Shortcut: Ctrl+Shift+Q
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Italic = False
    Selection.Font.Bold = False
    Selection.Font.Underline = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Sub NumberFormat()
' Keyboard Shortcut: Ctrl+Shift+J
    Selection.NumberFormat = "#,##0"
End Sub
Sub FormatDollar()
' Keyboard Shortcut: Ctrl+Shift+M
    Selection.NumberFormat = "$#,##0"
End Sub
Sub FormatPainter()
' Keyboard Shortcut: Ctrl+Shift+P
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
Sub SoftHighlight()
' Keyboard Shortcut: Ctrl+Shift+N
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
