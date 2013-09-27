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
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
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
Sub Autofit()
' Keyboard Shortcut: Ctrl+Shift+I
    ActiveCell.EntireColumn.Autofit
End Sub

Sub RowSelect()
    ActiveCell.EntireRow.Select
End Sub

Sub CycleThruColors()
    Static i 'counter
    
    If IsNull(i) Then i = 0
    
    i = i + 1
    
    If i >= 13 Then i = 1

    Select Case i
        Case 1: 'light grey
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
        Case 2: 'dark grey
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.499984740745262
                .PatternTintAndShade = 0
            End With
        Case 3: 'very dark grey
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349986266670736
                .PatternTintAndShade = 0
            End With
        Case 4: 'light brown
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = -9.99786370433668E-02
                .PatternTintAndShade = 0
            End With
        Case 5: 'medium brown
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
        Case 6: 'light blue
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        Case 7: 'dark blue
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Case 8: 'light red
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Case 9: 'light green
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Case 10: 'light purple
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Case 11: 'light turquoise
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
        Case 12: 'no fill
                With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
    End Select
        
End Sub

Sub CalibriFont10()
    Cells.Select
    Application.CutCopyMode = False
    With Selection.Font
        .Name = "Calibri"
        .Strikethrough = False
        .Size = 10
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Sub

