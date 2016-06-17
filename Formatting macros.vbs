' I would like to participate! Perhaps we can create a formatting tool which exceeds EY ;)
Option Explicit
Public Sub Right()
' Keyboard Shortcut: Ctrl+Shift+R
    With Selection
        .HorizontalAlignment = xlRight
    End With
End Sub
Public Sub Left()
' Keyboard Shortcut: Ctrl+Shift+L
    With Selection
        .HorizontalAlignment = xlLeft
    End With
End Sub
Public Sub BoldRed()
' Keyboard Shortcut: Ctrl+Shift+T
    Selection.Font.Bold = True
    With Selection.Font
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
    End With
End Sub
Public Sub Underline()
' Keyboard Shortcut: Ctrl+Shift+W
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub ClearFormatting()
' Keyboard Shortcut: Ctrl+Shift+Q
    Selection.ClearFormats
End Sub
Public Sub NumberFormat()
' Keyboard Shortcut: Ctrl+Shift+J
    Selection.NumberFormat = "#,##0"
End Sub
Public Sub FormatDollar()
' Keyboard Shortcut: Ctrl+Shift+M
    Selection.NumberFormat = "$#,##0"
End Sub
Public Sub FormatPainter()
' Keyboard Shortcut: Ctrl+Shift+P
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub
Public Sub SoftHighlight()
' Keyboard Shortcut: Ctrl+Shift+N
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 153)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Public Sub Autofit()
' Keyboard Shortcut: Ctrl+Shift+I
    ActiveCell.CurrentRegion.EntireColumn.Autofit
    ActiveCell.CurrentRegion.EntireRow.Autofit
End Sub

Public Sub RowSelect()
    ActiveCell.EntireRow.Select
End Sub

Public Sub CycleThruColors()
' Keyboard Shortcut: Ctrl+Shift+O
    Static i 'counter
    
    'initialize the variable, since it is not yet an object, but only set to zero if it is not yet an object
    If IsNull(i) Then i = 0
    
    'increment the counter by one for each time you run the macro
    i = i + 1
    
    'reset the counter
    If i > 6 Then i = 1

    'the magic happens here :)
    Select Case i
        Case 1: 'light grey
            With Selection.Interior
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
            End With
            
        Case 2: 'light blue
            With Selection.Interior
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.799981688894314
            End With
            
        Case 3: 'light red
            With Selection.Interior
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
            End With
            
        Case 4: 'light green
            With Selection.Interior
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.599993896298105
            End With
            
        Case 5: 'light purple
            With Selection.Interior
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.599993896298105
            End With
            
        Case 6: 'no fill
                With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

    End Select
        
End Sub


Public Sub CalibriFont10()
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

Public Sub FormatColumnWidth()
' Keyboard Shortcut: Ctrl+Shift+G
    Cells.Select
    Selection.ColumnWidth = 10.5
End Sub


