Sub MakeRowGreen()
'
' MakeRowGreen Macro
'
' Keyboard Shortcut: Ctrl+Shift+G
'
    Intersect(ActiveCell.EntireRow, ActiveCell.CurrentRegion).Select
    With Selection.Font
        .Color = -11489280
        .TintAndShade = 0
    End With
End Sub