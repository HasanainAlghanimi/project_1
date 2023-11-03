Sub Gotoimage()
'
' Gotoimage Macro
'
' Keyboard Shortcut: Ctrl+Shift+E
'
    ActiveSheet.Shapes.AddShape(msoShapeOval, 142.2, 229.8, 98.4, 60).Select
    Application.WindowState = xlNormal
    Selection.Delete
    Application.CutCopyMode = False
    ActiveWindow.ScrollRow = 2
    ActiveSheet.Shapes.AddShape(msoShapeOval, 102.6, 249, 148.2, 72).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "Check this out "
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 15). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 15).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Range("C23").Select
    ActiveSheet.Shapes.Range(Array("Oval 4")).Select
    Selection.OnAction = "Gotoimage"
    Range("D24").Select
End Sub

