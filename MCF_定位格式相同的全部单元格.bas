'���=��λ����
'˵��=��λѡ����Ԫ��ʽ��ͬ��ȫ����Ԫ��

Sub ��λ��ʽ��ͬ��ȫ����Ԫ��()
    Dim FirstCell As Range, FoundCell As Range
    Dim AllCells As Range
    With Application.FindFormat
        .Clear
        .NumberFormatLocal = Selection.NumberFormatLocal
        .HorizontalAlignment = Selection.HorizontalAlignment
        .VerticalAlignment = Selection.VerticalAlignment
        .WrapText = Selection.WrapText
        .Orientation = Selection.Orientation
        .AddIndent = Selection.AddIndent
        .IndentLevel = Selection.IndentLevel
        .ShrinkToFit = Selection.ShrinkToFit
        .MergeCells = Selection.MergeCells
        .Font.Name = Selection.Font.Name
        .Font.FontStyle = Selection.Font.FontStyle
        .Font.Size = Selection.Font.Size
        .Font.Strikethrough = Selection.Font.Strikethrough
        .Font.Subscript = Selection.Font.Subscript
        .Font.Underline = Selection.Font.Underline
        .Font.ColorIndex = Selection.Font.ColorIndex
        .Interior.ColorIndex = Selection.Interior.ColorIndex
        .Interior.Pattern = Selection.Interior.Pattern
        .Locked = Selection.Locked
        .FormulaHidden = Selection.FormulaHidden
    End With
    
    Set FirstCell = ActiveSheet.UsedRange.Find(what:="", searchformat:=True)
    If FirstCell Is Nothing Then
        Exit Sub
    End If
    
    Set AllCells = FirstCell
    Set FoundCell = FirstCell
        
    Do
        Set FoundCell = ActiveSheet.UsedRange.Find(After:=FoundCell, what:="", searchformat:=True)
        If FoundCell Is Nothing Then Exit Do
        
        Set AllCells = Union(FoundCell, AllCells)
        If FoundCell.Address = FirstCell.Address Then Exit Do
    Loop
    AllCells.Select
End Sub

