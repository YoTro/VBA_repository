'���=����¼��
'˵��=ѡ��¼�뵱ǰ��Ԫ��ַ




Sub ѡ��¼�뵱ǰ��Ԫ��ַ()
    Selection = "=ADDRESS(ROW(),COLUMN(),4,1)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub




