'���=��ֵת��
'˵��=����ת����Ҵ�д���

Sub ����ת����Ҵ�д���()
    On Error Resume Next
    dim M as Range
    set M = ActiveCell
    y = Int(Round(100 * Abs(M)) / 100)
    j = Round(100 * Abs(M) + 0.00001) - y * 100
    f = (j / 10 - Int(j / 10)) * 10
    A = IIf(y < 1, "", Application.Text(y, "[DBNum2]") & "Ԫ")
    b = IIf(j > 9.5, Application.Text(Int(j / 10), "[DBNum2]") & "��", IIf(y < 1, "", IIf(f > 1, "��", "")))
    c = IIf(f < 1, "��", Application.Text(Round(f, 0), "[DBNum2]") & "��")
    M = IIf(Abs(M) < 0.005, "", IIf(M < 0, "��" & A & b & c, A & b & c))
End Sub





