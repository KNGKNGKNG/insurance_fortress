Sub P���()

txt(1) = covcode
txt(2) = n
txt(3) = sex
txt(4) = insperiod
txt(5) = premperiod
txt(6) = renew
txt(7) = lev
txt(8) = age
txt(9) = youl
txt(10) = drv
txt(11) = "����P=" & ��������1��
txt(12) = "��ǰP=" & ��ǰp

Dim MAL As Long
MAL = 12

For a = 1 To MAL
If a < MAL Then
Print #1, Trim(txt(a)); " ; ";      ' ����Ʈ�ض� #1�� Ʈ���̶�� �迭
Else
Print #1, Trim(txt(MAL))
End If
Next

End Sub
Sub V���()

txt(1) = covcode
txt(2) = n
txt(3) = sex
txt(4) = insperiod
txt(5) = premperiod
txt(6) = renew
txt(7) = lev
txt(8) = age
txt(9) = youl
txt(10) = drv
txt(11) = "����V=" & Sum_����V
txt(12) = "��ǰV=" & Sum_��ǰV
txt(13) = "�����ѵ�=" & Int(�Ű����ѵ�)
txt(14) = "��ǰ�ѵ�=" & ��ǰ�ѵ�
txt(15) = "��������=" & ��p
txt(16) = "��ǰ����=" & ��ǰnp

Dim MAL As Long
'      MAL = 12 + nn
'MAL = 11
MAL = 16
For a = 1 To MAL
If a < MAL Then
Print #2, Trim(txt(a)); ";";
Else
Print #2, Trim(txt(MAL))
End If
Next

End Sub
Sub �ѵ����()

txt(1) = covcode
txt(2) = n
txt(3) = sex
txt(4) = insperiod
txt(5) = premperiod
txt(6) = renew
txt(7) = lev
txt(8) = age
txt(9) = youl
txt(10) = drv
txt(11) = "�ѵ�=" & Int(�Ű����ѵ�)
txt(12) = "�Ű���=" & ���Ű���

Dim MAL As Long
MAL = 12
For a = 1 To MAL
If a < MAL Then
Print #3, Trim(txt(a)); ";";
Else
Print #3, Trim(txt(MAL))
End If
Next

End Sub
Sub s���()
Select Case jong
Case 1
Sheets("�����1��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
Case 2
Sheets("�����2��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
Case 3
Sheets("�����3��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
Case 4
Sheets("�����4��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
End Select
End Sub

Sub ���̽����()
Sheets("���").Cells(k, 1) = jong
Sheets("���").Cells(k, 2) = sex
Sheets("���").Cells(k, 3) = covcode
Sheets("���").Cells(k, 4) = "" 'insperiod
Sheets("���").Cells(k, 5) = "" 'premperiod
Sheets("���").Cells(k, 6) = "" 'renew
Sheets("���").Cells(k, 7) = age
Sheets("���").Cells(k, 8) = lev
Sheets("���").Cells(k, 9) = "" 'youl
Sheets("���").Cells(k, 10) = qx����(n, 1, sex, lev, age)
Sheets("���").Cells(k, 11) = qx����(n, 2, sex, lev, age)
Sheets("���").Cells(k, 12) = ""
Sheets("���").Cells(k, 13) = ""
k = k + 1
End Sub