Sub ����������ν�()
For youl = 0 To 1
For sex = 1 To 2 '����
For x = 0 To 110 '���Գ���+����Ⱓ
����(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 170) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 40), 1)
��(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 52) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
������(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 108) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 14), 1)
�޼�(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 122) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 27), 1)
���ؼ�������(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 410) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 79), 1)
�����(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 884) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
��Ÿ(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 54) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
����(youl, 1, sex, 1, x) = Sheets("���������").Cells(x + 6, sex + 56) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
For lev = 1 To 3 '�޼�
����(youl, 1, sex, lev, x) = Sheets("���������").Cells(x + 6, sex + 8 + 2 * (lev - 1)) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 40), 1)
Next lev
Next x
Next sex
Next youl

End Sub
Sub ���������ν�()
For sex = 1 To 2 '����
For lev = 1 To 1 '�޼�
For x = 0 To 110 '���Գ���+����Ⱓ
qx�����(sex, lev, x) = Sheets("���������").Cells(x + 6, sex + 6 + 2 * (lev - 1))
Next x
Next lev
Next sex
End Sub
Sub ��������ν�()
For youl = youl_s To youl_e
For renew = renew_s To renew_e

For premperiod = 1 To 30
If nn = 0 Then
alpha2(n, youl, 0, 0) = 0
Else
Select Case jong
Case 1
alpha2(n, youl, renew, premperiod) = Sheets("�����1��").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
Case 2
alpha2(n, youl, renew, premperiod) = Sheets("�����2��").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
Case 3
alpha2(n, youl, renew, premperiod) = Sheets("�����3��").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
Case 4
alpha2(n, youl, renew, premperiod) = Sheets("�����4��").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
End Select
End If
Next premperiod

If nn = 0 Then
beta(n, youl, renew) = 0
Else
Select Case jong
Case 1
beta(n, youl, renew) = Sheets("�����1��").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
Case 2
beta(n, youl, renew) = Sheets("�����2��").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
Case 3
beta(n, youl, renew) = Sheets("�����3��").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
Case 4
beta(n, youl, renew) = Sheets("�����4��").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
End Select
End If

For mangi_k = mangi_k_s To mangi_k_e
For ipno_n = ipno_n_s To ipno_n_e

If nn = 0 Then
alpha1(n, youl, 0, 0, 0) = 0
Else
Select Case jong
Case 1
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("�����1��").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))
Case 2
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("�����2��").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))
Case 3
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("�����3��").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))
Case 4
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("�����4��").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))

End Select
End If
beta5 = Sheets("����").Range("J25") '����������� ���ݺ�
ce = Sheets("����").Range("J26") '���������
beta1 = Sheets("����").Range("J27") '(������)����������� ���ݺ�
ce1 = Sheets("����").Range("J28") '(������)���������
Next ipno_n
Next mangi_k
Next renew
Next youl

End Sub
Sub �������ν�()
For ipno_n = 0 To 4
Select Case ipno_n
Case 1
premperiod = 10
Case 2
premperiod = 15
Case 3
premperiod = 20
Case 4
premperiod = 30
End Select
For ������ = 0 To 1
For i = 0 To 110
If ������ = 1 Then
If jong = 2 Or jong = 4 Or ipno_n = 0 Then
������(������, ipno_n, i) = 0
w_rate(������, ipno_n, i) = 0
ElseIf i = premperiod - 2 Then '�ϳ�2����
������(������, ipno_n, i) = Sheets("������").Cells(5, 19)
w_rate(������, ipno_n, i) = 0 '����ȯ�ޱ� ���޷�
ElseIf i = premperiod - 1 Then '�ϳ�1����
������(������, ipno_n, i) = Sheets("������").Cells(5, 20)
w_rate(������, ipno_n, i) = 0 '����ȯ�ޱ� ���޷�
ElseIf i = premperiod Then '������1����
������(������, ipno_n, i) = Sheets("������").Cells(5, 21)
w_rate(������, ipno_n, i) = 0.5 '����ȯ�ޱ� ���޷�
ElseIf i >= premperiod + 1 Then '������1���ʰ�
������(������, ipno_n, i) = Sheets("������").Cells(5, 22)
w_rate(������, ipno_n, i) = 0.5 '����ȯ�ޱ� ���޷�
Else '�׿ܳ�����
������(������, ipno_n, i) = Sheets("������").Cells(3 + i, 2)
w_rate(������, ipno_n, i) = 0 '����ȯ�ޱ� ���޷�
End If
ElseIf ������ = 0 Then
������(������, ipno_n, i) = 0
w_rate(������, ipno_n, i) = 0
End If
Next i
Next ������
Next ipno_n
End Sub
Sub ��������ν�()
For youl = youl_s To youl_e
For sex = sex_s To sex_e
For x = 0 To 110 '���Գ���+����Ⱓ
���(youl, n, sex, x) = IIf((jong = 1 Or jong = 2) And si <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + si), 1) '(kk����, i����Ⱓ, sex����, lev�޼�, age(����), Ű��(n) )
Next x
Next sex
Next youl
End Sub
Sub ������ν�()
For sex = sex_s To sex_e
For lev = 1 To lev_e
For kk = 1 To n_rate_k
For x = 0 To 110

Select Case n_rate
Case 11
qx����(n, kk, sex, lev, x) = Sheets("���������").Cells(x + 6, (sex - 1) * 5 + n_rate_c(1) + 2 * (lev - 1) + kk) '(kk������ ����Ⱓ, i����Ⱓ, sex����, lev�޼�, age(����), Ű��(n) )
Case Else
qx����(n, kk, sex, lev, x) = Sheets("���������").Cells(x + 6, sex + n_rate_c(kk) + 2 * (lev - 1))
End Select

Next x
Next kk
Next lev
Next sex

End Sub