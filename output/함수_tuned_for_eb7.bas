Sub �����()

dx(1, 0) = 100000
dx1(1, 0) = 100000
dx2(1, 0) = 100000
dx����(1, 0) = 100000

For i = 0 To insperiod
cx1(1, i) = 0
cx2(1, i) = 0
x = age + i
Select Case calc_type '���������
Case 1      '#1ȸ��, ������� ����, ��������� ���� ��X
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 2      '#��Ż��, �����̶� ���X
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 3      '#Ż��, �����̶� ���X + ��/�� ���ܺ�
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 4      '#Ż��, ������ ����
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 5      '#Ż��, �޼��ɱٰ���� ����
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 6      '#Ż��, ��+Ż������� + ��/�����ܺ�
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 7      '#���庸��ᳳ�Ը��������
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (1 - (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 8      '#Ż��, ��, Ż��������� �� ����(���׹׸�å����), + ���վ����ܺ�
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.25, 0) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 9      '#Ż��, ��, Ż��������� �� ����(���׹׸�åO) + �����ܺ�
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 10      '#Ż��, �ϻ���Ư��
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 11      '#��Ż��, �� ����(���׹׸�å����)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 12      '#��Ż��, �� ����(���׹׸�åO)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 13
Case 14
Case 15      '#������ġ���Կ�/��纴���Կ�(���׹׸�å����)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + 0.1 * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + 0.2 * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + 0.1 * qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + 0.1 * qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 16      '#������ġ���Կ�/��纴���Կ�(���׹׸�åO)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (IIf(i = 0 And renew <> 2, 0.75, 1) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + 0.1 * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + 0.2 * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + 0.1 * qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + 0.1 * qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 17      '#��������ȯ����,������������ȯ����(1ȸ��)
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 18      '#7����
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + 0.5 * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 19      '#���ػ��.�������
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - ������(������, ipno_n, i) + qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * ������(������, ipno_n, i) / 2) * v
z(i) = (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) / 2) / (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x))
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) * z(i)) * v * (1 - ����(youl, 1, sex, lev, x) * z(i)) * (1 - ����(youl, 1, sex, 1, x) * z(i)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) * z(i)) * (1 - ������(youl, 1, sex, 1, x) * z(i)) * (1 - �޼�(youl, 1, sex, 1, x) * z(i)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x) * z(i)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - ������(������, ipno_n, i) + qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * ������(������, ipno_n, i) / 2) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 20      '#Ư��������׾Ϲ�缱/�๰
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
dx1(1, i + 1) = dx1(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
dx2(1, i + 1) = dx2(1, i) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * (dx1(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + dx2(1, i) * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 21      '#ǥ���׾Ͼ๰�㰡ġ���(1ȸ��) (���׹׸�åO)
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x))) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 22      '#ǥ���׾Ͼ๰�㰡ġ���(1ȸ��) (���׹׸�å����)
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + ��(youl, 1, sex, 1, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 23      '#ǥ���׾Ͼ๰�㰡ġ���(����1ȸ��) (���׹׸�åO)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.25, 0) * (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x))) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 24      '#������ġ�������,����ġ�������պ��������(���׹׸�å����)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 25      '#������ġ�������,����ġ�������պ��������(���׹׸�åO)
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (IIf(i = 0 And renew <> 2, 0.75, 1) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 26      '#�׾Ͼ缺��/����������缱(���׹׸�åO)
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x))) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 27      '##�׾Ͼ缺��/����������缱(���׹׸�å����)
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + ��(youl, 1, sex, 1, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x))) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 28      '#4������ϼ�����
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 29      '#���庸��ᳳ������
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (1 - (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x))) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5 * 12 * ((1 - v ^ (insperiod - i - 1)) / (1 - v) + 0.5 * v ^ (insperiod - i - 1))

Case 30      '#Ư�����ؼ����������ܺ�
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * (1 - ���ؼ�������(youl, 1, sex, 1, x))
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 31      '#�������ı�����������̾�
dx(1, i + 1) = dx(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.25, 0) * (1 - (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)))) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (1 - (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x))) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 32      '#�׾Ϲ�缱/�๰ġ���(ġ��1ȸ��) ���׸�åO
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (IIf(i = 0 And renew <> 2, 0.75, 1) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + 0.2 * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + 0.2 * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 33      '#�׾Ϲ�缱/�๰ġ���(ġ��1ȸ��) ���׸�å����
dx(1, i + 1) = dx(1, i) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + 0.2 * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + 0.2 * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5

Case 34

Case 35

Case 36      '##�ž�ġ���(��)
dx(1, i + 1) = dx(1, i) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx(1, i) * (1 - ��(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * ��(youl, 1, sex, 1, x) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5 * (v ^ (0.5) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + v ^ (1.5) * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + v ^ (2.5) * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + v ^ (3.5) * qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + v ^ (4.5) * qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) / ��(youl, 1, sex, 1, x)

Case 37      '##�ž�ġ���(Ư�������)
dx(1, i + 1) = dx(1, i) * (1 - ��Ÿ(youl, 1, sex, 1, x) - ����(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - ��Ÿ(youl, 1, sex, 1, x) - ����(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx(1, i) * (1 - ��Ÿ(youl, 1, sex, 1, x) - ����(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5 * (v ^ (0.5) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + v ^ (1.5) * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + v ^ (2.5) * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + v ^ (3.5) * qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + v ^ (4.5) * qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) / (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x))

Case 38      '##�ž�ġ���(��,Ư�������)
dx(1, i + 1) = dx(1, i) * (1 - ��(youl, 1, sex, 1, x) - ��Ÿ(youl, 1, sex, 1, x) - ����(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i)) * v
If ���Ը��� = 1 Then
dx����(1, i + 1) = dx����(1, i) * (1 - ������(������, ipno_n, i)) * v * (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - ��(youl, 1, sex, 1, x) - ��Ÿ(youl, 1, sex, 1, x) - ����(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)
Else
dx����(1, i + 1) = dx(1, i) * (1 - ��(youl, 1, sex, 1, x) - ��Ÿ(youl, 1, sex, 1, x) - ����(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (��(youl, 1, sex, 1, x) + ��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x)) * (1 - ������(������, ipno_n, i) / 2) * v ^ 0.5 * (v ^ (0.5) * qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) + v ^ (1.5) * qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x) + v ^ (2.5) * qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x) + v ^ (3.5) * qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x) + v ^ (4.5) * qx����(n, 5, sex, lev, x) * ���(youl, n, sex, x)) / (��(youl, 1, sex, 1, x) + ��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x))

Case 39      '##�ž�ġ���(��,�߻��ڿ�)
dx(1, i + 1) = dx(1, i) * v
dx����(1, i + 1) = dx����(1, i) * v
cx(1, i) = dx(1, i) * qx����(n, Application.Min(i + 1, 5), sex, lev, age) * ���(youl, n, sex, age) / ��(youl, 1, sex, 1, age) * v ^ 0.5

Case 40      '##�ž�ġ���(Ư�������,�߻��ڿ�)
dx(1, i + 1) = dx(1, i) * v
dx����(1, i + 1) = dx����(1, i) * v
cx(1, i) = dx(1, i) * qx����(n, Application.Min(i + 1, 5), sex, lev, age) * ���(youl, n, sex, age) / (��Ÿ(youl, 1, sex, 1, age) + ����(youl, 1, sex, 1, age)) * v ^ 0.5

Case 41      '##�ž�ġ���(��,Ư�������,�߻��ڿ�)
dx(1, i + 1) = dx(1, i) * v
dx����(1, i + 1) = dx����(1, i) * v
cx(1, i) = dx(1, i) * qx����(n, Application.Min(i + 1, 5), sex, lev, age) * ���(youl, n, sex, age) / (��(youl, 1, sex, 1, age) + ��Ÿ(youl, 1, sex, 1, age) + ����(youl, 1, sex, 1, age)) * v ^ 0.5
End Select

If ������ = 1 Then
Select Case calc_type
Case 1, 3, 4, 5, 6, 8, 9, 10, 17, 21, 22, 30 'Ż��
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) / 2) * v ^ 0.5 / face
Case 2, 11, 12, 15, 16, 18, 19, 20, 23, 24, 25, 28, 32, 33 '��Ż��
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * v ^ 0.5 / face
Case 7
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - (1 - (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - (1 - (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)) / 2) * v ^ 0.5 / face
Case 26, 27
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - (qx����(n, 1, sex, lev, x) + qx����(n, 2, sex, lev, x) + qx����(n, 3, sex, lev, x)) * ���(youl, n, sex, x) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - (qx����(n, 1, sex, lev, x) + qx����(n, 2, sex, lev, x) + qx����(n, 3, sex, lev, x)) * ���(youl, n, sex, x) / 2) * v ^ 0.5 / face
Case 29
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - (1 - (1 - ����(youl, 1, sex, lev, x)) * (1 - ����(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * ��(youl, 1, sex, 1, x) - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x) - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 3, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 4, sex, lev, x) * ���(youl, n, sex, x)) * (1 - ������(youl, 1, sex, 1, x)) * (1 - �޼�(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - ���ؼ�������(youl, 1, sex, 1, x)), 1)) / 2) * v ^ 0.5 / face
Case 31
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - (1 - (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x))) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - (1 - (1 - qx����(n, 1, sex, lev, x) * ���(youl, n, sex, x)) * (1 - qx����(n, 2, sex, lev, x) * ���(youl, n, sex, x))) / 2) * v ^ 0.5 / face
Case 36
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - ��(youl, 1, sex, 1, x) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - ��(youl, 1, sex, 1, x) / 2) * v ^ 0.5 / face
Case 37
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - (��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
Case 38
cx1(1, i) = w_rate(������, ipno_n, i) * (ȯ�ޱ�_ǥ��(i, 1) + ȯ�ޱ�_ǥ��(i + 1, 1)) / 2 * dx����(1, i) * ������(������, ipno_n, i) * (1 - (��(youl, 1, sex, 1, x) + ��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(������, ipno_n, i) * (�غ��_ǥ��(i, 1) + �غ��_ǥ��(i + 1, 1)) / 2 * (dx(1, i) - dx����(1, i)) * ������(������, ipno_n, i) * (1 - (��(youl, 1, sex, 1, x) + ��Ÿ(youl, 1, sex, 1, x) + ����(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
End Select
End If

Next i
���������:

End Sub
Sub �������()

j_e = insperiod

For j = 0 To j_e

mxsum = 0
nxsum = 0
nxsum_k = 0
For jj = j To j_e
mxsum = mxsum + cx(1, jj) + cx1(1, jj) + cx2(1, jj)
nxsum = nxsum + dx����(1, jj)
nxsum_k = nxsum_k + dx(1, jj)
Next jj
mx(j) = mxsum
nx����(j) = nxsum
nx(j) = nxsum_k
Next j

nx���� = (nx����(0) - nx����(premperiod)) - 11 / 24 * (dx����(1, 0) - dx����(1, premperiod))
End Sub
Sub ���������()

bunja_p = mx(0) - mx(insperiod)

If premperiod = 0 Then
������(irate) = bunja_p / dx(1, 0)
������(irate) = bunja_p / dx(1, 0)
���ѵ�(irate) = ������(irate)
Else
������(irate) = bunja_p / (nx����(0) - nx����(premperiod))
������(irate) = bunja_p / (nx���� * 12)
���ѵ�(irate) = bunja_p / (nx����(0) - nx����(handoprem))
End If

���ѵ�1��(irate) = ���ѵ�(irate) * face
������1��(irate) = Round(������(irate) * face)
������1��(irate) = Round(������(irate) * face)
End Sub
Sub �����������()

If premperiod = 0 Then
�������� = ������(1) '/ (1 - alpha2 - beta - ce) '�Ͻó�
�������� = ������(1)
Else
�������� = (������(1) + alpha1(n, youl, renew, IIf(gubun = "01", (mangi / 10 - 7), 0), ipno_n) * 100000 / (nx����(0) - nx����(premperiod))) / (1 - beta(n, youl, renew) - ce - beta5 - (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) / (nx����(0) - nx����(premperiod)) - alpha2(n, youl, renew, premperiod) * 100000 / (nx����(0) - nx����(premperiod)))
�������� = (������(1) + alpha1(n, youl, renew, IIf(gubun = "01", (mangi / 10 - 7), 0), ipno_n) * 100000 / (12 * nx����)) / (1 - beta(n, youl, renew) - ce - beta5 - (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) / nx���� - alpha2(n, youl, renew, premperiod) * 100000 / nx����)
End If
��������1�� = Round(�������� * face)
��������1�� = Round(�������� * face)

If renew = 0 And gubun = "01" Then
��p = Round((������(irate) + �������� * (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) / nx����) * face)
Else
��p = ������1��(1)
End If

End Sub
Sub �ѵ�üũ()

s��뿩�� = Sheets("NSP���̾ƿ�").Cells(3 + n, 42)

�ؾ������� = Application.Min(20, insperiod)

Select Case jong
Case 1
s = Sheets("�����1��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
Case 2
s = Sheets("�����2��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
Case 3
s = Sheets("�����3��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
Case 4
s = Sheets("�����4��").Cells(6 + nn, 39 + 4 * ������ + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
End Select


If premperiod = 0 Then
�Ű����ѵ� = 0
���Ű��� = 0

Else
If s��뿩�� <> 0 Then
�Ű����ѵ� = (���ѵ�(1) * face * 0.05 * �ؾ������� + s * face * 10 / 1000) '���ѵ�1��(4)=ǥ�ؿ����������
Else
�Ű����ѵ� = (���ѵ�(1) * face * (0.05 * �ؾ�������)) + ���ѵ�(1) * face * 0.45
End If

���Ű��� = ��������1�� * alpha2(n, youl, renew, premperiod) * 12 + face * alpha1(n, youl, renew, IIf(gubun = "01", (mangi / 10 - 7), 0), ipno_n)

End If
End Sub
Sub �غ�ݰ��()

For i = 0 To insperiod
If mangi = 5 Then
�غ��(i, irate) = (mx(i) - mx(insperiod)) / dx(1, i)
ElseIf i < premperiod Then
�غ��(i, irate) = (mx(i) - mx(insperiod) - ������(irate) * (nx����(i) - nx����(premperiod)) + �������� * (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) * (nx����(0) - nx����(i)) / (nx����(0) - nx����(premperiod))) / dx(1, i)
Else
�غ��(i, irate) = (mx(i) - mx(insperiod) + �������� * (beta1 + ce1) * (nx(i) - nx(insperiod))) / dx(1, i)
End If
�غ��1��(i, irate) = Round(�غ��(i, irate) * face)
Next i
End Sub
Sub ǥ���غ�ݰ��()
For i = 0 To insperiod
�غ��_ǥ��(i, irate) = Round(�غ��(i, irate) * face)
�غ��_ǥ��(i, irate) = Application.Max(�غ��_ǥ��(i, irate), 0)
���������Ⱓ = Application.Min(7, premperiod)
ȯ�ޱ�_ǥ��(i, irate) = Application.Max(�غ��_ǥ��(i, irate) - IIf(i > ���������Ⱓ, 0, Application.RoundDown((���������Ⱓ - i) / ���������Ⱓ * Application.Min(Application.RoundDown(�Ű����ѵ�, 0), ���Ű���), 0)), 0)
Next i

End Sub