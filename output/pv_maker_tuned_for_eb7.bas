Option Explicit

Public start As Integer, last As Integer, �������� As Integer, s���⿩�� As Integer, s��뿩�� As Integer, youl_s As Integer, youl_e As Integer, youl As Integer, a As Integer, b As Integer, c As Integer, ������ As Integer
Public alpha1(1163, 6, 2, 3, 4) As Double, alpha2(1163, 6, 2, 30) As Double, beta(1163, 6, 2) As Double, beta5 As Double, ce As Double, beta1 As Double, ce1 As Double
Public z(110) As Double, ����s��(1163, 6) As Double, s��(1163, 2, 6, 3, 3, 4) As Double, ���(6, 1163, 2, 110) As Double, qx����(1163, 5, 2, 3, 110) As Double, qx�����(2, 3, 110) As Double, s As Double, ����(6, 5, 2, 3, 110) As Double, ����(6, 5, 2, 3, 110) As Double, ��(6, 5, 2, 3, 110) As Double, ������(6, 5, 2, 3, 110) As Double, �޼�(6, 5, 2, 3, 110) As Double, �Ⱙ(6, 5, 2, 3, 110) As Double, ���ؼ�������(6, 5, 2, 3, 110) As Double, �����(6, 5, 2, 3, 110) As Double, ��Ÿ(6, 5, 2, 3, 110) As Double, ����(6, 5, 2, 3, 110) As Double
Public dx(5, 110) As Double, dx����(5, 110) As Double, dx�����(110) As Double, AA As Double, dx1(5, 110) As Double, dx2(5, 110) As Double, dx_ǥ��(5, 110) As Double, dx����_ǥ��(5, 110) As Double, dx1_ǥ��(5, 110) As Double, dx2_ǥ��(5, 110) As Double
Public cx(5, 110) As Double, cx1(5, 110) As Double, cx2(5, 110) As Double, cx�����(110) As Double, cx_ǥ��(5, 110) As Double, cx1_ǥ��(5, 110) As Double, cx2_ǥ��(5, 110) As Double
Public rate(4) As Double, v As Double
Public txt_num As Integer, i As Integer, j As Integer, j_e As Integer, jj As Integer, n As Integer, k As Integer, kk As Integer, kkk As Integer, no As Integer, t As Integer, irate As Integer, x As Integer, mangi_k As Integer, mangi_k_s As Integer, mangi_k_e As Integer, ���Ը��� As Integer
Public face As Long
Public sex As Integer, sex_check As Integer, sex_s As Integer, sex_e As Integer, cl_s As Integer, cl_e As Integer, cl As Integer, drv_s As Integer, drv_e As Integer, drv As Integer, lev_e As Integer, lev As Integer, age_s As Integer, age_e As Integer, age As Integer, age_s_check As Integer, renew_s As Integer, renew_e As Integer, renew As Integer, renewperi As Integer, renewperi_s As Integer, renewperi_e As Integer, �����ֱ� As Integer, si As Integer
Public ipno_p_s As Integer, ipno_p_e As Integer, ipno_p As Integer, ipno_m_s As Integer, ipno_m_e As Integer, ipno_m As Integer, ipno_n_e As Integer, ipno_n As Integer, mangi_type As Integer, mangi As Integer, ipno_n_s As Integer
Public insperiod As Integer, premperiod As Integer, handoprem As Integer
Public nx����(110) As Double, nx����_ǥ��(110) As Double, nx(110) As Double, nx_ǥ��(110) As Double, nx�ѵ�(110) As Double, nx���� As Double, nx����� As Double, nx����_ǥ�� As Double, nx��������� As Double, bunja_p As Double, mx(110) As Double, bunja_p_ǥ�� As Double, mx_ǥ��(110) As Double, mx����� As Double, mxsum As Double, nxsum As Double, nxsum_k As Double, mxsum_ǥ�� As Double, nxsum_ǥ�� As Double, nxsum_k_ǥ�� As Double
Public ����������� As Double, ����������� As Double, ������(1) As Double, ������(1) As Double, �������� As Double, �������� As Double, ������_ǥ��(1) As Double, ������_ǥ��(1) As Double, ��������_ǥ�� As Double, ��������_ǥ�� As Double
Public ���ѵ�(1) As Double, ���ѵ�_ǥ��(1) As Double, ���ѵ�1��(1) As Long, ���ѵ�1��_ǥ��(1) As Long, �ؾ������� As Long, ���������Ⱓ As Long
Public ������1��(1) As Long, ������1��(1) As Long, ��������1�� As Long, ��������1�� As Long, ������1��_ǥ��(1) As Long, ������1��_ǥ��(1) As Long, ��������1��_ǥ�� As Long, ��������1��_ǥ�� As Long, �����������1�� As Long, �����������1�� As Long, ��ǰp As Long, ��ǰnp As Long, Sum_��ǰV As Long, Sum_����V As Long, ��ǰ�ѵ� As Long, ��p As Long
Public �غ��(110, 1) As Double, �غ��_ǥ��(110, 1) As Double, ȯ�ޱ�_ǥ��(110, 1) As Double
Public �غ��1��(110, 1) As Double, �غ��1��_ǥ��(110, 1) As Double
Public plancode As String, covcode As String, zz As String
Public �㺸�� As String
Public nn As Integer
Public j1 As Integer, jj1 As Integer

Public ������(1, 4, 110) As Double, w_rate(1, 4, 110) As Double
Public ��ǰV(110) As Long
Public txt(430) As String
Public hh As Integer, hhh As Integer, scheck As Integer
Public �Ű����ѵ� As Double, ���Ű��� As Double, �Ű����ѵ�_ǥ�� As Double, ���Ű���_ǥ�� As Double, ��ü������� As Double
Public i2 As Integer, mm As Integer
Public ����ȯ�ޱ�(21) As Long, USUM As Long
Public �󰢱Ⱓ As Integer, ���⿩�� As Integer
Public gubun As String
Public ��θ� As String, ���ϸ� As String
Public �������μ�(1163) As Long, ��ǰ���μ�(1163) As Long
Public �������� As String
Public renewperiod() As Variant
Public jong As Integer, jong_k As Integer, jong_s As Integer, jong_e As Integer, jong_ss As Integer, jong_ee As Integer, n_rate As Integer, n_rate_k As Integer, n_rate_r As Integer, n_rate_c(5) As Integer
Public calc_type As Integer
Public ws As Worksheet
Public Rng1 As Range
Sub ����()

Application.StatusBar = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Sheets("����").Range("J11") = Now '���۽ð�
renewperiod() = Array(0, 3, 10, 20) '������ �����ֱ�
irate = 1 ' ����/ǥ��V(ǥ��i+����)/����V(ǥ��i*125%)/ǥ������������(�ѵ�üũ��,ǥ��i) ''15.01����
rate(1) = Sheets("����").Range("J7") '��������
v = 1 / (1 + rate(irate))
k = 3 '��¿�
�������� = Sheets("����").Range("J9") '1: �����Ǽ�üũ(��ü��) 2: s����(��ü��) 3: p���� 4: v����

If �������� = 1 Then
For a = Sheets("����").Range("J4") To Sheets("����").Range("J5")
�������μ�(a) = 0
Next a
ElseIf �������� = 3 Then
For a = Sheets("����").Range("J4") To Sheets("����").Range("J5")
��ǰ���μ�(a) = 0
Next a
End If

jong = CInt(Sheets("����").Range("J20"))

If �������� > 1 Then
Call ���������ν�
Call ����������ν�
Call �������ν�
End If

For n = Sheets("����").Range("J4") To Sheets("����").Range("J5")  '���̺���

jong_s = CInt(Sheets("NSP���̾ƿ�").Cells(3 + n, 39))
jong_e = CInt(Sheets("NSP���̾ƿ�").Cells(3 + n, 40))

If (jong < jong_s Or jong > jong_e) Then GoTo ����n

If �������� = 2 Then
���⿩�� = Sheets("NSP���̾ƿ�").Cells(3 + n, 42)
Else
���⿩�� = 1
End If

If ���⿩�� = 1 Then


�������� = Sheets("NSP���̾ƿ�").Cells(3 + n, 20) '* ���о��� 0 ǥ��ü 1 ǥ����ü �ǰ���ޱ׷��ڵ�ABCD
covcode = Sheets("NSP���̾ƿ�").Cells(3 + n, 6)
nn = Sheets("NSP���̾ƿ�").Cells(3 + n, 42 + jong) '����� ��ġ

calc_type = Sheets("NSP���̾ƿ�").Cells(3 + n, 60) '���������
AA = Sheets("NSP���̾ƿ�").Cells(3 + n, 61) '�ʳ⵵ Cx�� ���� ����(����/��å)

gubun = Sheets("NSP���̾ƿ�").Cells(3 + n, 18)    '���ⱸ���ڵ�(01:������,02:������,*:�¾� ��)

sex_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 28)        '0 ���о��� 1 ���� 2 ����
sex_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 29)       '����

drv_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 31)
drv_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 32)

renew_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 33)
renew_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 34)

renewperi_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 62)
renewperi_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 63)

mangi_k_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 65)
mangi_k_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 66)

n_rate = Sheets("NSP���̾ƿ�").Cells(3 + n, 47)  '���������
n_rate_k = Sheets("NSP���̾ƿ�").Cells(3 + n, 48)  '���������
n_rate_r = Sheets("NSP���̾ƿ�").Cells(3 + n, 49)  '�������
si = Sheets("NSP���̾ƿ�").Cells(3 + n, 55)  '�������

ipno_n_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 36)
ipno_n_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 35)

lev_e = Sheets("NSP���̾ƿ�").Cells(3 + n, 38)

For kk = 1 To n_rate_k
n_rate_c(kk) = Sheets("NSP���̾ƿ�").Cells(3 + n, 49 + kk)
Next kk

face = Sheets("NSP���̾ƿ�").Cells(3 + n, 27) * 10000 '���رݾ�
If �������� = "0013" Then
youl_s = 1
youl_e = 1
Else
youl_s = 0
youl_e = 0
End If
'-----------------------------------------------------------------------���̺�ȭ �̰ͱ��� s�����̵� p�����̵� v�����̵� ���ؾ���(�Ǽ�üũ ����)
If �������� > 1 Then
Call ��������ν�
Call ��������ν�
Call ������ν�
End If

'-------------------------------------------------------------------------
If �������� < 3 Then
'-------------------------------------------------------------------------
For sex = sex_s To sex_e '����
If �������� = 2 And (sex = 2 And sex_s <> 2) Then GoTo ����sex 'S����/����������: ������ �������� ��ƾ���

For drv = drv_s To drv_e                        '��������
For youl = youl_s To youl_e '��������
For renew = renew_s To renew_e
If �������� = 2 And renew = 2 Then GoTo ����renew

For renewperi = renewperi_s To renewperi_e
If renew = 0 Then
�����ֱ� = 0
Else
�����ֱ� = renewperiod(renewperi)
End If
mangi_type = Sheets("NSP���̾ƿ�").Cells(3 + n, 64)

'�ڡڡڡڡڡڡڡڡڡڡڡڡڡں�������
If Sheets("NSP���̾ƿ�").Cells(3 + n, 19) = 5 Then '�ž�ġ���
ipno_p_s = 1
ipno_p_e = 1
ElseIf renew = 1 Then '������ ���ʰ��
ipno_p_s = �����ֱ�
ipno_p_e = �����ֱ�
ElseIf renew = 2 And renewperi = renewperi_e Then '������ ���Ű��
ipno_p_s = 1
ipno_p_e = �����ֱ�
ElseIf renew = 0 Then '������&������񰻽�
ipno_p_s = 1
ipno_p_e = 1
Else
ipno_p_s = �����ֱ�
ipno_p_e = �����ֱ�
End If

For ipno_p = ipno_p_s To ipno_p_e '�ڡڡں���Ⱓ
If Sheets("NSP���̾ƿ�").Cells(3 + n, 19) = 5 Then
insperiod = 5
ElseIf renew <> 0 Then
insperiod = ipno_p
End If

If mangi_type = 0 Then '�¾ƺ���� ����/����=0, P�� ����� ������
ipno_m_s = 1
ipno_m_e = 1
Else
If mangi_type = 1 Then ipno_m_e = 2
If mangi_type = 2 Then ipno_m_e = 1
If mangi_type = 3 Then ipno_m_e = 3
If gubun = "01" Then
ipno_m_s = 1
ElseIf insperiod = �����ֱ� Then
ipno_m_s = ipno_m_e
Else
ipno_m_s = 1
End If
End If
For ipno_m = ipno_m_s To ipno_m_e '�ڡڡڸ���

If Sheets("NSP���̾ƿ�").Cells(3 + n, 19) = 5 Then
mangi = 5
ElseIf mangi_type = 0 Then
mangi = 0
ElseIf mangi_type = 1 Then
Select Case ipno_m
Case 1
mangi = 90
Case 2
mangi = 100
End Select
ElseIf mangi_type = 2 Then
Select Case ipno_m
Case 1
mangi = 80
End Select
ElseIf mangi_type = 3 Then
Select Case ipno_m
Case 1
mangi = 80
Case 2
mangi = 90
Case 3
mangi = 100
End Select
End If

For ipno_n = ipno_n_s To ipno_n_e '��������
If �������� = 2 And gubun = "01" And ipno_n <> ipno_n_e Then GoTo ��������
If (jong = 1 Or jong = 3) And ipno_n = 1 Then GoTo ��������
If Sheets("NSP���̾ƿ�").Cells(3 + n, 19) = 5 Then '�ž�ġ���
premperiod = 0
ElseIf renew <> 0 Then
premperiod = insperiod  '���ԱⰣ=����Ⱓ
ElseIf renew = 0 Then
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
End If
'�ڡڡڡڡڡڡڡڡڡڰ��Կ��ɹ���

'������
If Sheets("NSP���̾ƿ�").Cells(3 + n, 19) = 5 Then '�ž�ġ���
age_s = 15
age_e = 100

'������ ���ʰ��
ElseIf renew = 1 Then
age_e = Sheets("����").Cells(3 + n, 7 + renewperi)
age_s = Sheets("����").Cells(3 + n, 3 + renewperi)

'������ ���Ű��
ElseIf renew = 2 Then
age_e = mangi - insperiod
If insperiod = �����ֱ� Then
age_s = Sheets("����").Cells(3 + n, 3 + renewperi) + insperiod
Else
age_s = age_e
End If
ElseIf gubun = "02" Then
age_s = Sheets("����").Cells(3 + n, 3 + ipno_n + 24 * (1 - (jong Mod 2)))
age_e = Sheets("����").Cells(3 + n, 7 + ipno_n + 24 * (1 - (jong Mod 2)))
Else
age_s = Sheets("����").Cells(3 + n, 3 + ipno_n + 24 * (1 - (jong Mod 2)) + 8 * (10 - mangi / 10))
age_e = Sheets("����").Cells(3 + n, 7 + ipno_n + 24 * (1 - (jong Mod 2)) + 8 * (10 - mangi / 10))
End If

'�ڡڡڡڡڡڡڡڡڡڡڡڡڡڱ��ؿ���
If gubun = "02" And (age_s <= 40 And age_e >= 40) Then
age_s_check = 40
ElseIf (gubun = "02" And (age_s > 40 Or age_e < 40)) Or gubun = "01" Then '������&40���ʰ� ����, ������
'age_s_check = 40
age_s_check = Application.Round((age_s + age_e) / 2, 0)
End If

If �������� = 2 Then
age_s = age_s_check
age_e = age_s
End If

'�ڡڡڡڡ� ���Կ���

For lev = 1 To lev_e '�޼�
If �������� = 2 And lev <> 1 Then GoTo ����lev

For age = age_s To age_e
If renew = 2 And renewperi = renewperi_e And (insperiod = 3 Or insperiod = 10) And insperiod <> renewperiod(renewperi) Then GoTo ����age
'-----------------------------------------------------------------------------
If �������� = 1 Then
�������μ�(n) = �������μ�(n) + 1
ElseIf �������� = 2 Then
Sheets("���ؿ���").Cells(3 + n, 3 + IIf(gubun = "01", (mangi / 10 - 7), IIf(renew = 0, ipno_n, renewperi))) = age

If renew = 0 And Sheets("NSP���̾ƿ�").Cells(3 + n, 19) <> 5 Then
���Ը��� = 1
Else
���Ը��� = 0
End If
If gubun = "02" And renew = 0 Then
insperiod = premperiod
ElseIf gubun = "01" Then
insperiod = mangi - age
End If
handoprem = Application.Min(20, insperiod)
������ = 0
Call �����
Call �������
Call ���������

dx�����(0) = 100000
For i = 0 To (insperiod - 1)
x = age + i
cx�����(i) = dx�����(i) * qx�����(sex, lev, x) * v ^ 0.5
dx�����(i + 1) = dx�����(i) * (1 - qx�����(sex, lev, x)) * v
Next i
mx����� = 0
For j = 0 To (insperiod - 1)
mx����� = mx����� + cx�����(j)
Next j
nx����� = 0
For j = 0 To (premperiod - 1)
nx����� = nx����� + dx�����(j)
Next j
nx��������� = nx����� - 11 / 24 * (dx�����(0) - dx�����(premperiod))
����������� = mx����� / (nx��������� * 12)
s = ������(1) / �����������
Call s���
'---------------------------------------------------------------------------------
If (jong = 1 Or jong = 3) And renew = 0 And Sheets("NSP���̾ƿ�").Cells(3 + n, 19) <> 5 Then
Call �����������
Call �ѵ�üũ '��/������ s�����
Call �غ�ݰ�� '��/������ s�����
Call ǥ���غ�ݰ�� '��/������
������ = 1
Call �����
Call �������
Call ���������
s = ������(1) / �����������
Call s���
End If

End If
'---------------------------------------------------------------------------------
����age:
Next age

����lev:
Next lev

��������:
Next ipno_n

��������:
Next ipno_m '��������

��������:
Next ipno_p  '����
Next renewperi '�����ֱ�

����renew:
Next renew
����youl:
Next youl
Next drv

����sex:
Next sex

End If
End If
����n:
Next n
'------------------------------------------------------------------------------------------
If �������� = 3 Then

Sheets("����").Range("J12") = Now

'�ؽ�Ʈ �� ����
Dim DSN As String
Dim DATA_FIELD As String
Open ThisWorkbook.Path & "\���� �����.txt" For Output As #1 '����püũ
MsgBox ("P���̺�input")
DSN = Application.GetOpenFilename("TEXT files(*.TXT),(*TXT)", , "������ ���е� �ؽ�Ʈ ����")
If DSN = "FALSE" Then Exit Sub
Sheets("����").Range("J13") = Now
Open DSN For Input As #11
Do Until EOF(11)
Line Input #11, DATA_FIELD

covcode = Mid(DATA_FIELD, 11, 8)
jong = Mid(DATA_FIELD, 19, 3)
sex = Application.IfError(Val(Mid(DATA_FIELD, 29, 2)), 1)
insperiod = CInt(Mid(DATA_FIELD, 49, 2))
handoprem = Application.Min(20, insperiod)
premperiod = CInt(Mid(DATA_FIELD, 52, 2))
drv = Application.IfError(Val(Mid(DATA_FIELD, 65, 2)), 0)
renew = Application.IfError(Val(Mid(DATA_FIELD, 75, 2)), 0)
zz = Mid(DATA_FIELD, 105, 1)
If zz = "*" Then
lev = 1
Else
lev = Application.IfError(Val(Mid(DATA_FIELD, 105, 1)), 0)
End If
gubun = Mid(DATA_FIELD, 115, 2)
mangi = Mid(DATA_FIELD, 125, 3)
c = Application.IfError(Val(Mid(DATA_FIELD, 128, 4)), 0)
��ǰp = Val(Mid(DATA_FIELD, 197, 10))
Select Case c
Case 0
youl = 0
Case 101
youl = 1
Case 102
youl = 2
Case 103
youl = 3
Case 104
youl = 4
Case 105
youl = 5
Case 106
youl = 6
Case 13
youl = 1
End Select
age = Application.IfError(Val(Mid(DATA_FIELD, 148, 3)), 0)

Set ws = Worksheets("NSP���̾ƿ�")

n = Application.WorksheetFunction.VLookup(covcode, ws.Range("A4:B1166"), 2, False)
�㺸�� = Sheets("NSP���̾ƿ�").Cells(3 + n, 4)
calc_type = Sheets("NSP���̾ƿ�").Cells(3 + n, 60) '���������
AA = Sheets("NSP���̾ƿ�").Cells(3 + n, 61) '�ʳ⵵ Cx�� ���� ����(����/��å)
face = Sheets("NSP���̾ƿ�").Cells(3 + n, 27) * 10000 '���رݾ�
sex_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 28) '�ѵ�üũ ���� ����
si = Sheets("NSP���̾ƿ�").Cells(3 + n, 55)
nn = Sheets("NSP���̾ƿ�").Cells(3 + n, 42 + jong)
If renew = 1 Then    '���Ŵ㺸 ���ʰ��
ipno_n = 0
Select Case insperiod
Case 3
renewperi = 1
Case 10
renewperi = 2
Case 20
renewperi = 3
End Select
ElseIf renew = 2 Then  '���Ŵ㺸 ���Ű��, renewperi�� 0���� ��Ƶ� ���� ������ �ѹ� Ȯ���غ���(������ Į������ ���� �Ѵ�..)
renewperi = 0
ipno_n = 0
ElseIf renew = 0 Then  '���ź񰻽� ������ �ž�ġ���, ������񰻽�(��������), ������
renewperi = 0
Select Case premperiod
Case 0
ipno_n = 0
Case 10
ipno_n = 1
Case 15
ipno_n = 2
Case 20
ipno_n = 3
Case 30
ipno_n = 4
End Select
End If
If renew = 0 And Sheets("NSP���̾ƿ�").Cells(3 + n, 15) <> 5 Then
���Ը��� = 1
Else
���Ը��� = 0
End If
������ = 0
Call �����
Call �������
Call ���������

If (jong = 1 Or jong = 3) And renew = 0 And Sheets("NSP���̾ƿ�").Cells(3 + n, 19) <> 5 Then
Call �����������
Call �ѵ�üũ '��/������
Call �غ�ݰ�� '��/������
Call ǥ���غ�ݰ�� '��/������
������ = 1
Call �����
Call �������
Call ���������
End If

Call �����������
'Call ���

If ��������1�� <> ��ǰp Then Call P���
��ǰ���μ�(n) = ��ǰ���μ�(n) + 1

Loop
Close #1
Close #11
'-----------------------------------------------------------------------------------------------------
ElseIf �������� = 4 Then

Sheets("����").Range("J12") = Now

Dim DSNV As String
Dim DATA_FIELD_V As String
Open ThisWorkbook.Path & "\���� �غ��.txt" For Output As #2  '�غ��üũ. ��püũ, �ѵ���ġ üũ
Open ThisWorkbook.Path & "\�ѵ�üũ.txt" For Output As #3  '�Ű����ѵ��ʰ� üũ
MsgBox ("V���̺�input")
DSNV = Application.GetOpenFilename("TEXT files(*.TXT),(*TXT)", , "������ ���е� �ؽ�Ʈ ����")
If DSNV = "FALSE" Then Exit Sub
Sheets("����").Range("J13") = Now
Open DSNV For Input As #11
Do Until EOF(11)
Line Input #11, DATA_FIELD_V

covcode = Mid(DATA_FIELD_V, 11, 8)
jong = Mid(DATA_FIELD_V, 19, 3)
sex = Application.IfError(Val(Mid(DATA_FIELD_V, 29, 2)), 1)
insperiod = CInt(Mid(DATA_FIELD_V, 49, 2))
handoprem = Application.Min(20, insperiod)
premperiod = CInt(Mid(DATA_FIELD_V, 52, 2))
drv = Application.IfError(Val(Mid(DATA_FIELD_V, 65, 2)), 0)
renew = Application.IfError(Val(Mid(DATA_FIELD_V, 75, 2)), 0)
zz = Mid(DATA_FIELD_V, 105, 1)
If zz = "*" Then
lev = 1
Else
lev = Application.IfError(Val(Mid(DATA_FIELD_V, 105, 1)), 0)
End If
gubun = Mid(DATA_FIELD_V, 115, 2)
mangi = Mid(DATA_FIELD_V, 125, 3)
c = Application.IfError(Val(Mid(DATA_FIELD_V, 128, 4)), 0)
If mangi = 5 Then
��ǰnp = Val(Mid(DATA_FIELD_V, 1847, 10)) ''' ���� ����
Else
��ǰnp = Val(Mid(DATA_FIELD_V, 1862, 10)) ''' ���� ����
End If
��ǰ�ѵ� = Val(Mid(DATA_FIELD_V, 3692, 10)) ''' �ѵ� ��
Sum_��ǰV = 0
For t = 0 To insperiod
��ǰV(t) = Mid(DATA_FIELD_V, 182 + t * 15, 15)
Sum_��ǰV = Sum_��ǰV + ��ǰV(t)
Next t
Select Case c
Case 0
youl = 0
Case 101
youl = 1
Case 102
youl = 2
Case 103
youl = 3
Case 104
youl = 4
Case 105
youl = 5
Case 106
youl = 6
Case 13
youl = 1
End Select
age = Application.IfError(Val(Mid(DATA_FIELD_V, 148, 3)), 0)

Set ws = Worksheets("NSP���̾ƿ�")

n = Application.WorksheetFunction.VLookup(covcode, ws.Range("A4:B1166"), 2, False)
�㺸�� = Sheets("NSP���̾ƿ�").Cells(3 + n, 4)
calc_type = Sheets("NSP���̾ƿ�").Cells(3 + n, 60) '���������
AA = Sheets("NSP���̾ƿ�").Cells(3 + n, 61) '�ʳ⵵ Cx�� ���� ����(����/��å)
face = Sheets("NSP���̾ƿ�").Cells(3 + n, 27) * 10000 '���رݾ�
sex_s = Sheets("NSP���̾ƿ�").Cells(3 + n, 28) '�ѵ�üũ ���� ����
si = Sheets("NSP���̾ƿ�").Cells(3 + n, 55)
nn = Sheets("NSP���̾ƿ�").Cells(3 + n, 42 + jong)
If renew = 1 Then    '���Ŵ㺸 ���ʰ��
ipno_n = 0
Select Case insperiod
Case 3
renewperi = 1
Case 10
renewperi = 2
Case 20
renewperi = 3
End Select
ElseIf renew = 2 Then  '���Ŵ㺸 ���Ű��, renewperi�� 0���� ��Ƶ� ���� ������ �ѹ� Ȯ���غ���(������ Į������ ���� �Ѵ�..)
renewperi = 0
ipno_n = 0
ElseIf renew = 0 Then  '���ź񰻽� ������ �ž�ġ���, ������񰻽�(��������), ������
renewperi = 0
Select Case premperiod
Case 0
ipno_n = 0
Case 10
ipno_n = 1
Case 15
ipno_n = 2
Case 20
ipno_n = 3
Case 30
ipno_n = 4
End Select
End If

If renew = 0 And Sheets("NSP���̾ƿ�").Cells(3 + n, 19) <> 5 Then
���Ը��� = 1
Else
���Ը��� = 0
End If
������ = 0
Call �����
Call �������
Call ���������

If (jong = 1 Or jong = 3) And renew = 0 And Sheets("NSP���̾ƿ�").Cells(3 + n, 19) <> 5 Then
Call �����������
Call �ѵ�üũ '��/������
Call �غ�ݰ�� '��/������
Call ǥ���غ�ݰ�� '��/������
������ = 1
Call �����
Call �������
Call ���������
End If

Call �����������
Call �ѵ�üũ
Call �غ�ݰ��

Sum_����V = 0
For t = 0 To insperiod
Sum_����V = Sum_����V + �غ��1��(t, 1)
Next t
If Sum_����V <> Sum_��ǰV Or Int(�Ű����ѵ�) <> ��ǰ�ѵ� Or ��p <> ��ǰnp Or ��ǰnp = 0 Then Call V���

If Int(�Ű����ѵ�) < ���Ű��� And sex = sex_s And lev = 1 And (renew = 1 Or (renew = 0 And gubun = "02") Or (renew = 0 And gubun = "01" And ipno_n = 4)) Then
If Sheets("���ؿ���").Cells(3 + n, 3 + IIf(gubun = "01", (mangi / 10 - 7), IIf(renew = 0, ipno_n, renewperi))) = age Then Call �ѵ����
End If
'Call ���
Loop

Close #2
Close #3
Close #11

End If

If �������� = 1 Then

For i = Sheets("����").Range("J4") To Sheets("����").Range("J5")
Sheets("����").Cells(5 + i, 7) = �������μ�(i)
Next i

ElseIf �������� = 3 Then
For i = Sheets("����").Range("J4") To Sheets("����").Range("J5")
Sheets("����").Cells(5 + i, 6) = ��ǰ���μ�(i)
Next i

End If

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Sheets("����").Range("J14") = Now '����


MsgBox "�Ϸ�Ǿ����ϴ�"
End Sub