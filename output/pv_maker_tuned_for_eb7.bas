Option Explicit

Public start As Integer, last As Integer, 산출종류 As Integer, s산출여부 As Integer, s사용여부 As Integer, youl_s As Integer, youl_e As Integer, youl As Integer, a As Integer, b As Integer, c As Integer, 무해지 As Integer
Public alpha1(1163, 6, 2, 3, 4) As Double, alpha2(1163, 6, 2, 30) As Double, beta(1163, 6, 2) As Double, beta5 As Double, ce As Double, beta1 As Double, ce1 As Double
Public z(110) As Double, 갱신s값(1163, 6) As Double, s값(1163, 2, 6, 3, 3, 4) As Double, 계수(6, 1163, 2, 110) As Double, qx산출(1163, 5, 2, 3, 110) As Double, qx무배당(2, 3, 110) As Double, s As Double, 상해(6, 5, 2, 3, 110) As Double, 질병(6, 5, 2, 3, 110) As Double, 암(6, 5, 2, 3, 110) As Double, 뇌졸중(6, 5, 2, 3, 110) As Double, 급성(6, 5, 2, 3, 110) As Double, 기갑(6, 5, 2, 3, 110) As Double, 상해성뇌출혈(6, 5, 2, 3, 110) As Double, 유방암(6, 5, 2, 3, 110) As Double, 기타(6, 5, 2, 3, 110) As Double, 갑상선(6, 5, 2, 3, 110) As Double
Public dx(5, 110) As Double, dx납면(5, 110) As Double, dx무배당(110) As Double, AA As Double, dx1(5, 110) As Double, dx2(5, 110) As Double, dx_표준(5, 110) As Double, dx납면_표준(5, 110) As Double, dx1_표준(5, 110) As Double, dx2_표준(5, 110) As Double
Public cx(5, 110) As Double, cx1(5, 110) As Double, cx2(5, 110) As Double, cx무배당(110) As Double, cx_표준(5, 110) As Double, cx1_표준(5, 110) As Double, cx2_표준(5, 110) As Double
Public rate(4) As Double, v As Double
Public txt_num As Integer, i As Integer, j As Integer, j_e As Integer, jj As Integer, n As Integer, k As Integer, kk As Integer, kkk As Integer, no As Integer, t As Integer, irate As Integer, x As Integer, mangi_k As Integer, mangi_k_s As Integer, mangi_k_e As Integer, 납입면제 As Integer
Public face As Long
Public sex As Integer, sex_check As Integer, sex_s As Integer, sex_e As Integer, cl_s As Integer, cl_e As Integer, cl As Integer, drv_s As Integer, drv_e As Integer, drv As Integer, lev_e As Integer, lev As Integer, age_s As Integer, age_e As Integer, age As Integer, age_s_check As Integer, renew_s As Integer, renew_e As Integer, renew As Integer, renewperi As Integer, renewperi_s As Integer, renewperi_e As Integer, 갱신주기 As Integer, si As Integer
Public ipno_p_s As Integer, ipno_p_e As Integer, ipno_p As Integer, ipno_m_s As Integer, ipno_m_e As Integer, ipno_m As Integer, ipno_n_e As Integer, ipno_n As Integer, mangi_type As Integer, mangi As Integer, ipno_n_s As Integer
Public insperiod As Integer, premperiod As Integer, handoprem As Integer
Public nx납면(110) As Double, nx납면_표준(110) As Double, nx(110) As Double, nx_표준(110) As Double, nx한도(110) As Double, nx월납 As Double, nx무배당 As Double, nx월납_표준 As Double, nx월납무배당 As Double, bunja_p As Double, mx(110) As Double, bunja_p_표준 As Double, mx_표준(110) As Double, mx무배당 As Double, mxsum As Double, nxsum As Double, nxsum_k As Double, mxsum_표준 As Double, nxsum_표준 As Double, nxsum_k_표준 As Double
Public 순연납무배당 As Double, 순월납무배당 As Double, 순연납(1) As Double, 순월납(1) As Double, 영업연납 As Double, 영업월납 As Double, 순연납_표준(1) As Double, 순월납_표준(1) As Double, 영업연납_표준 As Double, 영업월납_표준 As Double
Public 순한도(1) As Long, 순한도_표준(1) As Double As Long, 순한도1원(1) As Long, 순한도1원_표준(1) As Long As Long, 해약공제계수 As Long, 해지공제기간 As Long
Public 순연납1원(1) As Long, 순월납1원(1) As Long, 영업연납1원 As Long, 영업월납1원 As Long, 순연납1원_표준(1) As Long, 순월납1원_표준(1) As Long, 영업연납1원_표준 As Long, 영업월납1원_표준 As Long, 순연납무배당1원 As Long, 순월납무배당1원 As Long, 상품p As Long, 상품np As Long, Sum_상품V As Long, Sum_계지V As Long, 상품한도 As Long, 순p As Long
Public 준비금(110, 1) As Double, 준비금_표준(110, 1) As Double, 환급금_표준(110, 1) As Double
Public 준비금1원(110, 1) As Double, 준비금1원_표준(110, 1) As Double
Public plancode As String, covcode As String, zz As String
Public 담보명 As String
Public nn As Integer
Public j1 As Integer, jj1 As Integer

Public 해지율(1, 4, 110) As Double, w_rate(1, 4, 110) As Double
Public 상품V(110) As Long
Public txt(430) As String
Public hh As Integer, hhh As Integer, scheck As Integer
Public 신계약비한도 As Double, 사용신계약비 As Double, 신계약비한도_표준 As Double, 사용신계약비_표준 As Double, 전체사업비율 As Double
Public i2 As Integer, mm As Integer
Public 해지환급금(21) As Long, USUM As Long
Public 상각기간 As Integer, 산출여부 As Integer
Public gubun As String
Public 경로명 As String, 파일명 As String
Public 계지라인수(1163) As Long, 상품라인수(1163) As Long
Public 요율구분 As String
Public renewperiod() As Variant
Public jong As Integer, jong_k As Integer, jong_s As Integer, jong_e As Integer, jong_ss As Integer, jong_ee As Integer, n_rate As Integer, n_rate_k As Integer, n_rate_r As Integer, n_rate_c(5) As Integer
Public calc_type As Integer
Public ws As Worksheet
Public Rng1 As Range
Sub 산출()
    
Application.StatusBar = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Sheets("산출").Range("J11") = Now '시작시간
renewperiod() = Array(0, 3, 10, 20) '갱신형 갱신주기
irate = 1 ' 적용/표준V(표준i+가산)/해지V(표준i*125%)/표준해지공제액(한도체크용,표준i) ''15.01수정
rate(1) = Sheets("산출").Range("J7") '예정이율
v = 1 / (1 + rate(irate))
k = 3 '출력용
산출종류 = Sheets("산출").Range("J9") '1: 계지건수체크(전체종) 2: s산출(전체종) 3: p검증 4: v검증

If 산출종류 = 1 Then
  For a = Sheets("산출").Range("J4") To Sheets("산출").Range("J5")
    계지라인수(a) = 0
  Next a
ElseIf 산출종류 = 3 Then
  For a = Sheets("산출").Range("J4") To Sheets("산출").Range("J5")
    상품라인수(a) = 0
  Next a
End If

jong = CInt(Sheets("산출").Range("J20"))

If 산출종류 > 1 Then
Call 정기사망률인식
Call 납면위험률인식
Call 해지율인식
End If

For n = Sheets("산출").Range("J4") To Sheets("산출").Range("J5")  '테이블세팅

jong_s = CInt(Sheets("NSP레이아웃").Cells(3 + n, 39))
jong_e = CInt(Sheets("NSP레이아웃").Cells(3 + n, 40))

If (jong < jong_s Or jong > jong_e) Then GoTo 다음n

If 산출종류 = 2 Then
  산출여부 = Sheets("NSP레이아웃").Cells(3 + n, 42)
Else
  산출여부 = 1
End If

If 산출여부 = 1 Then


요율구분 = Sheets("NSP레이아웃").Cells(3 + n, 20) '* 구분없음 0 표준체 1 표준하체 건강등급그룹코드ABCD
covcode = Sheets("NSP레이아웃").Cells(3 + n, 6)
nn = Sheets("NSP레이아웃").Cells(3 + n, 42 + jong) '사업비 위치

calc_type = Sheets("NSP레이아웃").Cells(3 + n, 60) '계산기수유형
AA = Sheets("NSP레이아웃").Cells(3 + n, 61) '초년도 Cx에 곱할 비율(감액/면책)

gubun = Sheets("NSP레이아웃").Cells(3 + n, 18)    '만기구분코드(01:세만기,02:연만기,*:태아 월)

sex_s = Sheets("NSP레이아웃").Cells(3 + n, 28)        '0 구분없음 1 남자 2 여자
sex_e = Sheets("NSP레이아웃").Cells(3 + n, 29)       '성별

drv_s = Sheets("NSP레이아웃").Cells(3 + n, 31)
drv_e = Sheets("NSP레이아웃").Cells(3 + n, 32)

renew_s = Sheets("NSP레이아웃").Cells(3 + n, 33)
renew_e = Sheets("NSP레이아웃").Cells(3 + n, 34)

renewperi_s = Sheets("NSP레이아웃").Cells(3 + n, 62)
renewperi_e = Sheets("NSP레이아웃").Cells(3 + n, 63)

mangi_k_s = Sheets("NSP레이아웃").Cells(3 + n, 65)
mangi_k_e = Sheets("NSP레이아웃").Cells(3 + n, 66)

n_rate = Sheets("NSP레이아웃").Cells(3 + n, 47)  '위험률유형
n_rate_k = Sheets("NSP레이아웃").Cells(3 + n, 48)  '위험률갯수
n_rate_r = Sheets("NSP레이아웃").Cells(3 + n, 49)  '위험률행
si = Sheets("NSP레이아웃").Cells(3 + n, 55)  '조정계수

ipno_n_e = Sheets("NSP레이아웃").Cells(3 + n, 36)
ipno_n_s = Sheets("NSP레이아웃").Cells(3 + n, 35)

lev_e = Sheets("NSP레이아웃").Cells(3 + n, 38)
    
For kk = 1 To n_rate_k
  n_rate_c(kk) = Sheets("NSP레이아웃").Cells(3 + n, 49 + kk)
Next kk

face = Sheets("NSP레이아웃").Cells(3 + n, 27) * 10000 '기준금액
If 요율구분 = "0013" Then
  youl_s = 1
  youl_e = 1
Else
  youl_s = 0
  youl_e = 0
End If
'-----------------------------------------------------------------------테이블화 이것까진 s산출이든 p산출이든 v산출이든 다해야함(건수체크 제외)
If 산출종류 > 1 Then
Call 조정계수인식
Call 사업비율인식
Call 위험률인식
End If

'-------------------------------------------------------------------------
If 산출종류 < 3 Then
'-------------------------------------------------------------------------
For sex = sex_s To sex_e '성별
If 산출종류 = 2 And (sex = 2 And sex_s <> 2) Then GoTo 다음sex 'S산출/여성보험고려: 기준을 여성으로 잡아야함

For drv = drv_s To drv_e                        '운전형태
For youl = youl_s To youl_e '요율구분
For renew = renew_s To renew_e
If 산출종류 = 2 And renew = 2 Then GoTo 다음renew

  For renewperi = renewperi_s To renewperi_e
    If renew = 0 Then
      갱신주기 = 0
    Else
      갱신주기 = renewperiod(renewperi)
    End If
     mangi_type = Sheets("NSP레이아웃").Cells(3 + n, 64)
    
    '★★★★★★★★★★★★★★보기종류
    If Sheets("NSP레이아웃").Cells(3 + n, 19) = 5 Then '신암치료비
      ipno_p_s = 1
      ipno_p_e = 1
    ElseIf renew = 1 Then '갱신형 최초계약
      ipno_p_s = 갱신주기
      ipno_p_e = 갱신주기
    ElseIf renew = 2 And renewperi = renewperi_e Then '갱신형 갱신계약
      ipno_p_s = 1
      ipno_p_e = 갱신주기
    ElseIf renew = 0 Then '세만기&연만기비갱신
      ipno_p_s = 1
      ipno_p_e = 1
    Else
      ipno_p_s = 갱신주기
      ipno_p_e = 갱신주기
    End If
    
  For ipno_p = ipno_p_s To ipno_p_e '★★★보험기간
    If Sheets("NSP레이아웃").Cells(3 + n, 19) = 5 Then
      insperiod = 5
    ElseIf renew <> 0 Then
      insperiod = ipno_p
    End If
      
    If mangi_type = 0 Then '태아보장용 보기/납기=0, P값 출력은 월납에
        ipno_m_s = 1
        ipno_m_e = 1
    Else
        If mangi_type = 1 Then ipno_m_e = 2
        If mangi_type = 2 Then ipno_m_e = 1
        If mangi_type = 3 Then ipno_m_e = 3
        If gubun = "01" Then
          ipno_m_s = 1
        ElseIf insperiod = 갱신주기 Then
          ipno_m_s = ipno_m_e
        Else
          ipno_m_s = 1
        End If
    End If
  For ipno_m = ipno_m_s To ipno_m_e '★★★만기
    
    If Sheets("NSP레이아웃").Cells(3 + n, 19) = 5 Then
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
    
  For ipno_n = ipno_n_s To ipno_n_e '납기종류
  If 산출종류 = 2 And gubun = "01" And ipno_n <> ipno_n_e Then GoTo 다음납기
  If (jong = 1 Or jong = 3) And ipno_n = 1 Then GoTo 다음납기
      If Sheets("NSP레이아웃").Cells(3 + n, 19) = 5 Then '신암치료비
        premperiod = 0
      ElseIf renew <> 0 Then
        premperiod = insperiod  '납입기간=보험기간
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
    '★★★★★★★★★★가입연령범위
    
    '연만기
    If Sheets("NSP레이아웃").Cells(3 + n, 19) = 5 Then '신암치료비
      age_s = 15
      age_e = 100
            
   '갱신형 최초계약
    ElseIf renew = 1 Then
      age_e = Sheets("연령").Cells(3 + n, 7 + renewperi)
      age_s = Sheets("연령").Cells(3 + n, 3 + renewperi)
      
    '갱신형 갱신계약
    ElseIf renew = 2 Then
      age_e = mangi - insperiod
      If insperiod = 갱신주기 Then
        age_s = Sheets("연령").Cells(3 + n, 3 + renewperi) + insperiod
      Else
        age_s = age_e
      End If
    ElseIf gubun = "02" Then
      age_s = Sheets("연령").Cells(3 + n, 3 + ipno_n + 24 * (1 - (jong Mod 2)))
      age_e = Sheets("연령").Cells(3 + n, 7 + ipno_n + 24 * (1 - (jong Mod 2)))
    Else
      age_s = Sheets("연령").Cells(3 + n, 3 + ipno_n + 24 * (1 - (jong Mod 2)) + 8 * (10 - mangi / 10))
      age_e = Sheets("연령").Cells(3 + n, 7 + ipno_n + 24 * (1 - (jong Mod 2)) + 8 * (10 - mangi / 10))
    End If
    
   '★★★★★★★★★★★★★★기준연령
    If gubun = "02" And (age_s <= 40 And age_e >= 40) Then
      age_s_check = 40
    ElseIf (gubun = "02" And (age_s > 40 Or age_e < 40)) Or gubun = "01" Then '연만기&40세초과 산출, 세만기
      'age_s_check = 40
      age_s_check = Application.Round((age_s + age_e) / 2, 0)
    End If
    
    If 산출종류 = 2 Then
      age_s = age_s_check
      age_e = age_s
    End If
                       
'★★★★★ 가입연령

    For lev = 1 To lev_e '급수
    If 산출종류 = 2 And lev <> 1 Then GoTo 다음lev
    
    For age = age_s To age_e
    If renew = 2 And renewperi = renewperi_e And (insperiod = 3 Or insperiod = 10) And insperiod <> renewperiod(renewperi) Then GoTo 다음age
    '-----------------------------------------------------------------------------
    If 산출종류 = 1 Then
      계지라인수(n) = 계지라인수(n) + 1
    ElseIf 산출종류 = 2 Then
      Sheets("기준연령").Cells(3 + n, 3 + IIf(gubun = "01", (mangi / 10 - 7), IIf(renew = 0, ipno_n, renewperi))) = age
      
      If renew = 0 And Sheets("NSP레이아웃").Cells(3 + n, 19) <> 5 Then
      납입면제 = 1
      Else
      납입면제 = 0
      End If
      If gubun = "02" And renew = 0 Then
        insperiod = premperiod
      ElseIf gubun = "01" Then
        insperiod = mangi - age
      End If
    handoprem = Application.Min(20, insperiod)
    무해지 = 0
    Call 계산기수
    Call 계산기수합
    Call 순보험료계산

    dx무배당(0) = 100000
    For i = 0 To (insperiod - 1)
      x = age + i
      cx무배당(i) = dx무배당(i) * qx무배당(sex, lev, x) * v ^ 0.5
      dx무배당(i + 1) = dx무배당(i) * (1 - qx무배당(sex, lev, x)) * v
    Next i
    mx무배당 = 0
    For j = 0 To (insperiod - 1)
      mx무배당 = mx무배당 + cx무배당(j)
    Next j
    nx무배당 = 0
    For j = 0 To (premperiod - 1)
      nx무배당 = nx무배당 + dx무배당(j)
    Next j
    nx월납무배당 = nx무배당 - 11 / 24 * (dx무배당(0) - dx무배당(premperiod))
    순월납무배당 = mx무배당 / (nx월납무배당 * 12)
    s = 순월납(1) / 순월납무배당
    Call s출력
    '---------------------------------------------------------------------------------
    If (jong = 1 Or jong = 3) And renew = 0 And Sheets("NSP레이아웃").Cells(3 + n, 19) <> 5 Then
    Call 영업보험료계산
    Call 한도체크 '무/저해지 s산출용
    Call 준비금계산 '무/저해지 s산출용
    Call 표준준비금계산 '무/저해지
    무해지 = 1
    Call 계산기수
    Call 계산기수합
    Call 순보험료계산
    s = 순월납(1) / 순월납무배당
    Call s출력
    End If
    
    End If
    '---------------------------------------------------------------------------------
다음age:
Next age

다음lev:
Next lev

다음납기:
Next ipno_n

다음만기:
Next ipno_m '갱신종료

다음보기:
Next ipno_p  '보기
Next renewperi '갱신주기

다음renew:
Next renew
다음youl:
Next youl
Next drv

다음sex:
Next sex

End If
End If
다음n:
Next n
'------------------------------------------------------------------------------------------
If 산출종류 = 3 Then

Sheets("산출").Range("J12") = Now

'텍스트 비교 시작
Dim DSN As String
Dim DATA_FIELD As String
Open ThisWorkbook.Path & "\검증 보험료.txt" For Output As #1 '영업p체크
MsgBox ("P테이블input")
DSN = Application.GetOpenFilename("TEXT files(*.TXT),(*TXT)", , "탭으로 구분된 텍스트 파일")
If DSN = "FALSE" Then Exit Sub
Sheets("산출").Range("J13") = Now
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
상품p = Val(Mid(DATA_FIELD, 197, 10))
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

Set ws = Worksheets("NSP레이아웃")

n = Application.WorksheetFunction.VLookup(covcode, ws.Range("A4:B1166"), 2, False)
담보명 = Sheets("NSP레이아웃").Cells(3 + n, 4)
calc_type = Sheets("NSP레이아웃").Cells(3 + n, 60) '계산기수유형
AA = Sheets("NSP레이아웃").Cells(3 + n, 61) '초년도 Cx에 곱할 비율(감액/면책)
face = Sheets("NSP레이아웃").Cells(3 + n, 27) * 10000 '기준금액
sex_s = Sheets("NSP레이아웃").Cells(3 + n, 28) '한도체크 기준 성별
si = Sheets("NSP레이아웃").Cells(3 + n, 55)
nn = Sheets("NSP레이아웃").Cells(3 + n, 42 + jong)
If renew = 1 Then    '갱신담보 최초계약
  ipno_n = 0
  Select Case insperiod
    Case 3
      renewperi = 1
    Case 10
      renewperi = 2
    Case 20
      renewperi = 3
  End Select
ElseIf renew = 2 Then  '갱신담보 갱신계약, renewperi를 0으로 잡아도 문제 없는지 한번 확인해보기(원래는 칼럼값에 들어가긴 한다..)
  renewperi = 0
  ipno_n = 0
ElseIf renew = 0 Then  '갱신비갱신 무관한 신암치료비, 연만기비갱신(납입지원), 세만기
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
If renew = 0 And Sheets("NSP레이아웃").Cells(3 + n, 15) <> 5 Then
      납입면제 = 1
      Else
      납입면제 = 0
      End If
무해지 = 0
Call 계산기수
Call 계산기수합
Call 순보험료계산

If (jong = 1 Or jong = 3) And renew = 0 And Sheets("NSP레이아웃").Cells(3 + n, 19) <> 5 Then
Call 영업보험료계산
Call 한도체크 '무/저해지
Call 준비금계산 '무/저해지
Call 표준준비금계산 '무/저해지
무해지 = 1
Call 계산기수
Call 계산기수합
Call 순보험료계산
End If

Call 영업보험료계산
'Call 출력

If 영업월납1원 <> 상품p Then Call P출력
  상품라인수(n) = 상품라인수(n) + 1
  
Loop
  Close #1
  Close #11
'-----------------------------------------------------------------------------------------------------
ElseIf 산출종류 = 4 Then

Sheets("산출").Range("J12") = Now

Dim DSNV As String
Dim DATA_FIELD_V As String
Open ThisWorkbook.Path & "\검증 준비금.txt" For Output As #2  '준비금체크. 순p체크, 한도일치 체크
Open ThisWorkbook.Path & "\한도체크.txt" For Output As #3  '신계약비한도초과 체크
MsgBox ("V테이블input")
DSNV = Application.GetOpenFilename("TEXT files(*.TXT),(*TXT)", , "탭으로 구분된 텍스트 파일")
If DSNV = "FALSE" Then Exit Sub
Sheets("산출").Range("J13") = Now
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
상품np = Val(Mid(DATA_FIELD_V, 1847, 10)) ''' 월납 순보
Else
상품np = Val(Mid(DATA_FIELD_V, 1862, 10)) ''' 월납 순보
End If
상품한도 = Val(Mid(DATA_FIELD_V, 3692, 10)) ''' 한도 값
Sum_상품V = 0
  For t = 0 To insperiod
    상품V(t) = Mid(DATA_FIELD_V, 182 + t * 15, 15)
    Sum_상품V = Sum_상품V + 상품V(t)
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

Set ws = Worksheets("NSP레이아웃")

n = Application.WorksheetFunction.VLookup(covcode, ws.Range("A4:B1166"), 2, False)
담보명 = Sheets("NSP레이아웃").Cells(3 + n, 4)
calc_type = Sheets("NSP레이아웃").Cells(3 + n, 60) '계산기수유형
AA = Sheets("NSP레이아웃").Cells(3 + n, 61) '초년도 Cx에 곱할 비율(감액/면책)
face = Sheets("NSP레이아웃").Cells(3 + n, 27) * 10000 '기준금액
sex_s = Sheets("NSP레이아웃").Cells(3 + n, 28) '한도체크 기준 성별
si = Sheets("NSP레이아웃").Cells(3 + n, 55)
nn = Sheets("NSP레이아웃").Cells(3 + n, 42 + jong)
If renew = 1 Then    '갱신담보 최초계약
  ipno_n = 0
  Select Case insperiod
    Case 3
      renewperi = 1
    Case 10
      renewperi = 2
    Case 20
      renewperi = 3
  End Select
ElseIf renew = 2 Then  '갱신담보 갱신계약, renewperi를 0으로 잡아도 문제 없는지 한번 확인해보기(원래는 칼럼값에 들어가긴 한다..)
  renewperi = 0
  ipno_n = 0
ElseIf renew = 0 Then  '갱신비갱신 무관한 신암치료비, 연만기비갱신(납입지원), 세만기
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

If renew = 0 And Sheets("NSP레이아웃").Cells(3 + n, 19) <> 5 Then
      납입면제 = 1
      Else
      납입면제 = 0
      End If
무해지 = 0
Call 계산기수
Call 계산기수합
Call 순보험료계산

If (jong = 1 Or jong = 3) And renew = 0 And Sheets("NSP레이아웃").Cells(3 + n, 19) <> 5 Then
Call 영업보험료계산
Call 한도체크 '무/저해지
Call 준비금계산 '무/저해지
Call 표준준비금계산 '무/저해지
무해지 = 1
Call 계산기수
Call 계산기수합
Call 순보험료계산
End If

Call 영업보험료계산
Call 한도체크
Call 준비금계산

  Sum_계지V = 0
  For t = 0 To insperiod
    Sum_계지V = Sum_계지V + 준비금1원(t, 1)
  Next t
  If Sum_계지V <> Sum_상품V Or Int(신계약비한도) <> 상품한도 Or 순p <> 상품np Or 상품np = 0 Then Call V출력
  
  If Int(신계약비한도) < 사용신계약비 And sex = sex_s And lev = 1 And (renew = 1 Or (renew = 0 And gubun = "02") Or (renew = 0 And gubun = "01" And ipno_n = 4)) Then
    If Sheets("기준연령").Cells(3 + n, 3 + IIf(gubun = "01", (mangi / 10 - 7), IIf(renew = 0, ipno_n, renewperi))) = age Then Call 한도출력
  End If
'Call 출력
Loop

  Close #2
  Close #3
  Close #11

End If

If 산출종류 = 1 Then

For i = Sheets("산출").Range("J4") To Sheets("산출").Range("J5")
Sheets("산출").Cells(5 + i, 7) = 계지라인수(i)
Next i

ElseIf 산출종류 = 3 Then
For i = Sheets("산출").Range("J4") To Sheets("산출").Range("J5")
Sheets("산출").Cells(5 + i, 6) = 상품라인수(i)
Next i

End If

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Sheets("산출").Range("J14") = Now '종료


MsgBox "완료되었습니다"
End Sub