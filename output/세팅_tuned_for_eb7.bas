Sub 납면위험률인식()
For youl = 0 To 1
For sex = 1 To 2 '성별
For x = 0 To 110 '가입나이+경과기간
질병(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 170) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 40), 1)
암(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 52) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
뇌졸중(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 108) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 14), 1)
급성(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 122) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 27), 1)
상해성뇌출혈(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 410) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 79), 1)
유방암(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 884) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
기타(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 54) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
갑상선(youl, 1, sex, 1, x) = Sheets("적용위험률").Cells(x + 6, sex + 56) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 1), 1)
For lev = 1 To 3 '급수
상해(youl, 1, sex, lev, x) = Sheets("적용위험률").Cells(x + 6, sex + 8 + 2 * (lev - 1)) * IIf((jong = 1 Or jong = 2) And youl <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + 40), 1)
Next lev
Next x
Next sex
Next youl

End Sub
Sub 정기사망률인식()
For sex = 1 To 2 '성별
For lev = 1 To 1 '급수
For x = 0 To 110 '가입나이+경과기간
qx무배당(sex, lev, x) = Sheets("적용위험률").Cells(x + 6, sex + 6 + 2 * (lev - 1))
Next x
Next lev
Next sex
End Sub
Sub 사업비율인식()
For youl = youl_s To youl_e
For renew = renew_s To renew_e

For premperiod = 1 To 30
If nn = 0 Then
alpha2(n, youl, 0, 0) = 0
Else
Select Case jong
Case 1
alpha2(n, youl, renew, premperiod) = Sheets("예사비1종").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
Case 2
alpha2(n, youl, renew, premperiod) = Sheets("예사비2종").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
Case 3
alpha2(n, youl, renew, premperiod) = Sheets("예사비3종").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
Case 4
alpha2(n, youl, renew, premperiod) = Sheets("예사비4종").Cells(6 + nn + IIf(renew = 2, 1, 0), 34 - premperiod) / 100
End Select
End If
Next premperiod

If nn = 0 Then
beta(n, youl, renew) = 0
Else
Select Case jong
Case 1
beta(n, youl, renew) = Sheets("예사비1종").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
Case 2
beta(n, youl, renew) = Sheets("예사비2종").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
Case 3
beta(n, youl, renew) = Sheets("예사비3종").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
Case 4
beta(n, youl, renew) = Sheets("예사비4종").Cells(6 + nn + IIf(renew = 2, 1, 0), 37) / 100
End Select
End If

For mangi_k = mangi_k_s To mangi_k_e
For ipno_n = ipno_n_s To ipno_n_e

If nn = 0 Then
alpha1(n, youl, 0, 0, 0) = 0
Else
Select Case jong
Case 1
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("예사비1종").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))
Case 2
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("예사비2종").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))
Case 3
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("예사비3종").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))
Case 4
alpha1(n, youl, renew, mangi_k, ipno_n) = Sheets("예사비4종").Cells(6 + nn + IIf(renew = 2, 1, 0), IIf(renew = 0, IIf(mangi_k <> 0, 37 - mangi_k, 37 - ipno_n), 36))

End Select
End If
beta5 = Sheets("산출").Range("J25") '영업보험료대비 수금비
ce = Sheets("산출").Range("J26") '손해조사비
beta1 = Sheets("산출").Range("J27") '(납입후)영업보험료대비 수금비
ce1 = Sheets("산출").Range("J28") '(납입후)손해조사비
Next ipno_n
Next mangi_k
Next renew
Next youl

End Sub
Sub 해지율인식()
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
For 무해지 = 0 To 1
For i = 0 To 110
If 무해지 = 1 Then
If jong = 2 Or jong = 4 Or ipno_n = 0 Then
해지율(무해지, ipno_n, i) = 0
w_rate(무해지, ipno_n, i) = 0
ElseIf i = premperiod - 2 Then '완납2년전
해지율(무해지, ipno_n, i) = Sheets("해지율").Cells(5, 19)
w_rate(무해지, ipno_n, i) = 0 '해지환급금 지급률
ElseIf i = premperiod - 1 Then '완납1년전
해지율(무해지, ipno_n, i) = Sheets("해지율").Cells(5, 20)
w_rate(무해지, ipno_n, i) = 0 '해지환급금 지급률
ElseIf i = premperiod Then '납입후1년차
해지율(무해지, ipno_n, i) = Sheets("해지율").Cells(5, 21)
w_rate(무해지, ipno_n, i) = 0.5 '해지환급금 지급률
ElseIf i >= premperiod + 1 Then '납입후1년초과
해지율(무해지, ipno_n, i) = Sheets("해지율").Cells(5, 22)
w_rate(무해지, ipno_n, i) = 0.5 '해지환급금 지급률
Else '그외납입중
해지율(무해지, ipno_n, i) = Sheets("해지율").Cells(3 + i, 2)
w_rate(무해지, ipno_n, i) = 0 '해지환급금 지급률
End If
ElseIf 무해지 = 0 Then
해지율(무해지, ipno_n, i) = 0
w_rate(무해지, ipno_n, i) = 0
End If
Next i
Next 무해지
Next ipno_n
End Sub
Sub 조정계수인식()
For youl = youl_s To youl_e
For sex = sex_s To sex_e
For x = 0 To 110 '가입나이+경과기간
계수(youl, n, sex, x) = IIf((jong = 1 Or jong = 2) And si <> 0, Sheets("SI").Cells(x + 6, Application.Max((youl - 1), 0) * 2 + sex + si), 1) '(kk갯수, i경과기간, sex성별, lev급수, age(나이), 키값(n) )
Next x
Next sex
Next youl
End Sub
Sub 위험률인식()
For sex = sex_s To sex_e
For lev = 1 To lev_e
For kk = 1 To n_rate_k
For x = 0 To 110

Select Case n_rate
Case 11
qx산출(n, kk, sex, lev, x) = Sheets("적용위험률").Cells(x + 6, (sex - 1) * 5 + n_rate_c(1) + 2 * (lev - 1) + kk) '(kk진단후 경과기간, i경과기간, sex성별, lev급수, age(나이), 키값(n) )
Case Else
qx산출(n, kk, sex, lev, x) = Sheets("적용위험률").Cells(x + 6, sex + n_rate_c(kk) + 2 * (lev - 1))
End Select

Next x
Next kk
Next lev
Next sex

End Sub