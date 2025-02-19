Sub 계산기수()

dx(1, 0) = 100000
dx1(1, 0) = 100000
dx2(1, 0) = 100000
dx납면(1, 0) = 100000

For i = 0 To insperiod
cx1(1, i) = 0
cx2(1, i) = 0
x = age + i
Select Case calc_type '계산기수유형
Case 1      '#1회한, 납면사유 같음, 납면사유에 포함 암X
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 2      '#무탈퇴, 납면이랑 상관X
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 3      '#탈퇴, 납면이랑 상관X + 제/경 진단비
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 4      '#탈퇴, 뇌졸중 포함
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 5      '#탈퇴, 급성심근경색증 포함
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 6      '#탈퇴, 암+탈퇴위험률 + 기/갑진단비
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 7      '#보장보험료납입면제대상보장
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (1 - (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 8      '#탈퇴, 암, 탈퇴위험률이 암 연관(감액및면책없음), + 통합암진단비
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.25, 0) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 9      '#탈퇴, 암, 탈퇴위험률이 암 연관(감액및면책O) + 암진단비
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 10      '#탈퇴, 암산정특례
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 11      '#무탈퇴, 암 연관(감액및면책없음)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 12      '#무탈퇴, 암 연관(감액및면책O)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 13
Case 14
Case 15      '#암직접치료입원/요양병원입원(감액및면책없음)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + 0.1 * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + 0.2 * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + 0.1 * qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + 0.1 * qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 16      '#암직접치료입원/요양병원입원(감액및면책O)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (IIf(i = 0 And renew <> 2, 0.75, 1) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + 0.1 * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + 0.2 * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + 0.1 * qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + 0.1 * qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 17      '#뇌혈관질환수술,허혈성심장질환수술(1회한)
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 18      '#7대기관
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + 0.5 * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 19      '#상해사망.질병사망
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - 해지율(무해지, ipno_n, i) + qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * 해지율(무해지, ipno_n, i) / 2) * v
z(i) = (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) / 2) / (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x))
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) * z(i)) * v * (1 - 상해(youl, 1, sex, lev, x) * z(i)) * (1 - 질병(youl, 1, sex, 1, x) * z(i)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) * z(i)) * (1 - 뇌졸중(youl, 1, sex, 1, x) * z(i)) * (1 - 급성(youl, 1, sex, 1, x) * z(i)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x) * z(i)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - 해지율(무해지, ipno_n, i) + qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * 해지율(무해지, ipno_n, i) / 2) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 20      '#특정유사암항암방사선/약물
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
dx1(1, i + 1) = dx1(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
dx2(1, i + 1) = dx2(1, i) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * (dx1(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + dx2(1, i) * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 21      '#표적항암약물허가치료비(1회한) (감액및면책O)
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x))) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 22      '#표적항암약물허가치료비(1회한) (감액및면책없음)
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + 암(youl, 1, sex, 1, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 23      '#표적항암약물허가치료비(연간1회한) (감액및면책O)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.25, 0) * (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x))) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 24      '#암직접치료통원비,직접치료상급종합병원통원비(감액및면책없음)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 25      '#암직접치료통원비,직접치료상급종합병원통원비(감액및면책O)
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (IIf(i = 0 And renew <> 2, 0.75, 1) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 26      '#항암양성자/세기조절방사선(감액및면책O)
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x))) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 27      '##항암양성자/세기조절방사선(감액및면책없음)
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * (qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + 암(youl, 1, sex, 1, x)) - IIf(i = 0 And renew <> 2, 0.25, 0) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x))) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 28      '#4대유사암수술비
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 29      '#보장보험료납입지원
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (1 - (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x))) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5 * 12 * ((1 - v ^ (insperiod - i - 1)) / (1 - v) + 0.5 * v ^ (insperiod - i - 1))

Case 30      '#특정상해성뇌출혈진단비
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * (1 - 상해성뇌출혈(youl, 1, sex, 1, x))
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 31      '#남성생식기관및유방전이암
dx(1, i + 1) = dx(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - IIf(i = 0 And renew <> 2, 0.25, 0) * (1 - (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)))) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (1 - (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x))) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 32      '#항암방사선/약물치료비(치료1회당) 감액면책O
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (IIf(i = 0 And renew <> 2, 0.75, 1) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + 0.2 * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + 0.2 * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 33      '#항암방사선/약물치료비(치료1회당) 감액면책없음
dx(1, i + 1) = dx(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + 0.2 * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + 0.2 * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5

Case 34

Case 35

Case 36      '##신암치료비(암)
dx(1, i + 1) = dx(1, i) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx(1, i) * (1 - 암(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * 암(youl, 1, sex, 1, x) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5 * (v ^ (0.5) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (1.5) * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (2.5) * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (3.5) * qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (4.5) * qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) / 암(youl, 1, sex, 1, x)

Case 37      '##신암치료비(특정유사암)
dx(1, i + 1) = dx(1, i) * (1 - 기타(youl, 1, sex, 1, x) - 갑상선(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - 기타(youl, 1, sex, 1, x) - 갑상선(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx(1, i) * (1 - 기타(youl, 1, sex, 1, x) - 갑상선(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5 * (v ^ (0.5) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (1.5) * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (2.5) * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (3.5) * qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (4.5) * qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) / (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x))

Case 38      '##신암치료비(암,특정유사암)
dx(1, i + 1) = dx(1, i) * (1 - 암(youl, 1, sex, 1, x) - 기타(youl, 1, sex, 1, x) - 갑상선(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
If 납입면제 = 1 Then
dx납면(1, i + 1) = dx납면(1, i) * (1 - 해지율(무해지, ipno_n, i)) * v * (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - 암(youl, 1, sex, 1, x) - 기타(youl, 1, sex, 1, x) - 갑상선(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)
Else
dx납면(1, i + 1) = dx(1, i) * (1 - 암(youl, 1, sex, 1, x) - 기타(youl, 1, sex, 1, x) - 갑상선(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i)) * v
End If
cx(1, i) = IIf(i = 0 And renew <> 2, AA, 1) * dx(1, i) * (암(youl, 1, sex, 1, x) + 기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x)) * (1 - 해지율(무해지, ipno_n, i) / 2) * v ^ 0.5 * (v ^ (0.5) * qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (1.5) * qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (2.5) * qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (3.5) * qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x) + v ^ (4.5) * qx산출(n, 5, sex, lev, x) * 계수(youl, n, sex, x)) / (암(youl, 1, sex, 1, x) + 기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x))

Case 39      '##신암치료비(암,발생자용)
dx(1, i + 1) = dx(1, i) * v
dx납면(1, i + 1) = dx납면(1, i) * v
cx(1, i) = dx(1, i) * qx산출(n, Application.Min(i + 1, 5), sex, lev, age) * 계수(youl, n, sex, age) / 암(youl, 1, sex, 1, age) * v ^ 0.5

Case 40      '##신암치료비(특정유사암,발생자용)
dx(1, i + 1) = dx(1, i) * v
dx납면(1, i + 1) = dx납면(1, i) * v
cx(1, i) = dx(1, i) * qx산출(n, Application.Min(i + 1, 5), sex, lev, age) * 계수(youl, n, sex, age) / (기타(youl, 1, sex, 1, age) + 갑상선(youl, 1, sex, 1, age)) * v ^ 0.5

Case 41      '##신암치료비(암,특정유사암,발생자용)
dx(1, i + 1) = dx(1, i) * v
dx납면(1, i + 1) = dx납면(1, i) * v
cx(1, i) = dx(1, i) * qx산출(n, Application.Min(i + 1, 5), sex, lev, age) * 계수(youl, n, sex, age) / (암(youl, 1, sex, 1, age) + 기타(youl, 1, sex, 1, age) + 갑상선(youl, 1, sex, 1, age)) * v ^ 0.5
End Select

If 무해지 = 1 Then
Select Case calc_type
Case 1, 3, 4, 5, 6, 8, 9, 10, 17, 21, 22, 30 '탈퇴
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) / 2) * v ^ 0.5 / face
Case 2, 11, 12, 15, 16, 18, 19, 20, 23, 24, 25, 28, 32, 33 '무탈퇴
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * v ^ 0.5 / face
Case 7
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - (1 - (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - (1 - (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)) / 2) * v ^ 0.5 / face
Case 26, 27
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - (qx산출(n, 1, sex, lev, x) + qx산출(n, 2, sex, lev, x) + qx산출(n, 3, sex, lev, x)) * 계수(youl, n, sex, x) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - (qx산출(n, 1, sex, lev, x) + qx산출(n, 2, sex, lev, x) + qx산출(n, 3, sex, lev, x)) * 계수(youl, n, sex, x) / 2) * v ^ 0.5 / face
Case 29
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - (1 - (1 - 상해(youl, 1, sex, lev, x)) * (1 - 질병(youl, 1, sex, 1, x)) * (1 - IIf(i = 0 And renew <> 2, 0.75, 1) * 암(youl, 1, sex, 1, x) - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x) - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 3, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 4, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - 뇌졸중(youl, 1, sex, 1, x)) * (1 - 급성(youl, 1, sex, 1, x)) * IIf(jong = 1 Or jong = 3, (1 - 상해성뇌출혈(youl, 1, sex, 1, x)), 1)) / 2) * v ^ 0.5 / face
Case 31
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - (1 - (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x))) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - (1 - (1 - qx산출(n, 1, sex, lev, x) * 계수(youl, n, sex, x)) * (1 - qx산출(n, 2, sex, lev, x) * 계수(youl, n, sex, x))) / 2) * v ^ 0.5 / face
Case 36
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - 암(youl, 1, sex, 1, x) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - 암(youl, 1, sex, 1, x) / 2) * v ^ 0.5 / face
Case 37
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - (기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
Case 38
cx1(1, i) = w_rate(무해지, ipno_n, i) * (환급금_표준(i, 1) + 환급금_표준(i + 1, 1)) / 2 * dx납면(1, i) * 해지율(무해지, ipno_n, i) * (1 - (암(youl, 1, sex, 1, x) + 기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
cx2(1, i) = w_rate(무해지, ipno_n, i) * (준비금_표준(i, 1) + 준비금_표준(i + 1, 1)) / 2 * (dx(1, i) - dx납면(1, i)) * 해지율(무해지, ipno_n, i) * (1 - (암(youl, 1, sex, 1, x) + 기타(youl, 1, sex, 1, x) + 갑상선(youl, 1, sex, 1, x)) / 2) * v ^ 0.5 / face
End Select
End If

Next i
계산기수점프:

End Sub
Sub 계산기수합()

j_e = insperiod

For j = 0 To j_e

mxsum = 0
nxsum = 0
nxsum_k = 0
For jj = j To j_e
mxsum = mxsum + cx(1, jj) + cx1(1, jj) + cx2(1, jj)
nxsum = nxsum + dx납면(1, jj)
nxsum_k = nxsum_k + dx(1, jj)
Next jj
mx(j) = mxsum
nx납면(j) = nxsum
nx(j) = nxsum_k
Next j

nx월납 = (nx납면(0) - nx납면(premperiod)) - 11 / 24 * (dx납면(1, 0) - dx납면(1, premperiod))
End Sub
Sub 순보험료계산()

bunja_p = mx(0) - mx(insperiod)

If premperiod = 0 Then
순연납(irate) = bunja_p / dx(1, 0)
순월납(irate) = bunja_p / dx(1, 0)
순한도(irate) = 순연납(irate)
Else
순연납(irate) = bunja_p / (nx납면(0) - nx납면(premperiod))
순월납(irate) = bunja_p / (nx월납 * 12)
순한도(irate) = bunja_p / (nx납면(0) - nx납면(handoprem))
End If

순한도1원(irate) = 순한도(irate) * face
순연납1원(irate) = Round(순연납(irate) * face)
순월납1원(irate) = Round(순월납(irate) * face)
End Sub
Sub 영업보험료계산()

If premperiod = 0 Then
영업연납 = 순연납(1) '/ (1 - alpha2 - beta - ce) '일시납
영업월납 = 순연납(1)
Else
영업연납 = (순연납(1) + alpha1(n, youl, renew, IIf(gubun = "01", (mangi / 10 - 7), 0), ipno_n) * 100000 / (nx납면(0) - nx납면(premperiod))) / (1 - beta(n, youl, renew) - ce - beta5 - (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) / (nx납면(0) - nx납면(premperiod)) - alpha2(n, youl, renew, premperiod) * 100000 / (nx납면(0) - nx납면(premperiod)))
영업월납 = (순월납(1) + alpha1(n, youl, renew, IIf(gubun = "01", (mangi / 10 - 7), 0), ipno_n) * 100000 / (12 * nx월납)) / (1 - beta(n, youl, renew) - ce - beta5 - (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) / nx월납 - alpha2(n, youl, renew, premperiod) * 100000 / nx월납)
End If
영업연납1원 = Round(영업연납 * face)
영업월납1원 = Round(영업월납 * face)

If renew = 0 And gubun = "01" Then
순p = Round((순월납(irate) + 영업월납 * (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) / nx월납) * face)
Else
순p = 순월납1원(1)
End If

End Sub
Sub 한도체크()

s사용여부 = Sheets("NSP레이아웃").Cells(3 + n, 42)

해약공제계수 = Application.Min(20, insperiod)

Select Case jong
Case 1
s = Sheets("예사비1종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
Case 2
s = Sheets("예사비2종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
Case 3
s = Sheets("예사비3종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
Case 4
s = Sheets("예사비4종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(renew = 2, 0, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))))
End Select


If premperiod = 0 Then
신계약비한도 = 0
사용신계약비 = 0

Else
If s사용여부 <> 0 Then
신계약비한도 = (순한도(1) * face * 0.05 * 해약공제계수 + s * face * 10 / 1000) '순한도1원(4)=표준연납순보험료
Else
신계약비한도 = (순한도(1) * face * (0.05 * 해약공제계수)) + 순한도(1) * face * 0.45
End If

사용신계약비 = 영업월납1원 * alpha2(n, youl, renew, premperiod) * 12 + face * alpha1(n, youl, renew, IIf(gubun = "01", (mangi / 10 - 7), 0), ipno_n)

End If
End Sub
Sub 준비금계산()

For i = 0 To insperiod
If mangi = 5 Then
준비금(i, irate) = (mx(i) - mx(insperiod)) / dx(1, i)
ElseIf i < premperiod Then
준비금(i, irate) = (mx(i) - mx(insperiod) - 순연납(irate) * (nx납면(i) - nx납면(premperiod)) + 영업연납 * (beta1 + ce1) * (nx(premperiod) - nx(insperiod)) * (nx납면(0) - nx납면(i)) / (nx납면(0) - nx납면(premperiod))) / dx(1, i)
Else
준비금(i, irate) = (mx(i) - mx(insperiod) + 영업연납 * (beta1 + ce1) * (nx(i) - nx(insperiod))) / dx(1, i)
End If
준비금1원(i, irate) = Round(준비금(i, irate) * face)
Next i
End Sub
Sub 표준준비금계산()
For i = 0 To insperiod
준비금_표준(i, irate) = Round(준비금(i, irate) * face)
준비금_표준(i, irate) = Application.Max(준비금_표준(i, irate), 0)
해지공제기간 = Application.Min(7, premperiod)
환급금_표준(i, irate) = Application.Max(준비금_표준(i, irate) - IIf(i > 해지공제기간, 0, Application.RoundDown((해지공제기간 - i) / 해지공제기간 * Application.Min(Application.RoundDown(신계약비한도, 0), 사용신계약비), 0)), 0)
Next i

End Sub