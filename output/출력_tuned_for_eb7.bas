Sub P출력()
      
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
      txt(11) = "계지P=" & 영업월납1원
      txt(12) = "상품P=" & 상품p
    
      Dim MAL As Long
      MAL = 12
      
              For a = 1 To MAL
              If a < MAL Then
                Print #1, Trim(txt(a)); " ; ";      ' 프린트해라 #1에 트림이라는 배열
              Else
                Print #1, Trim(txt(MAL))
              End If
              Next

End Sub
Sub V출력()

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
      txt(11) = "계지V=" & Sum_계지V
      txt(12) = "상품V=" & Sum_상품V
      txt(13) = "계지한도=" & Int(신계약비한도)
      txt(14) = "상품한도=" & 상품한도
      txt(15) = "계지순보=" & 순p
      txt(16) = "상품순보=" & 상품np
    
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
Sub 한도출력()

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
      txt(11) = "한도=" & Int(신계약비한도)
      txt(12) = "신계약비=" & 사용신계약비
    
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
Sub s출력()
Select Case jong
    Case 1
    Sheets("예사비1종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
    Case 2
    Sheets("예사비2종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
    Case 3
    Sheets("예사비3종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
    Case 4
    Sheets("예사비4종").Cells(6 + nn, 39 + 4 * 무해지 + IIf(renew = 1, renewperi, IIf(gubun = "01", (mangi / 10 - 7), ipno_n))) = s
End Select
End Sub

Sub 케이스출력()
Sheets("출력").Cells(k, 1) = jong
Sheets("출력").Cells(k, 2) = sex
Sheets("출력").Cells(k, 3) = covcode
Sheets("출력").Cells(k, 4) = "" 'insperiod
Sheets("출력").Cells(k, 5) = "" 'premperiod
Sheets("출력").Cells(k, 6) = "" 'renew
Sheets("출력").Cells(k, 7) = age
Sheets("출력").Cells(k, 8) = lev
Sheets("출력").Cells(k, 9) = "" 'youl
Sheets("출력").Cells(k, 10) = qx산출(n, 1, sex, lev, age)
Sheets("출력").Cells(k, 11) = qx산출(n, 2, sex, lev, age)
Sheets("출력").Cells(k, 12) = ""
Sheets("출력").Cells(k, 13) = ""
k = k + 1
End Sub