Sub WriteFXData(wsNew As Worksheet, wsFX As Worksheet, rowNum As Long, colNum As Long)
    ' ... 기존 코드 (Row 3-10) ...
    
    ' Row 11-14: VLOOKUP 수식 추가
    Dim colLetter As String
    
    ' 현재 열을 문자로 변환 (2=B, 3=C, 4=D...)
    colLetter = Chr(64 + colNum)
    
    ' Row 11: 거래통화 (복잡한 IF+VLOOKUP 수식)
    wsNew.Cells(11, colNum).Formula = _
        "=IF(VLOOKUP(" & colLetter & "6,FX!$H:$M,4,0)="" KRW - 한국 원""," & _
        "VLOOKUP(" & colLetter & "6,FX!$H:$M,6,0)," & _
        "IF(VLOOKUP(" & colLetter & "6,FX!$H:$M,6,0)="" KRW - 한국 원""," & _
        "VLOOKUP(" & colLetter & "6,FX!$H:$M,4,0)," & _
        "IF(VLOOKUP(" & colLetter & "6,FX!$H:$M,4,0)="" USD - 미국 달러""," & _
        "VLOOKUP(" & colLetter & "6,FX!$H:$M,6,0)," & _
        "VLOOKUP(" & colLetter & "6,FX!$H:$M,4,0))))"
    
    ' Row 12: 거래금액
    wsNew.Cells(12, colNum).Formula = "=거래금액수식"  ' 실제 수식으로 교체
    
    ' Row 13: 거래금액_미달러환산
    wsNew.Cells(13, colNum).Formula = "=거래금액USD수식"  ' 실제 수식으로 교체
    
    ' Row 14: 만기일자
    wsNew.Cells(14, colNum).Formula = "=VLOOKUP(" & colLetter & "6,FX!$H:$J,3,0)"
    
End Sub
