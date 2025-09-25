Sub AddVLookupFormulaRow11()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim colLetter As String
    Dim formula As String
    Dim i As Long
    
    Set wb = ActiveWorkbook
    
    ' 3번째 시트부터 마지막까지
    For i = 3 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        
        ' 3행에서 마지막 열 찾기
        lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
        
        ' B열(2열)부터 마지막 열까지
        For col = 2 To lastCol
            ' 열 번호를 열 문자로 변환 (간단한 방법)
            colLetter = Chr(64 + col)  ' 2=B, 3=C, 4=D...
            
            ' 26열(Z) 넘어가면 다른 방식
            If col > 26 Then
                colLetter = Chr(64 + Int((col - 1) / 26)) & Chr(65 + ((col - 1) Mod 26))
            End If
            
            ' 수식 생성
            formula = "=IF(VLOOKUP(" & colLetter & "6,FX!H:M,4,0)="" KRW - 한국 원"", " & _
                     "VLOOKUP(" & colLetter & "6,FX!H:M,6,0), " & _
                     "IF(VLOOKUP(" & colLetter & "6,FX!H:M,6,0)="" KRW - 한국 원"", " & _
                     "VLOOKUP(" & colLetter & "6,FX!H:M,4,0), " & _
                     "IF(VLOOKUP(" & colLetter & "6,FX!H:M,4,0)="" USD - 미국 달러"", " & _
                     "VLOOKUP(" & colLetter & "6,FX!H:M,6,0), " & _
                     "VLOOKUP(" & colLetter & "6,FX!H:M,4,0))))"
            
            ' 11행에 수식 입력
            ws.Cells(11, col).Formula = formula
        Next col
        
        Debug.Print ws.Name & " 시트 완료"
    Next i
    
    MsgBox "완료!", vbInformation
End Sub
