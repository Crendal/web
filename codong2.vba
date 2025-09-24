Sub SimpleFXOptionMacro()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim wsFXOpt As Worksheet
    Dim wsNew As Worksheet
    
    Dim lastRow As Long
    Dim i As Long, col As Long
    Dim companyName As String
    Dim companies As Object
    
    ' 현재 워크북과 FX_Option 시트
    Set wb = ThisWorkbook
    Set wsFXOpt = wb.Sheets("FX_Option")  ' 시트명 확인 필요 수정해도됨.
    Set companies = CreateObject("Scripting.Dictionary")
    
    ' 고객명 수집 (AK열)
    lastRow = wsFXOpt.Cells(wsFXOpt.Rows.Count, "AK").End(xlUp).Row
    For i = 2 To lastRow
        companyName = Trim(wsFXOpt.Cells(i, "AK").Value & "")
        If companyName <> "" And Not companies.exists(companyName) Then
            companies.Add companyName, True
        End If
    Next i
    
    ' 새 워크북 생성
    Set newWb = Workbooks.Add
    
    ' 회사별 시트 생성
    For Each companyName In companies.Keys
        Set wsNew = newWb.Worksheets.Add
        wsNew.Name = companyName
        
        ' 헤더 작성 (Sheet1의 A3:A14에서 가져오기)
        wsNew.Range("A3:A14").Value = wsSheet1.Range("A3:A14").Value
       
        
        ' 데이터 입력
        col = 2
        For i = 2 To lastRow
            If Trim(wsFXOpt.Cells(i, "AK").Value & "") = companyName Then
                ' Row 3: 컬럼명 (N열)
                wsNew.Cells(3, col).Value = wsFXOpt.Cells(i, "N").Value & ""
                
                ' Row 4: 고객명 (AK열)
                wsNew.Cells(4, col).Value = companyName
                
                ' Row 6: 관리번호 (L열)
                wsNew.Cells(6, col).Value = wsFXOpt.Cells(i, "L").Value & ""
                
                ' Row 7: 고객분류
                wsNew.Cells(7, col).Value = "2. 전문"
                
                ' Row 8: 상품종류 (AT열)
                wsNew.Cells(8, col).Value = "통화옵션 - 비정형(" & wsFXOpt.Cells(i, "AT").Value & "" & ")"
                
                ' Row 9: 거래구분 (K열 - 단순하게)
                wsNew.Cells(9, col).Value = wsFXOpt.Cells(i, "K").Value & ""
                
                ' Row 10: 거래방향 (U열)
                If InStr(wsFXOpt.Cells(i, "U").Value & "", "1 - 매입") > 0 Then
                    wsNew.Cells(10, col).Value = "매도"
                ElseIf InStr(wsFXOpt.Cells(i, "U").Value & "", "2 - 매도") > 0 Then
                    wsNew.Cells(10, col).Value = "매입"
                End If
                
                col = col + 1
            End If
        Next i
    Next companyName
    
    ' 기본 Sheet 삭제
    Application.DisplayAlerts = False
    newWb.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    ' 저장
    newWb.SaveAs "C:\Users\jesst\바탕 화면\매크로\download\FXOption_Simple.xlsx"
    MsgBox "완료!", vbInformation
End Sub

-------------------------------------------------------------------------------------

Sub UnifiedSimpleMacro()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim wsFX As Worksheet
    Dim wsFXOpt As Worksheet
    Dim wsNew As Worksheet
    
    Dim lastRowFX As Long, lastRowFXOpt As Long
    Dim i As Long, col As Long
    Dim companyName As String
    Dim companies As Object
    
    ' 현재 워크북과 시트들
    Set wb = ThisWorkbook
    Set wsFX = wb.Sheets("FX")
    Set wsFXOpt = wb.Sheets("FXoption")
    Set companies = CreateObject("Scripting.Dictionary")
    
    ' FX에서 고객명 수집 (AE열)
    lastRowFX = wsFX.Cells(wsFX.Rows.Count, "H").End(xlUp).Row
    For i = 2 To lastRowFX
        companyName = Trim(wsFX.Cells(i, "AE").Value & "")
        If companyName <> "" And Not companies.exists(companyName) Then
            companies.Add companyName, True
        End If
    Next i
    
    ' FXoption에서 고객명 수집 (AK열)
    lastRowFXOpt = wsFXOpt.Cells(wsFXOpt.Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastRowFXOpt
        companyName = Trim(wsFXOpt.Cells(i, "AK").Value & "")
        If companyName <> "" And Not companies.exists(companyName) Then
            companies.Add companyName, True
        End If
    Next i
    
    ' 새 워크북 생성
    Set newWb = Workbooks.Add
    
    ' 회사별 시트 생성
    For Each companyName In companies.Keys
        Set wsNew = newWb.Worksheets.Add
        wsNew.Name = companyName
        
        ' 헤더 작성
        wsNew.Cells(3, 1).Value = "컬럼명"
        wsNew.Cells(4, 1).Value = "고객명"
        wsNew.Cells(5, 1).Value = "법인등록번호"
        wsNew.Cells(6, 1).Value = "관리번호"
        wsNew.Cells(7, 1).Value = "고객분류"
        wsNew.Cells(8, 1).Value = "상품종류"
        wsNew.Cells(9, 1).Value = "거래구분"
        wsNew.Cells(10, 1).Value = "거래방향"
        wsNew.Cells(11, 1).Value = "거래통화"
        wsNew.Cells(12, 1).Value = "거래금액"
        wsNew.Cells(13, 1).Value = "거래금액_미달러환산"
        wsNew.Cells(14, 1).Value = "만기일자"
        
        col = 2
        
        ' --- FX 데이터 처리 ---
        For i = 2 To lastRowFX
            If Trim(wsFX.Cells(i, "AE").Value & "") = companyName Then
                Call WriteFXSimple(wsNew, wsFX, i, col)
                col = col + 1
            End If
        Next i
        
        ' --- FXoption 데이터 처리 ---
        For i = 2 To lastRowFXOpt
            If Trim(wsFXOpt.Cells(i, "AK").Value & "") = companyName Then
                Call WriteFXOptionSimple(wsNew, wsFXOpt, i, col)
                col = col + 1
            End If
        Next i
    Next companyName
    
    ' 기본 Sheet 삭제
    Application.DisplayAlerts = False
    If newWb.Worksheets.Count > 1 Then newWb.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    ' 저장
    newWb.SaveAs "C:\Users\jesst\바탕 화면\매크로\download\FX_FXOption_Unified.xlsx"
    MsgBox "통합 완료!", vbInformation
End Sub

' FX 데이터 간단 처리
Sub WriteFXSimple(wsNew As Worksheet, wsFX As Worksheet, rowNum As Long, colNum As Long)
    ' Row 3: 컬럼명 (I열)
    wsNew.Cells(3, colNum).Value = wsFX.Cells(rowNum, "I").Value & ""
    
    ' Row 4: 고객명 (AE열)
    wsNew.Cells(4, colNum).Value = wsFX.Cells(rowNum, "AE").Value & ""
    
    ' Row 6: 관리번호 (H열)
    wsNew.Cells(6, colNum).Value = wsFX.Cells(rowNum, "H").Value & ""
    
    ' Row 7: 고객분류 (통일)
    wsNew.Cells(7, colNum).Value = "2. 전문"
    
    ' Row 8: 상품종류 (AJ, AK열)
    If InStr(wsFX.Cells(rowNum, "AJ").Value & "", "YES") > 0 Then
        wsNew.Cells(8, colNum).Value = "비정형(" & wsFX.Cells(rowNum, "AK").Value & "" & ")"
    Else
        wsNew.Cells(8, colNum).Value = ""
    End If
    
    ' Row 9: 거래구분 (F열 - 단순하게)
    wsNew.Cells(9, colNum).Value = wsFX.Cells(rowNum, "F").Value & ""
    
    ' Row 10: 거래방향 (K, M열)
    If InStr(wsFX.Cells(rowNum, "K").Value & "", "KRW") > 0 Then
        wsNew.Cells(10, colNum).Value = "매입"
    ElseIf InStr(wsFX.Cells(rowNum, "M").Value & "", "KRW") > 0 Then
        wsNew.Cells(10, colNum).Value = "매도"
    Else
        wsNew.Cells(10, colNum).Value = "이종통화"
    End If
    
    ' --- Row 11-14: VBA로 처리 ---
    Dim buyingCurrency As String, sellingCurrency As String
    buyingCurrency = wsFX.Cells(rowNum, "K").Value & ""
    sellingCurrency = wsFX.Cells(rowNum, "M").Value & ""
    
    ' Row 11: 거래통화
    If InStr(buyingCurrency, "KRW") > 0 Then
        wsNew.Cells(11, colNum).Value = sellingCurrency
    ElseIf InStr(sellingCurrency, "KRW") > 0 Then
        wsNew.Cells(11, colNum).Value = buyingCurrency
    ElseIf InStr(buyingCurrency, "USD") > 0 Then
        wsNew.Cells(11, colNum).Value = sellingCurrency
    Else
        wsNew.Cells(11, colNum).Value = buyingCurrency
    End If
    
    ' Row 12: 거래금액
    If InStr(buyingCurrency, "USD") > 0 Then
        wsNew.Cells(12, colNum).Value = wsFX.Cells(rowNum, "L").Value & ""  ' L열
    ElseIf InStr(sellingCurrency, "USD") > 0 Then
        wsNew.Cells(12, colNum).Value = wsFX.Cells(rowNum, "N").Value & ""  ' N열
    Else
        wsNew.Cells(12, colNum).Value = wsFX.Cells(rowNum, "L").Value & ""
    End If
    
    ' Row 13: 거래금액_미달러환산 (12행과 동일)
    wsNew.Cells(13, colNum).Value = wsNew.Cells(12, colNum).Value
    
    ' Row 14: 만기일자 (J열)
    wsNew.Cells(14, colNum).Value = wsFX.Cells(rowNum, "J").Value & ""
End Sub

' FXoption 데이터 간단 처리
Sub WriteFXOptionSimple(wsNew As Worksheet, wsFXOpt As Worksheet, rowNum As Long, colNum As Long)
    ' Row 3: 컬럼명 (N열)
    wsNew.Cells(3, colNum).Value = wsFXOpt.Cells(rowNum, "N").Value & ""
    
    ' Row 4: 고객명 (AK열)
    wsNew.Cells(4, colNum).Value = wsFXOpt.Cells(rowNum, "AK").Value & ""
    
    ' Row 6: 관리번호 (L열)
    wsNew.Cells(6, colNum).Value = wsFXOpt.Cells(rowNum, "L").Value & ""
    
    ' Row 7: 고객분류 (통일)
    wsNew.Cells(7, colNum).Value = "2. 전문"
    
    ' Row 8: 상품종류 (AT열)
    wsNew.Cells(8, colNum).Value = "통화옵션 - 비정형(" & wsFXOpt.Cells(rowNum, "AT").Value & "" & ")"
    
    ' Row 9: 거래구분 (K열 - 단순하게)
    wsNew.Cells(9, colNum).Value = wsFXOpt.Cells(rowNum, "K").Value & ""
    
    ' Row 10: 거래방향 (U열)
    If InStr(wsFXOpt.Cells(rowNum, "U").Value & "", "1 - 매입") > 0 Then
        wsNew.Cells(10, colNum).Value = "매도"
    ElseIf InStr(wsFXOpt.Cells(rowNum, "U").Value & "", "2 - 매도") > 0 Then
        wsNew.Cells(10, colNum).Value = "매입"
    End If
    
    ' Row 11-14: 필요하면 FXoption용 로직 추가
End Sub
--------------------------------------------------------------------------------------

Sub FXOption_Only_Macro()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim wsFXOpt As Worksheet
    Dim wsNew As Worksheet
    
    Dim lastRow As Long
    Dim i As Long, col As Long
    Dim companyName As String
    Dim companies As Object
    
    ' 현재 워크북과 FXoption 시트
    Set wb = ThisWorkbook
    Set wsFXOpt = wb.Sheets("FXoption")  ' 또는 "FX_Option" (시트명 확인 필요)
    Set companies = CreateObject("Scripting.Dictionary")
    
    ' FXoption에서 고객명 수집 (AK열)
    lastRow = wsFXOpt.Cells(wsFXOpt.Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastRow
        companyName = Trim(wsFXOpt.Cells(i, "AK").Value & "")
        If companyName <> "" And Not companies.exists(companyName) Then
            companies.Add companyName, True
        End If
    Next i
    
    ' 새 워크북 생성 (FXoption 전용)
    Set newWb = Workbooks.Add
    
    ' 회사별 시트 생성 및 데이터 입력
    For Each companyName In companies.Keys
        ' 새 시트 생성
        Set wsNew = newWb.Worksheets.Add
        wsNew.Name = companyName
        


        ' 헤더 작성 (Sheet1의 A3:A14에서 가져오기)
        wsNew.Range("A3:A14").Value = wsSheet1.Range("A3:A14").Value
       

        
        ' 데이터 입력 시작 (B열부터)
        col = 2
        
        ' 해당 회사의 FXoption 데이터만 처리
        For i = 2 To lastRow
            If Trim(wsFXOpt.Cells(i, "AK").Value & "") = companyName Then
                ' Row 3: 컬럼명 (N열 - 신규/청산 거래일자)
                wsNew.Cells(3, col).Value = wsFXOpt.Cells(i, "N").Value & ""
                
                ' Row 4: 고객명 (AK열)
                wsNew.Cells(4, col).Value = companyName
                
                ' Row 5: 법인등록번호 (비워둠)
                
                ' Row 6: 관리번호 (L열)
                wsNew.Cells(6, col).Value = wsFXOpt.Cells(i, "L").Value & ""
                
                ' Row 7: 고객분류 (모두 "2. 전문"으로 통일)
                wsNew.Cells(7, col).Value = "2. 전문"
                
                ' Row 8: 상품종류 (AT열)
                wsNew.Cells(8, col).Value = "통화옵션 - 비정형(" & wsFXOpt.Cells(i, "AT").Value & "" & ")"
                
                ' Row 9: 거래구분 (K열 - 값 그대로)
                wsNew.Cells(9, col).Value = wsFXOpt.Cells(i, "K").Value & ""
                
                ' Row 10: 거래방향 (U열)
                If InStr(wsFXOpt.Cells(i, "U").Value & "", "1 - 매입") > 0 Then
                    wsNew.Cells(10, col).Value = "매도"
                ElseIf InStr(wsFXOpt.Cells(i, "U").Value & "", "2 - 매도") > 0 Then
                    wsNew.Cells(10, col).Value = "매입"
                Else
                    wsNew.Cells(10, col).Value = wsFXOpt.Cells(i, "U").Value & ""
                End If
                
                ' Row 11: 거래통화 (필요시 로직 추가)
                ' wsNew.Cells(11, col).Value = "추후 추가"
                
                ' Row 12: 거래금액 (필요시 컬럼 지정)
                ' wsNew.Cells(12, col).Value = wsFXOpt.Cells(i, "?").Value & ""
                
                ' Row 13: 거래금액_미달러환산 (필요시 컬럼 지정)
                ' wsNew.Cells(13, col).Value = wsFXOpt.Cells(i, "?").Value & ""
                
                ' Row 14: 만기일자 (필요시 컬럼 지정)
                ' wsNew.Cells(14, col).Value = wsFXOpt.Cells(i, "?").Value & ""
                
                col = col + 1  ' 다음 열로 이동
            End If
        Next i
    Next companyName
    
    ' 기본 Sheet1 삭제
    Application.DisplayAlerts = False
    If newWb.Worksheets.Count > 1 Then
        On Error Resume Next
        newWb.Sheets("Sheet1").Delete
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
    
    ' 파일 저장
    Dim savePath As String
    savePath = "C:\Users\jesst\바탕 화면\매크로\download\FXOption_Only.xlsx"
    
    ' 저장 및 완료 메시지
    newWb.SaveAs savePath
    MsgBox "FXOption 처리 완료!" & vbCrLf & "저장: " & savePath, vbInformation
End Sub


---------------------------------------------------------------------------
Sub FX_Only_Macro()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim wsFX As Worksheet
    Dim wsNew As Worksheet
    
    Dim lastRow As Long
    Dim i As Long, col As Long
    Dim companyName As String
    Dim companies As Object
    
    ' 현재 워크북과 FX 시트
    Set wb = ThisWorkbook
    Set wsFX = wb.Sheets("FX")
    Set companies = CreateObject("Scripting.Dictionary")
    
    ' FX 시트의 H열에서 데이터가 있는 마지막 행 번호를 찾기
    lastRow = wsFX.Cells(wsFX.Rows.Count, "H").End(xlUp).Row
    
    ' 2행부터 마지막 행까지 하나씩 확인하기
    For i = 2 To lastRow
        ' 현재 행의 AE열(고객명)에서 값을 가져와서 앞뒤 공백 제거
        companyName = Trim(wsFX.Cells(i, "AE").Value & "")
        
        ' 고객명이 비어있지 않고, Dictionary에 아직 추가 안 된 경우에만
        If companyName <> "" And Not companies.exists(companyName) Then
            ' Dictionary에 고객명 추가 (중복 방지용)
            companies.Add companyName, True
        End If
    Next i
    
    ' 새 워크북 생성 (FX 전용)
    Set newWb = Workbooks.Add
    
    ' 회사별 시트 생성 및 데이터 입력
    For Each companyName In companies.Keys
        ' 새 시트 생성
        Set wsNew = newWb.Worksheets.Add
        wsNew.Name = companyName
        

        ' 헤더 작성 (Sheet1의 A3:A14에서 가져오기)
        wsNew.Range("A3:A14").Value = wsSheet1.Range("A3:A14").Value
       

        
        ' 데이터 입력 시작 (B열부터)
        col = 2
        
        ' 해당 회사의 FX 데이터만 처리
        For i = 2 To lastRow
            If Trim(wsFX.Cells(i, "AE").Value & "") = companyName Then
                ' Row 3: 컬럼명 (I열 - 신규/청산 거래일자)
                wsNew.Cells(3, col).Value = wsFX.Cells(i, "I").Value & ""
                
                ' Row 4: 고객명 (AE열)
                wsNew.Cells(4, col).Value = companyName
                
                ' Row 5: 법인등록번호 (비워둠)
                
                ' Row 6: 관리번호 (H열)
                wsNew.Cells(6, col).Value = wsFX.Cells(i, "H").Value & ""
                
                ' Row 7: 고객분류 (모두 "2. 전문"으로 통일)
                wsNew.Cells(7, col).Value = "2. 전문"
                
                ' Row 8: 상품종류 (AJ, AK열)
                If InStr(wsFX.Cells(i, "AJ").Value & "", "YES") > 0 Then
                    wsNew.Cells(8, col).Value = "비정형(" & wsFX.Cells(i, "AK").Value & "" & ")"
                Else
                    wsNew.Cells(8, col).Value = ""
                End If
                
                ' Row 9: 거래구분 (F열 - 값 그대로)
                wsNew.Cells(9, col).Value = wsFX.Cells(i, "F").Value & ""
                
                ' Row 10: 거래방향 (K, M열)
                If InStr(wsFX.Cells(i, "K").Value & "", "KRW") > 0 Then
                    wsNew.Cells(10, col).Value = "매입"
                ElseIf InStr(wsFX.Cells(i, "M").Value & "", "KRW") > 0 Then
                    wsNew.Cells(10, col).Value = "매도"
                Else
                    wsNew.Cells(10, col).Value = "이종통화"
                End If
                
                ' --- Row 11-14: 상세 로직 처리 ---
                Dim kValue As String, mValue As String
                kValue = wsFX.Cells(i, "K").Value & ""  ' 매입통화
                mValue = wsFX.Cells(i, "M").Value & ""  ' 매도통화
                
                ' Row 11: 거래통화
                If InStr(kValue, "KRW") > 0 Then
                    ' K열에 KRW가 있으면 M열값 출력
                    wsNew.Cells(11, col).Value = mValue
                ElseIf InStr(mValue, "KRW") > 0 Then
                    ' M열에 KRW가 있으면 K열 출력
                    wsNew.Cells(11, col).Value = kValue
                ElseIf InStr(kValue, "USD") > 0 Then
                    ' K열에 USD가 있으면 M열 출력
                    wsNew.Cells(11, col).Value = mValue
                Else
                    ' 그 외에는 K열 출력
                    wsNew.Cells(11, col).Value = kValue
                End If
                
                ' Row 12: 거래금액
                If InStr(kValue, "USD - 미국 달러") > 0 Then
                    ' K열이 USD면 L열 출력
                    wsNew.Cells(12, col).Value = wsFX.Cells(i, "L").Value & ""
                ElseIf InStr(mValue, "USD - 미국 달러") > 0 Then
                    ' M열이 USD면 N열 출력
                    wsNew.Cells(12, col).Value = wsFX.Cells(i, "N").Value & ""
                Else
                    ' 기본값: L열
                    wsNew.Cells(12, col).Value = wsFX.Cells(i, "L").Value & ""
                End If
                
                ' Row 13: 거래금액_미달러환산 (12행과 같음)
                wsNew.Cells(13, col).Value = wsNew.Cells(12, col).Value
                
                ' Row 14: 만기일자 (J열)
                wsNew.Cells(14, col).Value = wsFX.Cells(i, "J").Value & ""
                
                col = col + 1  ' 다음 열로 이동
            End If
        Next i
    Next companyName
    
    ' 기본 Sheet1 삭제
    Application.DisplayAlerts = False
    If newWb.Worksheets.Count > 1 Then
        On Error Resume Next
        newWb.Sheets("Sheet1").Delete
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
    
    ' 파일 저장
    Dim savePath As String
    savePath = "C:\Users\jesst\바탕 화면\매크로\download\FX_Only.xlsx"
    
    ' 저장 및 완료 메시지
    newWb.SaveAs savePath
    MsgBox "FX 처리 완료!" & vbCrLf & "저장: " & savePath, vbInformation
End Sub
