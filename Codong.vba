Sub CreateCompanySheets()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim wsSheet1 As Worksheet
    Dim wsFX As Worksheet
    Dim wsFXOpt As Worksheet
    Dim wsNew As Worksheet
    
    Dim lastRowFX As Long
    Dim lastRowFXOpt As Long
    Dim i As Long, col As Long
    Dim companyName As String
    Dim companyDict As Object
    Dim companyKey As Variant
    
    ' 고객 분류 리스트
    Dim list1() As String
    Dim list2() As String
    Dim list3() As String
    
    list1 = Split("기아자동차,현대자동차", ",")
    list2 = Split("뱅크오브아메리카", ",")
    list3 = Split("CJ 제일제당", ",")
    
    ' 현재 워크북
    Set wb = ThisWorkbook
    Set wsSheet1 = wb.Sheets("Sheet1")
    Set wsFX = wb.Sheets("FX")
    Set wsFXOpt = wb.Sheets("FXoption")
    
    ' 고유 고객명 수집 (Dictionary 사용)
    Set companyDict = CreateObject("Scripting.Dictionary")
    
    ' FX 시트에서 고객명 수집 (AE열이 비어있을 수 있음)
    lastRowFX = wsFX.Cells(wsFX.Rows.Count, "H").End(xlUp).Row
    For i = 2 To lastRowFX
        companyName = Trim(wsFX.Cells(i, "AE").Value)
        If companyName <> "" And Not companyDict.exists(companyName) Then
            companyDict.Add companyName, True
        End If
    Next i
    
    ' FXoption 시트에서 고객명 수집 (AK열)
    lastRowFXOpt = wsFXOpt.Cells(wsFXOpt.Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastRowFXOpt
        companyName = Trim(wsFXOpt.Cells(i, "AK").Value)
        If companyName <> "" And Not companyDict.exists(companyName) Then
            companyDict.Add companyName, True
        End If
    Next i
    
    ' 새 워크북 생성
    Set newWb = Workbooks.Add
    
    ' 각 회사별로 시트 생성
    For Each companyKey In companyDict.Keys
        companyName = CStr(companyKey)
        
        ' 새 시트 생성
        Set wsNew = newWb.Worksheets.Add(After:=newWb.Worksheets(newWb.Worksheets.Count))
        wsNew.Name = companyName
        
        ' Sheet1의 템플릿 복사 (A3:A14)
        wsNew.Range("A3:A14").Value = wsSheet1.Range("A3:A14").Value
        
        ' 데이터 시작 열
        col = 2  ' B열부터 시작
        
        ' FX 데이터 처리
        For i = 2 To lastRowFX
            If Trim(wsFX.Cells(i, "AE").Value) = companyName Then
                Call WriteFXData(wsNew, wsFX, i, col, list1, list2, list3)
                col = col + 1
            End If
        Next i
        
        ' FXoption 데이터 처리
        For i = 2 To lastRowFXOpt
            If Trim(wsFXOpt.Cells(i, "AK").Value) = companyName Then
                Call WriteFXOptionData(wsNew, wsFXOpt, i, col, list1, list2, list3)
                col = col + 1
            End If
        Next i
    Next companyKey
    
    ' 기본 Sheet 삭제
    Application.DisplayAlerts = False
    For i = newWb.Worksheets.Count To 1 Step -1
        If newWb.Worksheets(i).Name Like "Sheet*" Then
            newWb.Worksheets(i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
    
    ' 새 파일 저장
    Dim savePath As String
    savePath = "C:\Users\jesst\바탕 화면\매크로\download\FX_FXoption_each_company.xlsx"
    
    ' 폴더가 없으면 생성
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists("C:\Users\jesst\바탕 화면\매크로\download") Then
        fso.CreateFolder "C:\Users\jesst\바탕 화면\매크로\download"
    End If
    
    newWb.SaveAs savePath
    MsgBox "완료되었습니다!" & vbCrLf & "저장 경로: " & savePath, vbInformation
End Sub

' FX 데이터 쓰기
Sub WriteFXData(wsNew As Worksheet, wsFX As Worksheet, rowNum As Long, colNum As Long, list1, list2, list3)
    Dim customerName As String
    Dim classification As String
    Dim productType As String
    Dim tradeType As String
    Dim direction As String
    
    customerName = Trim(wsFX.Cells(rowNum, "AE").Value)
    
    ' Row 3: 컬럼명 (I열)
    wsNew.Cells(3, colNum).Value = wsFX.Cells(rowNum, "I").Value
    
    ' Row 4: 고객명 (AE열)
    wsNew.Cells(4, colNum).Value = customerName
    
    ' Row 5: 비워둠
    
    ' Row 6: 관리번호 (H열)
    wsNew.Cells(6, colNum).Value = Trim(wsFX.Cells(rowNum, "H").Value)
    
    ' Row 7: 고객분류
    classification = GetCustomerClassification(customerName, list1, list2, list3)
    wsNew.Cells(7, colNum).Value = classification
    
    ' Row 8: 상품종류
    If Trim(wsFX.Cells(rowNum, "AJ").Value) = "YES" Then
        productType = "비정형(" & Trim(wsFX.Cells(rowNum, "AK").Value) & ")"
    Else
        productType = ""
    End If
    wsNew.Cells(8, colNum).Value = productType
    
    ' Row 9: 거래구분 (F열)
    tradeType = Trim(wsFX.Cells(rowNum, "F").Value)
    If InStr(tradeType, "1 - 신규") > 0 Then
        wsNew.Cells(9, colNum).Value = "신규"
    ElseIf InStr(tradeType, "2 - 중도청산") > 0 Then
        wsNew.Cells(9, colNum).Value = "중도청산"
    ElseIf InStr(tradeType, "3 - 부분청산") > 0 Then
        wsNew.Cells(9, colNum).Value = "부분청산"
    End If
    
    ' Row 10: 거래방향
    Dim buyingCurrency As String
    Dim sellingCurrency As String
    buyingCurrency = Trim(wsFX.Cells(rowNum, "K").Value)
    sellingCurrency = Trim(wsFX.Cells(rowNum, "M").Value)
    
    If InStr(buyingCurrency, "KRW") > 0 Then
        direction = "매입"
    ElseIf InStr(sellingCurrency, "KRW") > 0 Then
        direction = "매도"
    Else
        direction = "이종통화"
    End If
    wsNew.Cells(10, colNum).Value = direction
    
    ' Row 11-14는 사용자가 vlookup으로 처리
End Sub

' FXoption 데이터 쓰기
Sub WriteFXOptionData(wsNew As Worksheet, wsFXOpt As Worksheet, rowNum As Long, colNum As Long, list1, list2, list3)
    Dim customerName As String
    Dim classification As String
    Dim productType As String
    Dim tradeType As String
    Dim direction As String
    Dim buySell As String
    
    customerName = Trim(wsFXOpt.Cells(rowNum, "AK").Value)
    
    ' Row 3: 컬럼명 (N열)
    wsNew.Cells(3, colNum).Value = wsFXOpt.Cells(rowNum, "N").Value
    
    ' Row 4: 고객명 (AK열)
    wsNew.Cells(4, colNum).Value = customerName
    
    ' Row 5: 비워둠
    
    ' Row 6: 관리번호 (L열)
    wsNew.Cells(6, colNum).Value = Trim(wsFXOpt.Cells(rowNum, "L").Value)
    
    ' Row 7: 고객분류
    classification = GetCustomerClassification(customerName, list1, list2, list3)
    wsNew.Cells(7, colNum).Value = classification
    
    ' Row 8: 상품종류
    productType = "통화옵션 - 비정형(" & Trim(wsFXOpt.Cells(rowNum, "AT").Value) & ")"
    wsNew.Cells(8, colNum).Value = productType
    
    ' Row 9: 거래구분 (K열)
    tradeType = Trim(wsFXOpt.Cells(rowNum, "K").Value)
    If InStr(tradeType, "1 - 신규") > 0 Then
        wsNew.Cells(9, colNum).Value = "신규"
    ElseIf InStr(tradeType, "2 - 중도청산") > 0 Then
        wsNew.Cells(9, colNum).Value = "중도청산"
    ElseIf InStr(tradeType, "3 - 부분청산") > 0 Then
        wsNew.Cells(9, colNum).Value = "부분청산"
    End If
    
    ' Row 10: 거래방향 (U열)
    buySell = Trim(wsFXOpt.Cells(rowNum, "U").Value)
    If InStr(buySell, "1 - 매입") > 0 Then
        direction = "매도"
    ElseIf InStr(buySell, "2 - 매도") > 0 Then
        direction = "매입"
    End If
    wsNew.Cells(10, colNum).Value = direction
    
    ' Row 11-14는 사용자가 vlookup으로 처리
End Sub

' 고객 분류 함수
Function GetCustomerClassification(customerName As String, list1, list2, list3) As String
    Dim i As Integer
    
    ' list1 확인 (일반)
    For i = LBound(list1) To UBound(list1)
        If Trim(list1(i)) = customerName Then
            GetCustomerClassification = "1. 일반"
            Exit Function
        End If
    Next i
    
    ' list2 확인 (전문)
    For i = LBound(list2) To UBound(list2)
        If Trim(list2(i)) = customerName Then
            GetCustomerClassification = "2. 전문"
            Exit Function
        End If
    Next i
    
    ' list3 확인 (기업투자자)
    For i = LBound(list3) To UBound(list3)
        If Trim(list3(i)) = customerName Then
            GetCustomerClassification = "3. 기업투자자"
            Exit Function
        End If
    Next i
    
    GetCustomerClassification = ""
End Function
