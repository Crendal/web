' Row 8: 상품종류 (AJ, AK, G, V열 기준)
Dim akValue As String, gValue As String, vValue As String, alValue As String
akValue = Trim(wsFX.Cells(rowNum, "AK").Value & "")
gValue = Trim(wsFX.Cells(rowNum, "G").Value & "")
vValue = Trim(wsFX.Cells(rowNum, "V").Value & "")
alValue = Trim(wsFX.Cells(rowNum, "AL").Value & "")

If akValue = "YES" Then
    ' AK="YES"면 비정형
    wsNew.Cells(8, colNum).Value = "비정형(" & alValue & ")"
ElseIf akValue = "NO" Then
    ' AK="NO"일 때 세부 조건
    If gValue = " 2 - Forward" And vValue = " 1 - 명목원금교환(실물인수도 발생)" Then
        wsNew.Cells(8, colNum).Value = "선물환"
    ElseIf gValue = " 2 - Forward" And vValue = " 2 - 차액만결제" Then
        wsNew.Cells(8, colNum).Value = "NDF"
    ElseIf gValue = " 3 - F/X 스왑" Then
        wsNew.Cells(8, colNum).Value = "FX스왑"
    Else
        wsNew.Cells(8, colNum).Value = ""
    End If
Else
    wsNew.Cells(8, colNum).Value = ""
End If



'엑셀수식
=IF(VLOOKUP(B6,FX!$H:$AK,30,0)="YES",
"비정형(" & VLOOKUP(B6,FX!$H:$AL,31,0) & ")",
IF(VLOOKUP(B6,FX!$H:$AK,30,0)="NO",IF(AND(VLOOKUP(B6,FX!$H:$G,-1,0)=" 2 - Forward",VLOOKUP(B6,FX!$H:$V,15,0)=" 1 - 명목원금교환(실물인수도 발생)"),"선물환",
IF(AND(VLOOKUP(B6,FX!$H:$G,-1,0)=" 2 - Forward",VLOOKUP(B6,FX!$H:$V,15,0)=" 2 - 차액만결제"),"NDF",IF(VLOOKUP(B6,FX!$H:$G,-1,0)=" 3 - F/X 스왑","FX스왑",""))),""))
