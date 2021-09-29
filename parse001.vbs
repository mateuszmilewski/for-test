Public Function parseQtyFrom2510(strQty As String) As Double
    
    parseQtyFrom2510 = 0
    
    Dim tmpstr As String
    Dim separatorString As String
    
    If IsNumeric(strQty) Then
    
        tmpstr = Replace(strQty, ".", Application.DecimalSeparator)
        parseQtyFrom2510 = CDbl(tmpstr)
        
    ElseIf strQty Like "*.*" Then
        
        'tmpStr = Replace(strQty, ".", Application.DecimalSeparator)
        'parseQtyFrom2510 = CDbl(tmpStr)
        
        
        separatorString = Mid(ThisWorkbook.Sheets("register").Range("Q17").Value, 2, 1)
        
        tmpstr = Replace(strQty, ".", separatorString)
        parseQtyFrom2510 = CDbl(tmpstr)
        
        
        
    ElseIf strQty Like "*,*" Then
        
        'tmpStr = Replace(strQty, ".", Application.DecimalSeparator)
        'parseQtyFrom2510 = CDbl(tmpStr)
        
        
        separatorString = Mid(ThisWorkbook.Sheets("register").Range("Q17").Value, 2, 1)
        
        tmpstr = Replace(strQty, ",", separatorString)
        parseQtyFrom2510 = CDbl(tmpstr)
        
        
    End If
End Function