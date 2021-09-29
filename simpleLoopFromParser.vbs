Set tmp = New ConvertedData

If Not tb Is Nothing Then
    htmlString = tb.innerHTML
    If checkIfHTMLTable(htmlString) Then

        Set innerTb = tb.getElementsByTagName("table")(0)
        
        If Not innerTb Is Nothing Then
        
            For Each r In innerTb.Rows
            
                Set I = Nothing
                Set I = New CorailItem
            
                For Each c In r.Cells


                    ' ascii char 34 == " (double quote)
                    ' Dec	Hex	Description
                    ' 34	22	Quotation mark/Double quote
                    
                    If UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & Replace(UCase(Me.pDateCatcher), Chr(34), "") & "*" Then

                        s = CStr(c.innerHTML)
                        arr = Split(s, ">")
                        s = Left(arr(1), 10)
                        I.parseStringToDate s
                                                    
                    ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsPCCattcher, Chr(34), "")) & "*" Then
                        
                        s = CStr(c.innerHTML)
                        arr = Split(s, ">")
                        s = arr(1)
                        s = Replace(UCase(s), "</DIV", "")
                        s = adjustDecimalSeparator(s)
                        I.besoinsPC = CDbl(s)
                        
                    ElseIf UCase(Replace(c.innerHTML, Chr(34), "")) Like "*" & UCase(Replace(Me.pBesoinsBCCatcher, Chr(34), "")) & "*" Then
                    
                        s = CStr(c.innerHTML)
                        arr = Split(s, ">")
                        s = arr(1)
                        s = Replace(UCase(s), "</DIV", "")                            
                        s = adjustDecimalSeparator(s)
                        I.besoinsBC = CDbl(s)
                        
                    ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher1) & "*") Or _
                        (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher2) & "*") Or _
                        (UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher3) & "*") Or _
                        UCase(c.innerHTML) Like "*" & UCase(Me.pCmdCatcher4) & "*" Then
                    
                        s = CStr(c.innerHTML)
                        arr = Split(s, ">")
                        s = arr(1)
                        s = Replace(UCase(s), "</DIV", "")                            
                        s = adjustDecimalSeparator(s)
                        I.order = CDbl(s)
                        
                    ElseIf (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher1) & "*") Or _
                        (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher2) & "*") Or _
                        (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher3) & "*") Or _
                        (UCase(c.innerHTML) Like "*" & UCase(Me.pExpCatcher4) & "*") Then
                    
                        s = CStr(c.innerHTML)
                        arr = Split(s, ">")
                        s = arr(1)
                        s = Replace(UCase(s), "</DIV", "")                            
                        s = adjustDecimalSeparator(s)
                        I.ship = CDbl(s)
                        
                    End If

                    
                Next c
                
                If Not avoidHeading Then
                
                    ii.addItem I
                End If
                
                avoidHeading = False
            Next r
            
            
        End If
    End If
End If