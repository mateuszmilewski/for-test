Public Function getFirstDate() As Date

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        Set item = Nothing
        If dane.iteration.pItems.Count > 0 Then Set item = dane.iteration.pItems(1)
        Dim tmpDate As Date
        
        If Not item Is Nothing Then
        
            If CLng(item.getDate()) > 0 Then
            
                tmpDate = CDate(item.getDate())
                For Each item In dane.iteration.pItems
                    If CDate(item.getDate) < CDate(tmpDate) Then
                        tmpDate = item.getDate
                    End If
                Next item
                
                getFirstDate = tmpDate
            Else
                getFirstDate = Date
            End If
            
        Else
            getFirstDate = Date
        End If
        
    Else
        getFirstDate = Date
    End If
End Function


Public Function getLastDate() As Date

    If Not dane Is Nothing Then
    
        Dim item As CorailItem
        Set item = Nothing
        If dane.iteration.pItems.Count > 0 Then Set item = dane.iteration.pItems(1)
        Dim tmpDate As Date

        If Not item Is Nothing Then
            If CLng(item.getDate()) > 0 Then
                tmpDate = CDate(item.getDate())
                For Each item In dane.iteration.pItems
                    If CDate(item.getDate) > CDate(tmpDate) Then
                        tmpDate = item.getDate
                    End If
                Next item
                getLastDate = tmpDate
            Else
                getLastDate = Date
            End If
        Else
            getLastDate = Date
        End If
    Else
        getLastDate = Date
    End If
End Function