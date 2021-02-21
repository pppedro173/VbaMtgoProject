Public Function checkcellfill(address As String)
    
    If address = "" Then
        checkcellfill = 0
        Exit Function
    End If
    
    checkcellfill = 1
End Function


Public Function updateScore(acc1, acc2)

    Range(acc1).Offset(0, 3).Value = Range(acc1).Offset(0, 3).Value + 1
    
    Range(acc2).Offset(0, 3).Value = Range(acc2).Offset(0, 3).Value - 1
    
    If Range(acc1).Offset(0, 3).Value = 5 Then
        Range(acc1).Offset(0, 3).Value = 0
    End If
    
    If Range(acc2).Offset(0, 3).Value = -5 Then
        Range(acc2).Offset(0, 3).Value = 0
    End If
    
    updateScore = 1

End Function





Public Function update(acc1, acc2, hist1, hist2)
    
    Dim aux1 As Range
    Dim aux2 As Range
    
    If hist1 = 0 Then
        histaccnew1 = acc2 & ","
    End If
    If hist1 <> 0 Then
        histaccnew1 = hist1 & acc2 & ","
    End If
    
    If hist2 = 0 Then
        histaccnew2 = acc1 & ","
    End If
    If hist2 <> 0 Then
        histaccnew2 = hist2 & acc1 & ","
    End If
   
    Range(acc1).Offset(0, 2).Value = histaccnew1
    Range(acc2).Offset(0, 2).Value = histaccnew2
    
    update = 1
    
End Function



Sub runMatchPair1()
    
    Dim hist1 As String
    Dim acc1 As String
    Dim acc2 As String
    Dim hist2 As String
    
    acc1 = Range("J3").Value
    acc2 = Range("J7").Value
    
    cond1 = checkcellfill(acc1)
    cond2 = checkcellfill(acc2)
    
    If cond1 = 0 Or cond2 = 0 Then
        MsgBox "No account to run match"
    End If
    
    If cond1 = 1 And cond2 = 1 Then
        hist1 = Range("N3").Value
        hist2 = Range("N7").Value
        res = update(acc1, acc2, hist1, hist2)
        res = updateScore(acc1, acc2)
    End If
    
End Sub
Sub runMatchPair2()

    Dim hist1 As String
    Dim acc1 As String
    Dim acc2 As String
    Dim hist2 As String
    

    acc1 = Range("J3").Value
    acc2 = Range("J8").Value
    
    cond1 = checkcellfill(acc1)
    cond2 = checkcellfill(acc2)
    
    If cond1 = 0 Or cond2 = 0 Then
        MsgBox "No account to run match"
    End If
    
    If cond1 = 1 And cond2 = 1 Then
        hist1 = Range("N3").Value
        hist2 = Range("N8").Value
        res = update(acc1, acc2, hist1, hist2)
        res = updateScore(acc1, acc2)
    End If

End Sub
Sub runMatchPair3()

    Dim hist1 As String
    Dim acc1 As String
    Dim acc2 As String
    Dim hist2 As String
    
    acc1 = Range("J3").Value
    acc2 = Range("J9").Value
    
    cond1 = checkcellfill(acc1)
    cond2 = checkcellfill(acc2)
    
    If cond1 = 0 Or cond2 = 0 Then
        MsgBox "No account to run match"
    End If
    
    If cond1 = 1 And cond2 = 1 Then
        hist1 = Range("N3").Value
        hist2 = Range("N9").Value
        res = update(acc1, acc2, hist1, hist2)
        res = updateScore(acc1, acc2)
    End If
    
End Sub
Sub runMatchPair4()

    Dim hist1 As String
    Dim acc1 As String
    Dim acc2 As String
    Dim hist2 As String
    

    acc1 = Range("J3").Value
    acc2 = Range("J10").Value
    
    cond1 = checkcellfill(acc1)
    cond2 = checkcellfill(acc2)
    
    If cond1 = 0 Or cond2 = 0 Then
        MsgBox "No account to run match"
    End If
    
    If cond1 = 1 And cond2 = 1 Then
        hist1 = Range("N3").Value
        hist2 = Range("N10").Value
        res = update(acc1, acc2, hist1, hist2)
        res = updateScore(acc1, acc2)
    End If
End Sub
Sub runMatchPair5()

    Dim hist1 As String
    Dim acc1 As String
    Dim acc2 As String
    Dim hist2 As String
    
    acc1 = Range("J3").Value
    acc2 = Range("J11").Value
    
    cond1 = checkcellfill(acc1)
    cond2 = checkcellfill(acc2)
    
    If cond1 = 0 Or cond2 = 0 Then
        MsgBox "No account to run match"
    End If
    
    If cond1 = 1 And cond2 = 1 Then
        hist1 = Range("N3").Value
        hist2 = Range("N11").Value
        res = update(acc1, acc2, hist1, hist2)
        res = updateScore(acc1, acc2)
    End If

End Sub

