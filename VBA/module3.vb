Public Function checkcellfill(name As String, score As Long, email As String, pass As String)
    
    If name = "" Then
        checkcellfill = 0
        Exit Function
    End If
    
    If score < -4 Or score > 5 Then
        checkcellfill = 0
        Exit Function
    End If
    
    If email = "" Then
        checkcellfill = 0
        Exit Function
    End If
    
    If pass = "" Then
        checkcellfill = 0
        Exit Function
    End If
    
    checkcellfill = 1
End Function

Public Function retAddress()

    Dim counter As Long
    
    For counter = 1 To 5030
        If Range("B" & counter).Value = "Not account" Then
            retAddress = counter
            Exit For
        End If
    Next counter
End Function



Sub createAccount()
    
    ' Var declarations
    
    Dim name As String
    Dim score As Long
    Dim email As String
    Dim pass As String
    Dim numb As Long
    
    
    ' Var assignments
    name = Range("L24").Value
    score = Range("M24").Value
    email = Range("N24").Value
    pass = Range("O24").Value
    
    'Check account details
    test = checkcellfill(name, score, email, pass)
    
    ' Create account
    If test = 0 Then
        MsgBox "Account creation failled"
    End If
    
    If test = 1 Then
        numb = retAddress()
        If numb > 5029 Then
            MsgBox "Contact 919784975"
        End If
        
        If numb < 5029 Then
            MsgBox " Account creation sucessfull"
            Range("A" & numb).Value = name
            Range("D" & numb).Value = score
            Range("E" & numb).Value = email
            Range("F" & numb).Value = pass
        End If
    End If
End Sub

