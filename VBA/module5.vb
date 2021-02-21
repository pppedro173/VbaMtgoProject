Public Function checkcellfill(address As String)
    
    If address = "" Then
        checkcellfill = 0
        Exit Function
    End If
    
    checkcellfill = 1
End Function




Sub deleteAccount()

    ' Var declarations
    
    Dim address As String
     
    ' Var assignments
    address = Range("L28").Value
    
    'Check account details
    test = checkcellfill(address)
    
    ' Create account
    If test = 0 Then
        MsgBox "Account delete failled"
    End If
    
    If test = 1 Then
    
        MsgBox " Account delete sucessfull"
        Range(address).Value = "Account deleted"
        Range(address).Offset(0, 3).Value = 999
        Range(address).Offset(0, 2).Value = "Account deleted"
        Range(address).Offset(0, 4).Value = "Account deleted"
        Range(address).Offset(0, 5).Value = "Account deleted"
        
    End If
        
End Sub