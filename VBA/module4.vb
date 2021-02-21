Public Function getAdressByType(score As String)
    
    Dim counter As Long
    
    For counter = 1 To 5030
        If Range("B" & counter).Value = score And Range("J3").Value <> ("A" & counter) Then
            getAdressByType = counter
            Exit For
        End If
    Next counter

End Function

Public Function checkCellType()
    If Range("L3").Value <> "A" And Range("L3").Value <> "G" And Range("L3").Value <> "H" And Range("L3").Value <> "I" And Range("L3").Value <> "J" And Range("L3").Value <> "L" Then
        Range("J3").Value = ""
        checkCellType = 0
        Exit Function
    End If
    checkCellType = 1

End Function
Public Function brain()

    'Variable declaration

    Dim X As Integer
    Dim result As Integer
    
    'Checks if account is valid
    result = checkCellType()
    If result = 0 Then
        brain = "Error cell Type"
        Exit Function
    End If
    
    result = getAdressByType(Range("L3").Value)
    
    MsgBox result

    'For X = 7 To 11
        'Range("M" & X).Value = "Test"
    'Next X
    
    brain = "Clean exec"
    
End Function


Sub hpGodModeOn()
    Dim text As String
    text = brain()
    MsgBox text
End Sub



