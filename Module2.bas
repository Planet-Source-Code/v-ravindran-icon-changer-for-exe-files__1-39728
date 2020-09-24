Attribute VB_Name = "Module2"
Function HexToDec(HexNum As String)
    'HexNum is the hexadecimal string to be converted to decimal
    
    Dim TempValue As Integer
    Dim Digit As String
    Dim DecValue As Long

    'Trim the '&H' characters if exist
    If UCase(Left$(HexNum, 2)) = "&H" Then HexNum = Right$(HexNum, Len(HexNum) - 2)

    'Evaluate the decimal value from the Hex string
    For x = Len(HexNum) To 1 Step -1
        Digit = Mid$(HexNum, x, 1)
        Select Case UCase(Digit)
            Case Is = "A"
                TempValue = 10
            Case Is = "B"
                TempValue = 11
            Case Is = "C"
                TempValue = 12
            Case Is = "D"
                TempValue = 13
            Case Is = "E"
                TempValue = 14
            Case Is = "F"
                TempValue = 15
            Case "0" To "9"
                TempValue = Val(Digit)
            Case Else   'If the hex character is invalid, report it
                MsgBox "Error in Hex string! (" & HexNum & ")"
                Exit Function
        End Select
        DecValue = DecValue + TempValue * 16 ^ (Len(HexNum) - x)
    Next

    HexToDec = DecValue
    
End Function


