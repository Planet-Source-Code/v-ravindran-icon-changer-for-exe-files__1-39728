Attribute VB_Name = "Module1"
Public Function Count(Source As String, Countee As String) As Long

   'Declare our variables

    Dim i As Long, iCount As Integer

    'Initialize our variables
    iCount = 0
    i = 1

    Do
            'If we've gone past the end of the string, exit

            If (i > Len(Source)) Then Exit Do

            'Look for the next occurrence, and store it in I.

              i = InStr(i, Source, Countee, vbTextCompare)

            'If we've found another occurrence, add one to iCount and two to
            ' our current position(I)
               If i Then
                'Increment our "found" count
                      iCount = iCount + 1
                'Increment our position
                      i = i + 1
                'This code is in here so that we don't eat up all available
                ' CPU time
                      DoEvents
               End If

    Loop While i

        sCount = iCount
    Exit Function


CountError:

         sCount = 0
        Exit Function

End Function


