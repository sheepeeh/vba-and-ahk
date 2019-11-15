Function ConcatDelim(ConcatRange As Variant, Delimiter As Variant) As String

Dim Test As Boolean
Test = True

For Each i In ConcatRange
    If Test Then
        ConcatDelim = i
        Test = False
    Else
        ConcatDelim = ConcatDelim & Delimiter & i
    End If
Next i

End Function