'Division
Function divNums(ByVal arg1 As Double, ByVal arg2 As Double) As Double
    divNums = WorksheetFunction.Round(arg1 / arg2, 2)
End Function

'Multiplication
Function multiplyNums(ByVal arg1 As Double, ByVal arg2 As Double)
    multiplyNums = WorksheetFunction.Round(arg1 * arg2, 2)
End Function

'Find last row
Function lastRow(ByVal ws As Worksheet, Optional ByVal colNum As Long = 1)
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row + 1
End Function

'Convert to integer or return blank string if not integer
Function toInt(ByRef arg As String)
    arg = Trim(arg)
    If Len(arg) <> 0 And IsNumeric(arg) Then
        toInt = CInt(arg)
    Else
        toInt = ""
    End If
End Function

'Convert to double or return blank string if not double
Function toDbl(ByRef arg As String)
    arg = Trim(arg)
    If Len(arg) <> 0 And IsNumeric(arg) Then
        toDbl = CDbl(arg)
    Else
        toDbl = ""
    End If
End Function

'Convert string to lower case and remove extra whitespace
Function toLower(ByRef arg As String)
    arg = Trim(LCase(arg))
    toLower = arg
End Function

'Convert Boolean value to Yes/No
Function boolToString(ByRef arg As Boolean, Optional ByRef retNo As Boolean = True) As String
    If arg = True Then
        boolToString = "Kyllä"
    Else
        If retNo = True Then
            boolToString = "Ei"
        Else
            boolToString = ""
        End If
    End If
End Function
