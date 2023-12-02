Attribute VB_Name = "Module1"
Sub AoC_1_1()

Dim i, j As Integer
Dim left, right As String
Dim total As Double


i = 1
total = 0

Do Until Cells(i, 1).Value = ""
'loop through left of string to find number
    For j = 1 To Len(Cells(i, 1).Value)
        If IsNumeric(Mid(Cells(i, 1).Value, j, 1)) Then
            left = Mid(Cells(i, 1).Value, j, 1)
            Exit For
        End If
    Next
'loop through right of string to find number
    For j = Len(Cells(i, 1).Value) To 1 Step -1
        If IsNumeric(Mid(Cells(i, 1).Value, j, 1)) Then
            right = Mid(Cells(i, 1).Value, j, 1)
            Exit For
        End If
    Next
    total = total + CInt(left & right)
    Cells(i, 2).Value = CInt(left & right)
    i = i + 1
Loop

Cells(1, 3).Value = total
End Sub

Sub AoC_1_2()

Dim i, j As Integer
Dim strNumLeft, strNumRight As String
Dim intTextLeft, intTextRight As Integer
Dim posNumLeft, posNumRight As Integer
Dim totalLeft, totalRight As String
Dim total As Double
Dim arr(0 To 9) As Integer
Dim leftmost, rightmost As Integer


i = 1
total = 0

Do Until Cells(i, 1).Value = ""
'loop through left of string to find a number
    For j = 1 To Len(Cells(i, 1).Value)
        If IsNumeric(Mid(Cells(i, 1).Value, j, 1)) Then
            strNumLeft = Mid(Cells(i, 1).Value, j, 1) 'value of the numerical number
            posNumLeft = j 'position of the numerical number
            Exit For
        End If
    Next
    'check for occurances of zero to nine
    arr(0) = InStr(1, Cells(i, 1).Value, "zero")
    arr(1) = InStr(1, Cells(i, 1).Value, "one")
    arr(2) = InStr(1, Cells(i, 1).Value, "two")
    arr(3) = InStr(1, Cells(i, 1).Value, "three")
    arr(4) = InStr(1, Cells(i, 1).Value, "four")
    arr(5) = InStr(1, Cells(i, 1).Value, "five")
    arr(6) = InStr(1, Cells(i, 1).Value, "six")
    arr(7) = InStr(1, Cells(i, 1).Value, "seven")
    arr(8) = InStr(1, Cells(i, 1).Value, "eight")
    arr(9) = InStr(1, Cells(i, 1).Value, "nine")
    
    'compare array entries to each other
    leftmost = posNumLeft
    For j = 0 To 9
        If arr(j) <> 0 And (arr(j) < leftmost Or leftmost = 0) Then
            leftmost = arr(j)
            intTextLeft = j
        End If
    Next
    'if the leftmost position is the numerical number, return that
    If leftmost = posNumLeft Then
        Cells(i, 2).Value = CInt(strNumLeft)
        totalLeft = strNumLeft
    Else 'otherwise  return the value of text number
        Cells(i, 2).Value = intTextLeft
        totalLeft = CStr(intTextLeft)
    End If
    
    
'loop through right of string to find a number
    For j = Len(Cells(i, 1).Value) To 1 Step -1
        If IsNumeric(Mid(Cells(i, 1).Value, j, 1)) Then
            strNumRight = Mid(Cells(i, 1).Value, j, 1) 'value of the numerical number
            posNumRight = j 'position of the numerical number
            Exit For
        End If
    Next
    'check for occurances of zero to nine
    arr(0) = InStrRev(Cells(i, 1).Value, "zero")
    arr(1) = InStrRev(Cells(i, 1).Value, "one")
    arr(2) = InStrRev(Cells(i, 1).Value, "two")
    arr(3) = InStrRev(Cells(i, 1).Value, "three")
    arr(4) = InStrRev(Cells(i, 1).Value, "four")
    arr(5) = InStrRev(Cells(i, 1).Value, "five")
    arr(6) = InStrRev(Cells(i, 1).Value, "six")
    arr(7) = InStrRev(Cells(i, 1).Value, "seven")
    arr(8) = InStrRev(Cells(i, 1).Value, "eight")
    arr(9) = InStrRev(Cells(i, 1).Value, "nine")
    
    'compare array entries to each other
    rightmost = posNumRight
    For j = 0 To 9
        If arr(j) > rightmost Then
            rightmost = arr(j)
            intTextRight = j
        End If
    Next
    'if the leftmost position is the numerical number, return that
    If rightmost = posNumRight Then
        Cells(i, 3).Value = CInt(strNumRight)
        totalRight = strNumRight
    Else 'otherwise  return the value of text number
        Cells(i, 3).Value = intTextRight
        totalRight = CStr(intTextRight)
    End If
        
    total = total + CInt(totalLeft & totalRight)
    Cells(i, 4).Value = CInt(totalLeft & totalRight)
    i = i + 1
Loop

Cells(1, 5).Value = total
End Sub

