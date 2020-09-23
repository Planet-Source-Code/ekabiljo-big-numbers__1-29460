Attribute VB_Name = "BigNumbers"
Option Explicit

Public Const MaxLength = 20000
Dim Pom0(MaxLength) As Long
Dim LengthPom0 As Long
Dim Pom1(MaxLength) As Long
Dim LengthPom1 As Long
Dim Pom2(MaxLength) As Long
Dim LengthPom2 As Long
Dim Pom3(MaxLength) As Long
Dim LengthPom3 As Long
Dim Pom4(MaxLength) As Long
Dim LengthPom4 As Long
Dim Pom5(MaxLength) As Long
Dim LengthPom5 As Long
Dim Pom6(MaxLength) As Long
Dim LengthPom6 As Long
Dim Pom7(MaxLength) As Long
Dim LengthPom7 As Long
Dim K1(MaxLength) As Long
Dim K10(MaxLength) As Long
Dim K100(MaxLength) As Long
Dim K200(MaxLength) As Long
Dim K10000(MaxLength) As Long

Public Function CompareB(a() As Long, LengthA As Long, b() As Long, LengthB As Long) As Long
Dim i As Long
Select Case LengthA - LengthB
Case Is < 0
    CompareB = -1
Case Is = 0
    CompareB = 0
    For i = 1 To LengthA
        Select Case a(LengthA - i + 1) - b(LengthA - i + 1)
        Case Is < 0
            CompareB = -1
            Exit For
        Case Is = 0
        Case Is > 0
            CompareB = 1
            Exit For
        End Select
    Next i
Case Is > 0
    CompareB = 1
End Select

End Function



Public Function BigNumberToText(c() As Long, Length As Long) As String
Dim i As Long
Dim pom As String
If c(0) = -1 Then
    pom = "-"
Else
    pom = ""
End If
pom = pom & Format$(c(Length), "0")
For i = Length - 1 To 1 Step -1
    pom = pom & Format$(c(i), "0000")
Next i
BigNumberToText = pom
End Function



Public Sub TextToBigNumber(Tekst As String, a() As Long, LengthA As Long)
Dim i As Long
Dim Prvi As String
Dim MaxLengthPrvi As Long
Dim ostatak As Long
Prvi = Trim$(Tekst)
If IsItBigNumber(Prvi) Then
    a(0) = 0
    If Left$(Prvi, 1) = "+" Then
        Prvi = Right$(Prvi, Len(Prvi) - 1)
    End If
    If Left$(Prvi, 1) = "-" Then
        Prvi = Right$(Prvi, Len(Prvi) - 1)
        a(0) = -1
    End If
    
    MaxLengthPrvi = Len(Prvi)
    If (MaxLengthPrvi \ 4) * 4 = MaxLengthPrvi Then
        LengthA = MaxLengthPrvi \ 4
        For i = 1 To LengthA
            a(i) = Mid$(Prvi, MaxLengthPrvi - i * 4 + 1, 4)
        Next i
    Else
        LengthA = MaxLengthPrvi \ 4 + 1
        ostatak = MaxLengthPrvi Mod 4
        For i = 1 To LengthA - 1
            a(i) = Mid$(Prvi, MaxLengthPrvi - i * 4 + 1, 4)
        Next i
        a(LengthA) = Mid$(Prvi, 1, ostatak)
    End If
Else
    a(1) = 0
    LengthA = 1
End If
End Sub







Public Sub MinusB(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
Select Case CompareB(a, LengthA, b, LengthB)
Case Is < 0
    Call MinusBV(b, LengthB, a, LengthA, c, LengthC)
    c(0) = -1
Case Is = 0
    c(1) = 0
    LengthC = 1
    c(0) = 0
Case Is > 0
    Call MinusBV(a, LengthA, b, LengthB, c, LengthC)
    c(0) = 0
End Select





End Sub


Public Function IsItBigNumber(Ulaz As String) As Boolean
Dim pom As String
Dim Pom1 As String
Dim i As Long
Dim IsItBigNumber1 As Boolean
Pom1 = Ulaz
pom = Left$(Pom1, 300)
If IsNumeric(pom) Then
    If InStr(1, pom, "e", 1) > 0 Then
        IsItBigNumber1 = False
    Else
        If InStr(1, pom, ".", 1) > 0 Then
            IsItBigNumber1 = False
        Else
           If InStr(1, pom, ",", 1) > 0 Then
               IsItBigNumber1 = False
           Else
               IsItBigNumber1 = True
           End If
        End If
    End If
Else
    IsItBigNumber1 = False
End If
If IsItBigNumber1 Then
    For i = 1 To Len(Pom1) \ 300
        pom = Mid$(Pom1, 300 * i, 300)
        If IsNumeric(pom) Then
            If InStr(1, pom, "e", 1) > 0 Then
                IsItBigNumber1 = False
            Else
                If InStr(1, pom, ".", 1) > 0 Then
                    IsItBigNumber1 = False
                Else
                    If InStr(1, pom, ",", 1) > 0 Then
                        IsItBigNumber1 = False
                    Else
                        If InStr(1, pom, "-", 1) > 0 Then
                            IsItBigNumber1 = False
                        Else
                            If InStr(1, pom, "+", 1) > 0 Then
                                IsItBigNumber1 = False
                            Else
                                IsItBigNumber1 = True
                            End If
                        End If
                    End If
                End If
            End If
        Else
            IsItBigNumber1 = False
        End If
        
    Next i
End If
IsItBigNumber = IsItBigNumber1
End Function

Public Sub AddBSigned(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
If a(0) = 0 And b(0) = 0 Then
    Call AddB(a, LengthA, b, LengthB, c, LengthC)
End If
If a(0) < 0 And b(0) < 0 Then
    Call AddB(a, LengthA, b, LengthB, c, LengthC)
    c(0) = -1
End If
If a(0) = 0 And b(0) < 0 Then
    Call MinusB(a, LengthA, b, LengthB, c, LengthC)
End If
If a(0) < 0 And b(0) = 0 Then
    Call MinusB(b, LengthB, a, LengthA, c, LengthC)
End If

End Sub

Public Sub MinusBSigned(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
If a(0) = 0 And b(0) = 0 Then
    Call MinusB(a, LengthA, b, LengthB, c, LengthC)
End If
If a(0) < 0 And b(0) < 0 Then
    Call MinusB(b, LengthB, a, LengthA, c, LengthC)
End If
If a(0) = 0 And b(0) < 0 Then
    Call AddB(a, LengthA, b, LengthB, c, LengthC)
    c(0) = 0
    
End If
If a(0) < 0 And b(0) = 0 Then
    Call AddB(b, LengthB, a, LengthA, c, LengthC)
    c(0) = -1
End If

End Sub

Public Sub MultBSigned(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
Call MultB(a, LengthA, b, LengthB, c, LengthC)
If (a(0) = 0 And b(0) = 0) Or (a(0) < 0 And b(0) < 0) Then
    c(0) = 0
Else
    If LengthC = 1 And c(1) = 0 Then
        c(0) = 0
    Else
        c(0) = -1
    End If
End If
End Sub

Public Sub CopyB(a() As Long, LengthA As Long, b() As Long, LengthB As Long)
Dim i As Long
LengthB = LengthA
For i = 0 To LengthA
    b(i) = a(i)
Next i
End Sub


Public Sub DivBSigned(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long, d() As Long, LengthD As Long)
Call DivB(a, LengthA, b, LengthB, c, LengthC, d, LengthD)
If (a(0) = 0 And b(0) = 0) Or (a(0) < 0 And b(0) < 0) Then
    c(0) = 0
Else
    If LengthC = 1 And c(1) = 0 Then
        c(0) = 0
    Else
        c(0) = -1
    End If
End If

End Sub

Public Sub DivB(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long, d() As Long, LengthD As Long)
If LengthB = 1 And b(1) = 0 Then
    c(1) = 0
    LengthC = 1
    c(0) = 0
    Exit Sub
End If
If LengthB = 1 And b(1) = 1 Then
    Call CopyB(a, LengthA, c, LengthC)
    Exit Sub
End If
If LengthA = 1 And a(1) = 0 Then
    c(1) = 0
    LengthC = 1
    c(0) = 0
    Exit Sub
End If
Select Case CompareB(a, LengthA, b, LengthB)
Case Is < 0
    c(1) = 0
    LengthC = 1
    c(0) = 0
Case Is = 0
    c(1) = 1
    LengthC = 1
    c(0) = 0
Case Is > 0
    Call DivBInt(a, LengthA, b, LengthB, c, LengthC, d, LengthD)
End Select

End Sub

Public Sub DivBInt(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long, d() As Long, LengthD As Long)
Dim i As Long
Dim j As Long
Dim StrA As String
Dim StrB As String
Dim StrC As String
Dim MaxLengthStrA As Long
Dim MaxLengthStrB As Long
Dim tr As String
K10(1) = 10
StrA = BigNumberToText(a, LengthA)
If Left$(StrA, 1) = "-" Then StrA = Right$(StrA, Len(StrA) - 1)
StrB = BigNumberToText(b, LengthB)
If Left$(StrA, 1) = "-" Then StrA = Right$(StrA, Len(StrA) - 1)
MaxLengthStrA = Len(StrA)
MaxLengthStrB = Len(StrB)
j = 0
Call TextToBigNumber(Left$(StrA, MaxLengthStrB), Pom2, LengthPom2)
Do While CompareB(Pom2, LengthPom2, b, LengthB) >= 0
    j = j + 1
    Call MinusBV(Pom2, LengthPom2, b, LengthB, Pom3, LengthPom3)
    Call CopyB(Pom3, LengthPom3, Pom2, LengthPom2)
Loop
StrC = Format$(j, "0")

For i = 1 To MaxLengthStrA - MaxLengthStrB
    j = 0
    Call MultB(Pom2, LengthPom2, K10, 1, Pom1, LengthPom1)
    Call TextToBigNumber(Mid$(StrA, MaxLengthStrB + i, 1), Pom2, LengthPom2)
    tr = BigNumberToText(Pom1, LengthPom1)
    tr = BigNumberToText(Pom2, LengthPom2)
    
    Call AddB(Pom1, LengthPom1, Pom2, LengthPom2, Pom3, LengthPom3)
    Call CopyB(Pom3, LengthPom3, Pom2, LengthPom2)
    Do While CompareB(Pom2, LengthPom2, b, LengthB) >= 0
        j = j + 1
        Call MinusBV(Pom2, LengthPom2, b, LengthB, Pom3, LengthPom3)
        Call CopyB(Pom3, LengthPom3, Pom2, LengthPom2)
    Loop
    StrC = StrC & Format$(j, "0")
Next i
Call CopyB(Pom2, LengthPom2, d, LengthD)
Call TextToBigNumber(StrC, c, LengthC)


End Sub



Public Sub AddB(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
Dim prenos As Long
Dim i As Long
Dim j As Long
prenos = 0
If LengthA > LengthB Then
    LengthC = LengthA + 1
    For i = 1 To LengthB
        c(i) = a(i) + b(i) + prenos
        prenos = c(i) \ 10000
        c(i) = c(i) Mod 10000
    Next i
    i = LengthB + 1
    Do While prenos > 0 And i <= LengthA
        c(i) = a(i) + prenos
        prenos = c(i) \ 10000
        c(i) = c(i) Mod 10000
        i = i + 1
    Loop
    If i > LengthA Then
        c(i) = prenos
    Else
        For j = i To LengthA
            c(j) = a(j)
        Next j
        c(LengthA + 1) = 0
    End If
Else
    LengthC = LengthB + 1
    For i = 1 To LengthA
        c(i) = a(i) + b(i) + prenos
        prenos = c(i) \ 10000
        c(i) = c(i) Mod 10000
    Next i
    i = LengthA + 1
    Do While prenos > 0 And i <= LengthB
        c(i) = b(i) + prenos
        prenos = c(i) \ 10000
        c(i) = c(i) Mod 10000
        i = i + 1
    Loop
    If i > LengthB Then
        c(i) = prenos
    Else
        For j = i To LengthB
            c(j) = b(j)
        Next j
        c(LengthB + 1) = 0
    End If
End If
If c(LengthC) = 0 Then LengthC = LengthC - 1

End Sub

Public Sub MinusBV(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
Dim prenos As Long
Dim i As Long
Dim j As Long
prenos = 0
LengthC = LengthA
For i = 1 To LengthB
    c(i) = a(i) - b(i) - prenos
    If c(i) < 0 Then
        c(i) = c(i) + 10000
        prenos = 1
    Else
        prenos = 0
    End If
Next i
i = LengthB + 1
Do While prenos > 0 And i <= LengthA
    c(i) = a(i) - prenos
    If c(i) < 0 Then
        c(i) = c(i) + 10000
        prenos = 1
    Else
        prenos = 0
    End If
    i = i + 1
Loop
If i > LengthA Then
    c(i) = prenos
Else
    For j = i To LengthA
        c(j) = a(j)
    Next j
End If
Do Until c(LengthC) <> 0 Or LengthC = 1
    LengthC = LengthC - 1
Loop

End Sub


Public Sub PowerB(a() As Long, LengthA As Long, PowerB As Long, c() As Long, LengthC As Long)
Dim i As Long
c(1) = 1
LengthC = 1
For i = 1 To PowerB
        Call MultBSigned(a, LengthA, c, LengthC, c, LengthC)
Next i

End Sub


Public Sub MultB(a() As Long, LengthA As Long, b() As Long, LengthB As Long, c() As Long, LengthC As Long)
Dim prenos As Long
Dim i As Long
Dim j As Long
If (LengthB = 1 And b(1) = 0) Or (LengthA = 1 And a(1) = 0) Then
    c(1) = 0
    LengthC = 1
    c(0) = 0
    Exit Sub
End If
If LengthB = 1 And b(1) = 1 Then
    Call CopyB(a, LengthA, c, LengthC)
    Exit Sub
End If
If LengthA = 1 And a(1) = 1 Then
    Call CopyB(b, LengthB, c, LengthC)
    Exit Sub
End If
prenos = 0
For i = 1 To LengthA + LengthB
    Pom0(i) = 0
Next i
For i = 1 To LengthB
    For j = 1 To LengthA
        Pom0(i + j - 1) = Pom0(i + j - 1) + a(j) * b(i)
        prenos = Pom0(i + j - 1) \ 10000
        Pom0(i + j - 1) = Pom0(i + j - 1) Mod 10000
        Pom0(i + j) = Pom0(i + j) + prenos
    Next j
Next i

LengthPom0 = LengthA + LengthB
Do Until Pom0(LengthPom0) <> 0 Or LengthPom0 = 1
    LengthPom0 = LengthPom0 - 1
Loop
Call CopyB(Pom0, LengthPom0, c, LengthC)
End Sub

Public Sub Factorial(Ulaz As Long, c() As Long, LengthC As Long)
Dim i As Long
c(1) = 1
LengthC = 1
For i = 2 To Ulaz
    Pom4(1) = i
    LengthPom4 = 1
    Call MultB(c, LengthC, Pom4, LengthPom4, c, LengthC)
Next i


End Sub

Public Sub SqrtB(a() As Long, LengthA As Long, c() As Long, LengthC As Long, d() As Long, LengthD As Long)
Dim Prvi As Long
Dim i As Long
Dim j As Long
Dim tr As String
For i = 0 To MaxLength
    Pom4(i) = 0
    Pom5(i) = 0
    Pom6(i) = 0
    Pom7(i) = 0
    d(i) = 0
Next i
LengthPom4 = 1
LengthPom5 = 1
LengthPom6 = 1
LengthPom7 = 1
LengthD = 1
K100(1) = 100
K200(1) = 200
K10000(2) = 1
K1(1) = 1
If a(0) = 0 Then
    Prvi = Int(Sqr(a(LengthA)))
    c(1) = Prvi
    c(0) = 0
    LengthC = 1
    d(1) = a(LengthA) - Prvi * Prvi
    LengthD = 1
    For i = LengthA - 1 To 1 Step -1
        Call MultB(d, LengthD, K10000, 2, d, LengthD)
        d(1) = a(i)
        tr = BigNumberToText(d, LengthD)
        Call MultB(c, LengthC, K200, 1, Pom4, LengthPom4)
        tr = BigNumberToText(Pom4, LengthPom4)
        Call DivB(d, LengthD, Pom4, LengthPom4, Pom5, LengthPom5, Pom7, LengthPom7)
        tr = BigNumberToText(Pom5, LengthPom5)
        Call AddB(Pom5, LengthPom5, Pom4, LengthPom4, Pom4, LengthPom4)
        tr = BigNumberToText(Pom4, LengthPom4)
        Call MultB(Pom5, LengthPom5, Pom4, LengthPom4, Pom6, LengthPom6)
        tr = BigNumberToText(Pom6, LengthPom6)
        Do While CompareB(d, LengthD, Pom6, LengthPom6) < 0 And Pom5(1) > 0
            Call MinusB(Pom4, LengthPom4, K1, 1, Pom4, LengthPom4)
            tr = BigNumberToText(Pom4, LengthPom4)
            Call MinusB(Pom5, LengthPom5, K1, 1, Pom5, LengthPom5)
            tr = BigNumberToText(Pom5, LengthPom5)
            Call MultB(Pom5, LengthPom5, Pom4, LengthPom4, Pom6, LengthPom6)
            tr = BigNumberToText(Pom6, LengthPom6)
        Loop
        Call MinusB(d, LengthD, Pom6, LengthPom6, d, LengthD)
        tr = BigNumberToText(d, LengthD)
        Call MultB(c, LengthC, K100, 1, c, LengthC)
        tr = BigNumberToText(c, LengthC)
        Call AddB(c, LengthC, Pom5, LengthPom5, c, LengthC)
        tr = BigNumberToText(c, LengthC)
    Next i
Else
    c(1) = 0
    c(0) = 0
    LengthC = 1
End If
End Sub
