Attribute VB_Name = "Ri_To_Rw"
Function RiToRwx(Ri As Variant) As Double
    Dim ref, Ccorr, Ctr, Ccorr3, Ctr3 As Variant
    ref = Array(36, 45, 52, 55, 56)
    Dim shiftedRefCurve(5) As Double
    Dim shift, deviation, temp, Rw As Double
    
    shift = 0
    deviation = calc_deviation2(Ri, ref, 5)
    If deviation = 10 Then
        Rw = Ri.Cells(1, 3).Value
    ElseIf deviation < 10 Then
        Do While deviation < 10
            shift = shift + 1
            For i = 1 To 5
                shiftedRefCurve(i) = ref(i - 1) + shift
            Next i
            deviation = calc_deviation(Ri, shiftedRefCurve, 5)
        Loop
        Rw = shiftedRefCurve(3) - 1
    Else
        Do While deviation > 10
            shift = shift - 1
            For i = 1 To 5
                shiftedRefCurve(i) = ref(i - 1) + shift
            Next i
            deviation = calc_deviation(Ri, shiftedRefCurve, 5)
        Loop
        Rw = shiftedRefCurve(3)
    End If
    RiToRwx = Rw
End Function
Function calc_deviation(Ri As Variant, shifted As Variant, NUM As Integer) As Double
    Dim deviation, temp As Double
    deviation = 0
    For i = 1 To NUM
        temp = shifted(i) - Ri(i)
        If temp <= 0 Then
            temp = 0
        End If
        deviation = deviation + temp
    Next i
    calc_deviation = deviation
End Function
Function calc_deviation2(Ri As Variant, shifted As Variant, NUM As Integer) As Double
    Dim deviation, temp As Double
    deviation = 0
    For i = 1 To NUM
        temp = shifted(i - 1) - Ri(i)
        If temp <= 0 Then
            temp = 0
        End If
        deviation = deviation + temp
    Next i
    calc_deviation2 = deviation
End Function

