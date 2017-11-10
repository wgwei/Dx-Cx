Attribute VB_Name = "RiToRwx"
Function RiToRwx(Ri As Range) As String
    Dim ref, Ccorr, Ctr, Ccorr3, Ctr3 As Variant
    ref = Array(36, 45, 52, 55, 56)
    Ccorr = Array(-21, -14, -8, -5, -4) 'vairant index start from 0 such as Ccorr(0)
    Ctr = Array(-14, -10, -7, -4, -6)
    
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
    Dim Ccorrv, Ctrv As Integer
    Ccorrv = calc_correction(Ri, Rw, Ccorr)
    Ctrv = calc_correction(Ri, Rw, Ctr)
    RiToRwx = CStr(Rw) & "(" & CStr(Ccorrv) & ";" & CStr(Ctrv) & ")"
End Function
Function calc_correction(Ri As Range, Rw As Double, refvalues As Variant)
    Dim Ld As Double
    Ld = 0#
    For i = 1 To 5
        Ld = Ld + 10 ^ ((refvalues(i - 1) - Ri.Cells(1, i).Value) / 10)
    Next i
    calc_correction = CInt(Round(-10 * WorksheetFunction.Log10(Ld) - Rw))
End Function
Function calc_deviation(Ri As Range, shifted As Variant, NUM As Integer) As Double
    Dim deviation, temp As Double
    deviation = 0
    For i = 1 To NUM
        temp = shifted(i) - Ri.Cells(1, i).Value
        If temp <= 0 Then
            temp = 0
        End If
        deviation = deviation + temp
    Next i
    calc_deviation = deviation
End Function
Function calc_deviation2(Ri As Range, shifted As Variant, NUM As Integer) As Double
    Dim deviation, temp As Double
    deviation = 0
    For i = 1 To NUM
        temp = shifted(i - 1) - Ri.Cells(1, i).Value
        If temp <= 0 Then
            temp = 0
        End If
        deviation = deviation + temp
    Next i
    calc_deviation2 = deviation
End Function
Function Ri3ToRwx(Ri As Range) As String
    Dim ref, Ccorr, Ctr, Ccorr3, Ctr3 As Variant
    ref = Array(33, 36, 39, 42, 45, 48, 51, 52, 53, 54, 55, 56, 56, 56, 56, 56)
    Ccorr = Array(-29, -26, -23, -21, -19, -17, -15, -13, -12, -11, -10, -9, -9, -9, -9, -9)     'vairant index start from 0 such as Ccorr(0)
    Ctr = Array(-20, -20, -18, -16, -15, -14, -13, -12, -11, -9, -8, -9, -10, -11, -13, -15)
    
    Dim shiftedRefCurve(16) As Double
    Dim shift, deviation, temp, Rw As Double
    
    shift = 0
    deviation = calc_deviation2(Ri, ref, 16)
    If deviation = 32 Then
        Rw = Ri.Cells(1, 8).Value
    ElseIf deviation < 10 Then
        Do While deviation < 32
            shift = shift + 1
            For i = 1 To 16
                shiftedRefCurve(i) = ref(i - 1) + shift
            Next i
            deviation = calc_deviation(Ri, shiftedRefCurve, 16)
        Loop
        Rw = shiftedRefCurve(8) - 1
    Else
        Do While deviation > 32
            shift = shift - 1
            For i = 1 To 16
                shiftedRefCurve(i) = ref(i - 1) + shift
            Next i
            deviation = calc_deviation(Ri, shiftedRefCurve, 16)
        Loop
        Rw = shiftedRefCurve(8)
    End If
    Dim Ccorrv, Ctrv As Integer
    Ccorrv = calc_correction(Ri, Rw, Ccorr)
    Ctrv = calc_correction(Ri, Rw, Ctr)
    Ri3ToRwx = CStr(Rw) & "(" & CStr(Ccorrv) & ";" & CStr(Ctrv) & ")"
End Function

