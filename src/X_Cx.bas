Attribute VB_Name = "X_Cx"
Sub main()
Dim V, S, T, n, IANLwin, IANLvent, roomCondi, roomCond2, L2i, L2iEng As Double
Const NUM = 5000
' define variation, C and Ctr correction
Dim variation(5), sspec(5), C(5), Ctr(5) As Double
variation(1) = 6#
variation(2) = 5#
variation(3) = 6#
variation(4) = 11#
variation(5) = 11#
C(1) = -21#
C(2) = -14#
C(3) = -8#
C(4) = -5#
C(5) = -4#
Ctr(1) = -14#
Ctr(2) = -10#
Ctr(3) = -7#
Ctr(4) = -4#
Ctr(5) = -6#

' read info
V = Cells(5, "C")
S = Cells(5, "D")
T = Cells(5, "E")
n = Cells(5, "F")

IANLwin = Cells(5, "G")
IANLvent = Cells(5, "H")

'read source spec
Dim sourceSpec As Range
Set sourceSpec = Range("C8:G8")
For x = 1 To 5
    sspec(x) = sourceSpec.Cells(1, x).Value
Next x

' room condition
roomCondi = 10# * WorksheetFunction.Log10(T) + 10# * WorksheetFunction.Log10(S / V) + 11
roomCond2 = 10# * WorksheetFunction.Log10(T) + 10# * WorksheetFunction.Log10(n / V) + 21

' Generate internal spectrum
Dim L2specsWin(NUM, 5) As Double
Dim L2specsVent(NUM, 5) As Double
Dim L2specTemp(5) As Double
For i = 1 To NUM
    L2iEng = 0#
    For j = 1 To 5
        L2i = WorksheetFunction.RandBetween(-0.5 * variation(j), 0.5 * variation(j))
        L2iEng = L2iEng + 10# ^ (L2i / 10#)
        L2specTemp(j) = L2i
    Next j
    Total = 10 * WorksheetFunction.Log10(L2iEng)
    For a = 1 To 5
        L2specsWin(i, a) = L2specTemp(a) - Total + IANLwin
        L2specsVent(i, a) = L2specTemp(a) - Total + IANLvent
    Next a
Next i

'Calculate the source minus Correction
Dim Deltai_C(5), Deltai_Ctr(5) As Double
For x = 1 To 5
    Deltai_C(x) = sspec(x) - C(x)
    Deltai_Ctr(x) = sspec(x) - Ctr(x)
Next x

' calculate required Rx+C/Ctr or Dnew+C/Ctr
Dim RwC(NUM), RwCtr(NUM), DnewC(NUM), DnewCtr(NUM) As Double
Dim L2winC, L2winCtr, L2ventC, L2ventCtr As Double
For m = 1 To NUM
    L2winC = 0#
    L2winCtr = 0#
    L2ventC = 0#
    L2ventCtr = 0#
    For n = 1 To 5
        L2winC = L2winC + 10# ^ ((L2specsWin(m, n) - Deltai_C(n)) / 10#)
        L2winCtr = L2winCtr + 10# ^ ((L2specsWin(m, n) - Deltai_Ctr(n)) / 10#)
        L2ventC = L2ventC + 10# ^ ((L2specsVent(m, n) - Deltai_C(n)) / 10#)
        L2ventCtr = L2ventCtr + 10# ^ ((L2specsVent(m, n) - Deltai_Ctr(n)) / 10#)
    Next n
    RwC(m) = roomCondi - 10 * WorksheetFunction.Log10(L2winC)
    RwCtr(m) = roomCondi - 10 * WorksheetFunction.Log10(L2winCtr)
    DnewC(m) = roomCond2 - 10 * WorksheetFunction.Log10(L2ventC)
    DnewCtr(m) = roomCond2 - 10 * WorksheetFunction.Log10(L2ventCtr)
Next m

'clear all the output
Dim blank As Range
Set blank = Sheets("output").Range("A2:H11000")
blank = ""

'write the results to seperate sheet
Sheets("output").Cells(1, 1).Value = "Rw+C"
Sheets("output").Cells(1, 2).Value = "Rw+Ctr"
Sheets("output").Cells(1, 3).Value = "Dnew+C"
Sheets("output").Cells(1, 4).Value = "Dnew+Ctr"
For m = 1 To NUM
    Sheets("output").Cells(m + 1, 1).Value = RwC(m)
    Sheets("output").Cells(m + 1, 2).Value = RwCtr(m)
    Sheets("output").Cells(m + 1, 3).Value = DnewC(m)
    Sheets("output").Cells(m + 1, 4).Value = DnewCtr(m)
Next m

Call sortOutput(NUM)
Call Scan_database(IANLwin, IANLvent, roomCondi, roomCond2)
MsgBox "Calculation completed!"
End Sub

Sub sortOutput(NUM As Long)
    ' sort the results
    Dim NUMstr, rgi, rg2, rg3, rg4 As String
    NUMstr = CStr(NUM + 1)
    rgi = "A2:A" & NUMstr
    rg2 = "B2:B" & NUMstr
    rg3 = "C2:C" & NUMstr
    rg4 = "D2:D" & NUMstr
    Dim RwCsort, RwCtrSort, DnewCsort, DnewCtrSort As Range
    Set RwCsort = Sheets("output").Range(rgi)
    Set RwCtrSort = Sheets("output").Range(rg2)
    Set DnewCsort = Sheets("output").Range(rg3)
    Set DnewCtrSort = Sheets("output").Range(rg4)
    RwCsort.Sort key1:=Sheets("output").Range("A2")
    RwCtrSort.Sort key1:=Sheets("output").Range("B2")
    DnewCsort.Sort key1:=Sheets("output").Range("C2")
    DnewCtrSort.Sort key1:=Sheets("output").Range("D2")
    For m = 1 To NUM
        Sheets("output").Cells(m + 1, 5).Value = RwCsort(m)
        Sheets("output").Cells(m + 1, 6).Value = RwCtrSort(m)
        Sheets("output").Cells(m + 1, 7).Value = DnewCsort(m)
        Sheets("output").Cells(m + 1, 8).Value = DnewCtrSort(m)
    Next m
    
    ' convert to Long to get the index
    Dim fivePerc, twfivePerc, svfivePerc As Long
    fivePecr = CLng(NUM * 0.95)
    twfivePerc = CLng(NUM * 0.75)
    svfivePerc = CLng(NUM * 0.25)
    
    ' show statistical values
    Cells(11, "C").Value = RwCsort(NUM)
    Cells(11, "D").Value = RwCsort(fivePecr)
    Cells(11, "E").Value = RwCsort(twfivePerc)
    Cells(11, "F").Value = RwCsort(svfivePerc)
    Cells(11, "G").Value = RwCsort(twfivePerc) - RwCsort(svfivePerc)
    
    Cells(12, "C").Value = RwCtrSort(NUM)
    Cells(12, "D").Value = RwCtrSort(fivePecr)
    Cells(12, "E").Value = RwCtrSort(twfivePerc)
    Cells(12, "F").Value = RwCtrSort(svfivePerc)
    Cells(12, "G").Value = RwCtrSort(twfivePerc) - RwCtrSort(svfivePerc)
    
    Cells(13, "C").Value = DnewCsort(NUM)
    Cells(13, "D").Value = DnewCsort(fivePecr)
    Cells(13, "E").Value = DnewCsort(twfivePerc)
    Cells(13, "F").Value = DnewCsort(svfivePerc)
    Cells(13, "G").Value = DnewCsort(twfivePerc) - DnewCsort(svfivePerc)
    
    Cells(14, "C").Value = DnewCtrSort(NUM)
    Cells(14, "D").Value = DnewCtrSort(fivePecr)
    Cells(14, "E").Value = DnewCtrSort(twfivePerc)
    Cells(14, "F").Value = DnewCtrSort(svfivePerc)
    Cells(14, "G").Value = DnewCtrSort(twfivePerc) - DnewCtrSort(svfivePerc)
    
End Sub
Sub Scan_database(IANLwin As Variant, IANLvent As Variant, roomCondi As Variant, roomCond2 As Variant)
    'read source spec
    Dim sspec(5) As Double
    Dim sourceSpec As Range
    Set sourceSpec = Range("C8:G8")
    For x = 1 To 5
        sspec(x) = sourceSpec.Cells(1, x).Value
    Next x
       
    '''''
    Dim lRow, lCol, vRow, vCol, w As Long
    Dim glass, vent, blank As Range
    Dim eng As Double
    Dim L2i(5) As Double
    
    'clear all the output
    Set blank = Range("A17:J1000")
    blank = ""
    
    ' read glass data
    'Find the last non-blank cell in column H(1)
    lRow = Sheets("Glass").Cells(Rows.Count, 8).End(xlUp).Row
    'Find the last non-blank cell in row 1
    lCol = Sheets("Glass").Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim rg As String
    rg = "A1:" & "J" & CStr(lRow)
    Set glass = Sheets("Glass").Range(rg)
    
    ' calculate glass and output
    w = 0
    For m = 2 To lRow
        eng = 0
        For n = 6 To 10
            L2i(n - 5) = sspec(n - 5) - glass.Cells(m, n).Value + roomCondi
            eng = eng + 10 ^ (L2i(n - 5) / 10)
        Next n
        L2 = 10# * WorksheetFunction.Log10(eng)
        If L2 <= IANLwin Then
            Cells(17 + w, 2).Value = glass.Cells(m, "B")
            Cells(17 + w, 3).Value = glass.Cells(m, "D")
            Cells(17 + w, 4).Value = glass.Cells(m, "E")
            Cells(17 + w, 5).Value = L2
            For p = 6 To 10
                Cells(17 + w, p).Value = L2i(p - 5)
            Next p
            w = w + 1
        End If
    Next m
    
    
    ' read vent data
    vRow = Sheets("Vent").Cells(Rows.Count, 8).End(xlUp).Row
    vCol = Sheets("Vent").Cells(1, Columns.Count).End(xlToLeft).Column
    Dim rgv As String
    rgv = "A1:" & "J" & CStr(vRow)
    Set vent = Sheets("Vent").Range(rgv)
    
    w = w + 1
    'calculate vent and output
    For m = 2 To vRow
        eng = 0
        For n = 6 To 10
            L2i(n - 5) = sspec(n - 5) - vent.Cells(m, n).Value + roomCond2
            eng = eng + 10 ^ (L2i(n - 5) / 10)
        Next n
        L2 = 10# * WorksheetFunction.Log10(eng)
        If L2 <= IANLvent Then
            Cells(17 + w, 2).Value = vent.Cells(m, "B")
            Cells(17 + w, 3).Value = vent.Cells(m, "D")
            Cells(17 + w, 4).Value = vent.Cells(m, "E")
            Cells(17 + w, 5).Value = L2
            For p = 6 To 10
                Cells(17 + w, p).Value = L2i(p - 5)
            Next p
            w = w + 1
        End If
    Next m
End Sub
