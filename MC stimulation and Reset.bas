Attribute VB_Name = "Module1"
Option Explicit
Sub MCSimulation()
Dim n As Long, Randsp As Double, Ws As Worksheet, Row As Long, Cellsp As Range, RandspArray() As String, WS2 As Worksheet, Osp As String, RandNu As Double, ONu As String, Randvc As Double, Ovc As String, Randfc As Double, Ofc As String, Cellnu As Range, Cellvc As Range, Cellfc As Range, WS3 As Worksheet, Profit As Double
ReDim RandspArray(1 To 10000), RandnuArray(1 To 10000), RandvcArray(1 To 10000), RandfcArray(1 To 10000)
Set Ws = ThisWorkbook.Sheets("Rand value")
Set WS2 = ThisWorkbook.Sheets("Question")
Set WS3 = ThisWorkbook.Sheets("iteration")
Call Reset
For n = 1 To 10000
        Row = n + 2
        Set Cellsp = Ws.Cells(Row, 2)
        Randsp = Cellsp.Value
        RandspArray(n) = Format(Randsp, "0.0000")
        If Randsp <= WS2.Range("E11") Then Osp = WS2.Range("A11") Else If Randsp <= WS2.Range("E12") Then Osp = WS2.Range("A12") Else If Randsp <= WS2.Range("E13") Then Osp = WS2.Range("A13") Else If Randsp <= WS2.Range("E14") Then Osp = WS2.Range("A14") Else If Randsp <= WS2.Range("E15") Then Osp = WS2.Range("A15")

        Set Cellnu = Ws.Cells(Row, 3)
        Randsp = Cellnu.Value
        RandnuArray(n) = Format(RandNu, "0.0000")
        If RandNu <= WS2.Range("K11") Then ONu = WS2.Range("G11") Else If RandNu <= WS2.Range("K12") Then ONu = WS2.Range("G12") Else If RandNu <= WS2.Range("K13") Then ONu = WS2.Range("G13")

        Set Cellvc = Ws.Cells(Row, 4)
        Randvc = Cellvc.Value
        RandvcArray(n) = Format(Randvc, "0.0000")
        If Randvc <= WS2.Range("Q11") Then Ovc = WS2.Range("M11") Else If Randvc <= WS2.Range("Q12") Then Ovc = WS2.Range("M12") Else If Randvc <= WS2.Range("Q13") Then Ovc = WS2.Range("M13") Else If Randvc <= WS2.Range("Q14") Then Ovc = WS2.Range("M14")

        Set Cellfc = Ws.Cells(Row, 5)
        Randfc = Cellfc.Value
        RandfcArray(n) = Format(Randfc, "0.0000")
        If Randfc <= WS2.Range("W11") Then Ofc = WS2.Range("S11") Else If Randfc <= WS2.Range("W12") Then Ofc = WS2.Range("S12") Else If Randfc <= WS2.Range("W13") Then Ofc = WS2.Range("S13")

        WS3.Cells(n, 1).Value = "Iteration " & n
        WS3.Cells(n, 2).Value = "Selling Price"
        WS3.Cells(n, 3).Value = Osp
        WS3.Cells(n, 4).Value = "Number of Units"
        WS3.Cells(n, 5).Value = ONu
        WS3.Cells(n, 6).Value = "Variable cost"
        WS3.Cells(n, 7).Value = Ovc
        WS3.Cells(n, 8).Value = "Fixed Cost"
        WS3.Cells(n, 9).Value = Ofc
        WS3.Cells(n, 10).Value = "Profit"
        Profit = (((Osp * ONu) - (Ovc * ONu)) - Ofc)
        WS3.Cells(n, 11).Value = Profit
    Next n
        With WS3.Columns("A:A")
            .Font.Bold = True
            .Font.Size = 12
        End With
        With WS3.Range("B:B,D:D,F:F,H:H,J:J")
            .Font.Bold = True
        End With
        With WS3.Range("A:A,B:B,D:D,F:F,H:H")
            .Columns.AutoFit
        End With
End Sub

Sub Reset()
Dim Ws As Worksheet
Set Ws = ThisWorkbook.Sheets("iteration")
    Ws.Cells.Clear
End Sub
