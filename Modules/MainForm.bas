VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "MainForm"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

Dim tWB As Workbook
Dim a1 As Double, b1 As Double, c1 As Double
Dim alpha As Double, beta As Double
Dim costOfRoyalties As Double
Dim totDepCap As Double
Dim workCap As Double
Dim startCosts As Double
Dim salesRevenue As Double
Dim P As Double
Dim prodCosts As Double
Dim intRate As Double
Dim nPos As Integer

Dim datamin As Double, datamax As Double, datarange As Double
Dim lowbins As Integer, highbins As Integer, nbins As Double
Dim binrangeinit As Double, binrangefinal As Double
Dim bins() As Double, bincenters() As Double, j As Integer
Dim c As Integer, i As Integer, R() As Double
Dim bincounts() As Integer, ChartRange As String, nr As Integer
Dim x1 As Double, x2 As Double, x3 As Double, randNum As Double

'Input Validation------------------------------------------------------------------------------

If Not IsNumeric(MainForm.costOfLandChance1) Or Not IsNumeric(MainForm.costOfLandChance2) Or Not IsNumeric(MainForm.costOfLandChance3) _
    Or Not IsNumeric(MainForm.costOfLandCost1) Or Not IsNumeric(MainForm.costOfLandCost2) Or Not IsNumeric(MainForm.costOfLandCost3) _
    Or Not IsNumeric(MainForm.costOfRoyaltiesLow) Or Not IsNumeric(MainForm.costOfRoyaltiesMode) Or Not IsNumeric(MainForm.costOfRoyaltiesHigh) _
    Or Not IsNumeric(MainForm.totalDepCapitalAve) Or Not IsNumeric(MainForm.totalDepCapitalStDev) _
    Or Not IsNumeric(MainForm.workingCapitalMin) Or Not IsNumeric(MainForm.workingCapitalMax) _
    Or Not IsNumeric(MainForm.startupCostsAve) Or Not IsNumeric(MainForm.startupCostsStDev) _
    Or Not IsNumeric(MainForm.salesRevenueLow) Or Not IsNumeric(MainForm.salesRevenueMode) Or Not IsNumeric(MainForm.salesRevenueHigh) _
    Or Not IsNumeric(MainForm.prodCostsLow) Or Not IsNumeric(MainForm.prodCostsMode) Or Not IsNumeric(MainForm.prodCostsHigh) _
    Or Not IsNumeric(MainForm.taxChance1) Or Not IsNumeric(MainForm.taxChance2) Or Not IsNumeric(MainForm.taxRate1) Or Not IsNumeric(MainForm.taxRate2) _
    Or Not IsNumeric(MainForm.interestRateMin) Or Not IsNumeric(MainForm.interestRateMax) _
    Or Not IsNumeric(MainForm.numOfSimulations) Then
    MsgBox ("One or more values are not numbers, please try again.")
    GoTo Reset
End If

If MainForm.costOfLandChance1 < 0 Or MainForm.costOfLandChance2 < 0 Or MainForm.costOfLandChance3 < 0 Then
    MsgBox ("One or more cost of land chance inputs are negative, please enter positive values.")
    GoTo Reset
ElseIf MainForm.costOfLandChance1 > 100 Or MainForm.costOfLandChance2 > 100 Or MainForm.costOfLandChance3 > 100 Then
    MsgBox ("One or more cost of land chance inputs exceed 100%, please try again.")
    GoTo Reset
End If

If MainForm.costOfLandCost1 > 0 Or MainForm.costOfLandCost2 > 0 Or MainForm.costOfLandCost3 > 0 Then
    MsgBox ("One or more cost of land cost inputs are positive, please enter negative values.")
    GoTo Reset
End If

If MainForm.costOfRoyaltiesLow > 0 Or MainForm.costOfRoyaltiesMode > 0 Or MainForm.costOfRoyaltiesHigh > 0 Then
    MsgBox ("One or more cost of royalties cost inputs are positive, please enter negative values.")
    GoTo Reset
End If

If MainForm.totalDepCapitalStDev < 0 Then
    MsgBox ("Total depreciable capital standard deviation cannot be negative, please enter a positive value.")
    GoTo Reset
End If

If MainForm.startupCostsAve > 0 Then
    MsgBox ("Start-up cost cannot be positive, please enter a negative value.")
    GoTo Reset
ElseIf MainForm.startupCostsStDev < 0 Then
    MsgBox ("Start-up cost standard deviation cannot be negative, please enter a positive value.")
    GoTo Reset
End If
    
If MainForm.prodCostsLow > 0 Or MainForm.prodCostsMode > 0 Or MainForm.prodCostsHigh > 0 Then
    MsgBox ("One or more cost of production inputs are positive please enter negative values.")
    GoTo Reset
End If

If MainForm.taxChance1 < 0 Or MainForm.taxChance2 < 0 Or MainForm.taxRate1 < 0 Or MainForm.taxRate2 < 0 Then
    MsgBox ("One or more inputs for tax are negative, please enter positive values.")
    GoTo Reset
End If

If MainForm.interestRateMin < 0 Or MainForm.interestRateMax < 0 Then
    MsgBox ("One or more inputs for interest rate are negative, please enter positive values.")
    GoTo Reset
End If

If MainForm.numOfSimulations < 0 Then
    MsgBox ("Number of simulations to run cannot be negative, please enter a positive value.")
    GoTo Reset
End If

'--------------------------------------------------------------------------------------------------------------
    
Set tWB = ThisWorkbook

ReDim R(numOfSimulations) As Double
On Error Resume Next
Application.DisplayAlerts = False

tWB.Sheets("Main").Select

For i = 1 To numOfSimulations

'Cost of Land - Discrete Distribution-----------------------------

If Rnd < (MainForm.costOfLandChance1 / 100) Then
    x1 = MainForm.costOfLandCost1
ElseIf Rnd < (MainForm.costOfLandChance2 / 100) Then
    x1 = MainForm.costOfLandCost2
Else:
    x1 = MainForm.costOfLandCost3
End If

[Cland] = x1

'Cost of Royalties - Beta-PERT Distribution----------------------

a1 = Abs(MainForm.costOfRoyaltiesLow)
b1 = Abs(MainForm.costOfRoyaltiesMode)
c1 = Abs(MainForm.costOfRoyaltiesHigh)

alpha = (4 * b1 + c1 - 5 * a1) / (c1 - a1)
beta = (5 * c1 - a1 - 4 * b1) / (c1 - a1)

costOfRoyalties = WorksheetFunction.Beta_Inv(Rnd, alpha, beta, a1, c1)

[Croyal] = (costOfRoyalties * -1)

'Total Depreciable Capital - Normal Distribution---------------------

totDepCap = WorksheetFunction.Norm_Inv(Rnd, MainForm.totalDepCapitalAve, MainForm.totalDepCapitalStDev)

[CTDC] = totDepCap

'Working Capital - Uniform Distribution------------------------

workCap = MainForm.workingCapitalMin + (MainForm.workingCapitalMax - MainForm.workingCapitalMin) * Rnd

[WC] = workCap

'Start-Up Costs - Normal Distribution---------------------

startCosts = WorksheetFunction.Norm_Inv(Rnd, MainForm.startupCostsAve, MainForm.startupCostsStDev)

[Cstart] = startCosts

'Sales Revenue - Beta-PERT Distribution----------------------

a1 = MainForm.salesRevenueLow
b1 = MainForm.salesRevenueMode
c1 = MainForm.salesRevenueHigh

alpha = (4 * b1 + c1 - 5 * a1) / (c1 - a1)
beta = (5 * c1 - a1 - 4 * b1) / (c1 - a1)

salesRevenue = WorksheetFunction.Beta_Inv(Rnd, alpha, beta, a1, c1)

[S] = salesRevenue

'Production Costs - Triangular Distribution------------------

prodCosts = triangular_inverse(P, MainForm.prodCostsLow, MainForm.prodCostsMode, MainForm.prodCostsHigh)

Range("H3") = prodCosts

'Tax - Discrete Distribution-----------------------------

If Rnd < (MainForm.taxChance1 / 100) Then
    x1 = MainForm.taxRate1
Else:
    x1 = MainForm.taxRate2
End If

[tax] = x1

'Interest Rate - Uniform Distribution------------------------

intRate = MainForm.interestRateMin + (MainForm.interestRateMax - MainForm.interestRateMin) * Rnd

Range("H4") = intRate

R(i) = Range("N24")

If Range("N24") > 0 Then nPos = nPos + 1

Next i

'Histogram Creation

datamin = WorksheetFunction.Min(R)
datamax = WorksheetFunction.Max(R)
datarange = datamax - datamin
lowbins = Int(WorksheetFunction.Log(numOfSimulations, 2)) + 1
highbins = Int(Sqr(numOfSimulations))
nbins = (lowbins + highbins) / 2
binrangeinit = datarange / nbins
ReDim bins(1) As Double
If binrangeinit < 1 Then
    c = 1
    Do
        If 10 * binrangeinit > 1 Then
            binrangefinal = 10 * binrangeinit Mod 10
            Exit Do
        Else
            binrangeinit = 10 * binrangeinit
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal / 10 ^ c
ElseIf binrangeinit < 10 Then
    binrangefinal = binrangeinit Mod 10
Else
    c = 1
    Do
        If binrangeinit / 10 < 10 Then
            binrangefinal = binrangeinit / 10 Mod 10
            Exit Do
        Else
            binrangeinit = binrangeinit / 10
            c = c + 1
        End If
    Loop
    binrangefinal = binrangefinal * 10 ^ c
End If
i = 1
bins(1) = (datamin - ((datamin) - (binrangefinal * Fix(datamin / binrangefinal))))
Do
    i = i + 1
    ReDim Preserve bins(i) As Double
    bins(i) = bins(i - 1) + binrangefinal
Loop Until bins(i) > datamax
nbins = i
ReDim Preserve bincounts(nbins - 1) As Integer
ReDim Preserve bincenters(nbins - 1) As Double
For j = 1 To nbins - 1
    c = 0
    For i = 1 To numOfSimulations
        If R(i) > bins(j) And R(i) <= bins(j + 1) Then
            c = c + 1
        End If
    Next i
    bincounts(j) = c
    bincenters(j) = (bins(j) + bins(j + 1)) / 2
Next j
Sheets("Histogram Data").Select
Cells.Clear
Range("A1").Select
Range("A1:A" & nbins - 1) = WorksheetFunction.Transpose(bincenters)
Range("B1:B" & nbins - 1) = WorksheetFunction.Transpose(bincounts)
MainForm.Hide
Application.ScreenUpdating = False
Charts("Histogram").Delete
ActiveCell.Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    nr = Selection.Rows.Count
    ChartRange = Selection.Address
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Histogram Data'!" & ChartRange)
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.PlotArea.Select
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).Delete
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveChart.FullSeriesCollection(1).XValues = "='Histogram Data'!" & "$A$1:$A$" & nr
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Caption = "Count"
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Caption = "Bin Center"
    ActiveChart.ChartArea.Select
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Histogram"
    
    MsgBox ((nPos / numOfSimulations) * 100 & " " & "percent of simulations were positive.")

Reset:
    
End Sub

Private Sub CommandButton2_Click()

Unload MainForm

End Sub

Private Sub Label35_Click()

End Sub

Function triangular_inverse(P As Double, Low As Double, Mode As Double, High As Double) As Double

Dim a As Double, b As Double, c As Double
Dim prodCosts As Double

If Rnd < (Mode - Low) / (High - Low) Then
    a = 1
    b = -2 * Low
    c = Low ^ 2 - Rnd * (Mode - Low) * (High - Low)
    triangular_inverse = (-b + Sqr(b ^ 2 - 4 * a * c)) / 2 / a
ElseIf Rnd <= 1 Then
    a = 1
    b = -2 * High
    c = High ^ 2 - (1 - Rnd) * (High - Low) * (High - Mode)
    triangular_inverse = (-b - Sqr(b ^ 2 - 4 * a * c)) / 2 / a
Else:
End If

End Function

Private Sub Label52_Click()

End Sub

Private Sub startupCostsAve_Change()

End Sub
