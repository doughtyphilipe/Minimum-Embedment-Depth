Option Explicit
Dim Gw As Worksheet, Fk As Worksheet, Sv As Worksheet, MET As Worksheet


Private Sub cmbGewinde_Change()

End Sub

Private Sub cmdClear_Click()
   Call clearAll
End Sub

Private Sub cmdMET_Click()
Dim Rm As Double, Rmmax As Double, Rs As Double
Dim D As Double, P As Double, d2 As Double, D1 As Double, A_s As Double, s As Double, sd As Double, SFV As Double
Dim C1 As Double, C3 As Double, tauBM As Double, Mgesmin As Double
'Script by Philipe Doughty - TKE Intern from 01.09.21 until 28.02.22 -  doughtyphilipe@gmail.com

'Values that depend on the selection
D = Gw.Cells(Me.cmbGewinde.ListIndex + 2, 2)
P = Gw.Cells(Me.cmbGewinde.ListIndex + 2, 3)
d2 = Gw.Cells(Me.cmbGewinde.ListIndex + 2, 4)
D1 = Gw.Cells(Me.cmbGewinde.ListIndex + 2, 5)
A_s = Gw.Cells(Me.cmbGewinde.ListIndex + 2, 6)
s = Gw.Cells(Me.cmbGewinde.ListIndex + 2, 7)
Rm = Fk.Cells(Me.cmbFestigkeitsklasse.ListIndex + 2, 2)
SFV = Sv.Cells(Me.cmbWerkstoff.ListIndex + 2, 2)

'Calculations
Rmmax = 1.2 * Rm
sd = s / D
Rs = D * (P / 2 + (D - d2) * Tan(0.523599)) / (D1 * (P / 2 + (d2 - D1) * Tan(0.523599))) * (Rm / Rm)
    If Rs <= 0.4 Then
        Rs = 0.4
    ElseIf Rs >= 1 Then
    C3 = 0.897
    End If
    
    If sd >= 1.4 And sd <= 1.9 Then
        C1 = 3.8 * sd - sd ^ 2 - 2.61
    ElseIf sd > 1.9 Then
        C1 = 1
        C3 = 0.728 + 1.769 * Rs - 2.896 * Rs ^ 2 + 1.296 * Rs ^ 3
    End If

tauBM = SFV * Rm
Mgesmin = ((Rmmax * A_s * P) / (C1 * C3 * tauBM * (P / 2 + (D - d2) * Tan(0.523598775)) * WorksheetFunction.Pi() * D)) + 2 * P

'Output
txtMET.Text = Round(Mgesmin, 3)

'Print Values to Spreadsheet
Set MET = ThisWorkbook.Sheets("Mindesteinschraubtiefe")
    If cbErgebnisse.Value = True Then
        'Create Array
        Dim list1(1 To 14, 1 To 2) As Variant
        Dim i As Single, j As Single
        list1(1, 1) = "d [mm]"
        list1(1, 2) = D
        list1(2, 1) = "P [mm]"
        list1(2, 2) = P
        list1(3, 1) = "d2 [mm]"
        list1(3, 2) = d2
        list1(4, 1) = "d1 [mm]"
        list1(4, 2) = D1
        list1(5, 1) = "As [mm^2]"
        list1(5, 2) = A_s
        list1(6, 1) = "s [mm]"
        list1(6, 2) = s
        list1(7, 1) = "Rm [N/mm^2]"
        list1(7, 2) = Rm
        list1(8, 1) = "SFV"
        list1(8, 2) = SFV
        list1(9, 1) = "sd"
        list1(9, 2) = sd
        list1(10, 1) = "C1"
        list1(10, 2) = C1
        list1(11, 1) = "C3"
        list1(11, 2) = C3
        list1(12, 1) = "Rs"
        list1(12, 2) = Rs
        list1(13, 1) = "tBM [N/mm^2]"
        list1(13, 2) = tauBM
        list1(14, 1) = "Mgesmin [mm]"
        list1(14, 2) = Mgesmin
        
        'Populate Spreadsheet
        For i = 1 To UBound(list1, 1)
            For j = 1 To UBound(list1, 2)
            MET.Cells(i, j).Value = list1(i, j)
            If i = UBound(list1, 1) Then
                MET.Cells(i, j).Interior.Color = vbRed
                MET.Cells(i, j).Font.ColorIndex = 4
            End If
            Next j
        Next i
        
    End If

    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Activate()
Dim i As Integer, j As Integer, k As Integer

'Populate Combo Boxes with values from tables
Set Gw = ThisWorkbook.Sheets("Metrische Gewinde")
For i = 2 To Gw.Range("A" & Application.Rows.Count).End(xlUp).Row
    Me.cmbGewinde.AddItem Gw.Range("A" & i).Value
Next i

Set Fk = ThisWorkbook.Sheets("Festigkeitsklasse")
For j = 2 To Fk.Range("A" & Application.Rows.Count).End(xlUp).Row
    Me.cmbFestigkeitsklasse.AddItem Fk.Range("A" & j).Value
Next j

Set Sv = ThisWorkbook.Sheets("Werkstoff")
For k = 2 To Sv.Range("A" & Application.Rows.Count).End(xlUp).Row
    Me.cmbWerkstoff.AddItem Sv.Range("A" & k).Value
Next k


End Sub


Private Sub clearAll()
    cmbGewinde.Text = ""
    cmbFestigkeitsklasse.Text = ""
    cmbWerkstoff.Text = ""
    txtMET.Text = ""
End Sub
