VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ResetSpreadsheet_Click()

Dim Path As String
Dim FileName As String

Dim Reset As ButtonProperties
Set Reset = New ButtonProperties

Path = "C:\NCOA Historical Data\"
FileName = Range("A1")

Call Reset.PerformanceOpt(False, xlCalculationManual)

'Performance improvement - Do not update screen while clearing sheets

Result = MsgBox("Do you wish to save to historical data?", vbYesNo + vbQuestion, "RESET DATA")

    If Result = 6 Then
    'User hits yes
    
    ActiveWorkbook.SaveAs FileName:=Path & FileName & ".xls", FileFormat:=xlNormal
    GoTo Line1
    
    ElseIf Result = 7 Then
    'User hits no
Line1:
        Result = MsgBox("Are you sure you want to clear data?", vbYesNo + vbQuestion, "WARNING!")

        If Result = 7 Then
        'User hits no
        
        End If
        
        If Result = 6 Then
        'User hits yes
                    
            'Define Sheets/ranges to clear below
            Call Reset.ClearSheet("Admin", "C5:J34", "L5:L34", "N5:AB34")
            Call Reset.ClearSheet("Demographics", "F4:G33", "J4:J33", "F4:G33", "M4:P33")
            Call Reset.ClearSheet("Attendance", "Z3", "AA52", "D4", "D53", "D5:BA34", "D54:BA83")
            Call Reset.ClearSheet("LDR Duties", "AV2", "BF2", "S4", "S5:BP34")
            Call Reset.ClearSheet("Ht-Wt", "G4:K33")
            Call Reset.ClearSheet("APFT Data", "G3", "F5:G34", "I5:I34", "K5:K34", "M5:M34")
            Call Reset.ClearSheet("Evals", "K4:M33", "R4:T33", "Y4:AA33", "AF4:AH33", "AM4:AO33", "D37:F66", "D37:F66", "K37:M66", "R37:T66")
            Call Reset.ClearSheet("Honors", "C4", "G4", "K4", "O4", "S4", "D5", "H5", "L5", "P5", "T5", "C6:V35", "AO6:AP35")
            Call Reset.ClearSheet("1059 Data", "L7:L36", "N7:N36", "W7:W36")
                               
        End If
        
     End If
     
Call Reset.PerformanceOpt(True, xlCalculationAutomatic)

End Sub

Private Sub HideEmptyRows_Change()

Dim Prop As ButtonProperties
Set Prop = New ButtonProperties

calcState = Application.Calculation
Call Prop.PerformanceOpt(False, x)

If Cells(2, 1).Value = False Then
    Cells(2, 1).Value = Prop.HideRows("Admin", 4, 5, 34, vbNullString)
                   Call Prop.HideRows("Demographics", 2, 4, 33, vbNullString)
                   Call Prop.HideRows("Ht-Wt", 3, 4, 33, vbNullString)
                   Call Prop.HideRows("APFT Data", 1, 5, 34, vbNullString)
                   Call Prop.HideRows("Evals", 1, 4, 66, vbNullString)
                   Call Prop.UnhideRows("Evals", 1, 37, 54)
                   Call Prop.HideRows("Honors", 1, 6, 35, vbNullString)
                   Call Prop.HideRows("1059 Data", 1, 7, 36, vbNullString)
                   Call Prop.PerformanceOpt(True, calcState)
    Exit Sub

ElseIf Cells(2, 1).Value = True Then
    Cells(2, 1).Value = Prop.UnhideRows("Admin", 4, 5, 34)
                   Call Prop.UnhideRows("Demographics", 2, 4, 33)
                   Call Prop.UnhideRows("Ht-Wt", 3, 4, 33)
                   Call Prop.UnhideRows("APFT Data", 1, 5, 34)
                   Call Prop.UnhideRows("Evals", 1, 4, 66)
                   Call Prop.UnhideRows("Honors", 1, 6, 35)
                   Call Prop.UnhideRows("1059 Data", 1, 7, 36)
                   Call Prop.PerformanceOpt(True, calcState)

    Exit Sub
End If

Call Prop.PerformanceOpt(True, calcState)

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
