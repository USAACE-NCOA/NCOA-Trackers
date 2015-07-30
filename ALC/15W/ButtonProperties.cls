VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private RowHiding As Boolean

Public Property Get Hiding() As Boolean
    Hiding = RowHiding
End Property

Public Property Let Hiding(Value As Boolean)
    RowHiding = Value
End Property
    
Public Function HideRows(Sheet As String, ChkCol As Integer, BeginRow As Integer, EndRow As Integer, cValue As Variant) As Boolean
    'Option for empty rows on all sheets
    
    For RowCnt = BeginRow To EndRow
        If Worksheets(Sheet).Cells(RowCnt, ChkCol).Value = cValue Then
            Worksheets(Sheet).Cells(RowCnt, ChkCol).EntireRow.Hidden = True
        End If
    Next RowCnt
    HideRows = 1
    
End Function

Public Function UnhideRows(Sheet As String, ChkCol As Integer, BeginRow As Integer, EndRow As Integer) As Boolean
    'Option for empty rows on all sheets
    
    For RowCnt = BeginRow To EndRow
        If Worksheets(Sheet).Cells(RowCnt, ChkCol).EntireRow.Hidden = True Then
            Worksheets(Sheet).Cells(RowCnt, ChkCol).EntireRow.Hidden = False
        End If
    Next RowCnt
    UnhideRows = 0
    
End Function

Sub ClearSheet(Sheet As String, ParamArray sRange() As Variant)
    'sRange() is the field or list of fields you want to be cleared.
    'Any number of fields can be entered because sRange is a variant argument.
    'So it can be zero or keep adding lists or single fields as arguments.
                
    For Each element In sRange
       Call Clear(Sheet, element)
    Next element
    
End Sub
Sub Clear(Sheet, ByVal e As String)
'Finally the business end - input Null into specified field/fields
'vbNullString inputs 0 bytes. "" will input a value and will increase file size. Overall slowdown.

Worksheets(Sheet).Range(e) = vbNullString

End Sub

Public Sub PerformanceOpt(bValue As Boolean, calcState As Variant)

Application.ScreenUpdating = bValue
Application.DisplayStatusBar = bValue
Application.Calculation = calcState
Application.EnableEvents = bValue

End Sub

Private Sub Class_Initialize()

End Sub
