Attribute VB_Name = "Buttons"
Sub ClearSheet(Sheet As String, ParamArray sRange() As Variant)
    'sRange() is the field or list of fields you want to be cleared.
    'Any number of fields can be entered because sRange is a variant argument.
    'So it can be zero or keep adding lists or single fields as arguments.
        
    Worksheets(Sheet).Unprotect Password:="505401"
        
    For Each element In sRange
       Call Clear(Sheet, element)
    Next element
    
    Worksheets(Sheet).Protect Password:="505401"
    
End Sub
Sub Clear(Sheet, ByVal e As String)
'Finally the business end - input Null into specified field/fields
'vbNullString inputs 0 bytes. "" will input a value and will increase file size. Overall slowdown.

Worksheets(Sheet).Range(e) = vbNullString

End Sub

