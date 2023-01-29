Attribute VB_Name = "mod_utilities"
Option Explicit

'This is set of utlities function that were created to render any given variant onto a sheet given the variant and range
'credit is given to Mike Jang GitHub: Mjang714

Sub RenderVariant(data As Variant, input_range As Range)
    
    'this is for the range that will hold all the data of the Variant
    Dim destination As Range
    
    Dim Rows As Long
    Dim Cols As Long
    
    'Since variants and arrays start off at 0 we need to add 1 to both columns and rows
    Rows = UBound(data, 1) - LBound(data, 1) + 1
    Cols = UBound(data, 2) - LBound(data, 2) + 1
    'create a range with proper dimensons
    Set destination = input_range.Resize(Rows, Cols)
    
    destination.Value = data

End Sub


