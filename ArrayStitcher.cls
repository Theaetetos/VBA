VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayStitcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum Dimension
    row = 1
    Column = 2
End Enum

Public Sub ConvertHeaders(ByRef arr As Variant, ByRef ReplacePairs As Variant)

'converts headers in Arr by matching odd-numbered elements in ReplacePairs to even
'Arr must be 2d with headers at top of columns
'ReplacePairs must be 1d

'--------------------------------------------------------------

'ADD: param checks


Dim i As Integer, col_id As Integer
For i = LBound(ReplacePairs) To UBound(ReplacePairs) - 1 Step 2
    col_id = GetArrColId(CStr(ReplacePairs(i)), arr)
    If col_id > -1 Then arr(LBound(arr, 1), col_id) = ReplacePairs(i + 1)
Next i

End Sub

Public Function Join(Columns As Boolean, ParamArray Arrays() As Variant) As Variant

'takes as input a set of equal-dimensioned arrays and stitches them together into a table
'ADD:
'   functionality to handle 2d arrays of any size
'   checks on params:
'       make sure all arrs are same size
'       more?
'   more?
'Columns determines whether or not they are treated as columns in the new table or rows

'-------------------------------------------------

Dim return_arr As Variant, v As Variant
Dim bound As Integer, i As Integer, j As Integer, _
ret_row As Integer, ret_col As Integer
bound = UBound(Arrays(1)) + 1 - LBound(Arrays(1))

If Columns = True Then
    
    ReDim return_arr(1 To UBound(Arrays) + 1, 1 To bound)
    
    ret_row = 1
    ret_col = 1
    For i = LBound(Arrays) To UBound(Arrays)
        For j = LBound(Arrays(i)) To UBound(Arrays(i))
            
            ret_row = ret_row + 1
        Next j
        ret_col = ret_col + 1
    Next i
    
Else
    
    
    
End If

End Function

Public Function AddHeaders(DataArr As Variant, HeaderArr As Variant, Optional Base As Integer = 0) As Variant

'stitches passed-in headers onto passed-in data array and returns new 2d array with headers
'headers must be 2d with values in first row (no other rows will be read)
'ADD:
'   checks:
'       arrays:
'           must be 2d
'           have same num of cols

'-------------------------------------------

'checks


'get vars for sizing return array
Dim data_ubound As Long, headers_ubound As Long
data_ubound = UBound(DataArr, 1)
headers_ubound = UBound(HeaderArr, 2)

Dim data_corr As Integer, headers_corr As Integer, _
return_row_corr As Integer, return_col_corr As Integer
Select Case Base
    Case Is = 0
        If LBound(HeaderArr, 2) = 1 Then headers_corr = 1
        If LBound(DataArr, 1) = 1 Then data_corr = 1
    Case Is = 1
        If LBound(HeaderArr, 2) = 0 Then headers_corr = -1
        If LBound(DataArr, 1) = 0 Then data_corr = -1
End Select

Dim return_arr As Variant
ReDim return_arr(Base To data_ubound - LBound(DataArr, 1) + Base + 1, Base To headers_ubound - LBound(HeaderArr, 2) + Base)

'store headers in return array
Dim i As Long
For i = Base To UBound(return_arr, 2)
    return_arr(Base, i) = HeaderArr(LBound(HeaderArr), i + headers_corr)
Next i

'store data in return array
Dim j As Long
For i = Base + 1 To UBound(return_arr, 1)
    For j = Base To UBound(return_arr, 2)
        return_arr(i, j) = DataArr(i + data_corr - 1, j + data_corr)
    Next j
Next i

AddHeaders = return_arr

Erase DataArr, HeaderArr

End Function

Public Function RemoveHeaders(arr As Variant) As Variant

'removes headers from passed-in 2d array

'-------------------------------------------



End Function

Public Function To1D(arr As Variant, Optional InputDimension As Dimension = Column) As Variant

'takes first row or column of arr and transforms it into a 1d array

'-------------------------------------------------



End Function

Public Function To2D(arr As Variant, Optional OutputDimension As Dimension = Column) As Variant

'transforms a 1d array into a 2d array with either one row or one col

'-----------------------------------------------------------------------------



End Function
