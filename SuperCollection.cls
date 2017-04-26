VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class mirrors some of the functionality of recordsets and arrays
'while adding some flexibility and efficiency:
'   e.g., the Subset method is much faster and uses less memory than a similar operation on an array,
'    and adding and deleting columns is much easier

Option Explicit

Private cols As New Collection

Property Set ColumnCollection(Coll As Collection)

Set cols = Coll

End Property

Property Get ColumnCollection() As Collection

Set ColumnCollection = cols

End Property

Property Get FieldCount() As Long

FieldCount = cols.Count

End Property

Property Get RecordCount() As Long

If cols.Count > 0 Then
    RecordCount = cols(1).Count
Else
    RecordCount = 0
End If

End Property

Public Sub InsertColumn(Coll As Collection, Optional Index As Integer)

'adds a passed-in collection to cols at index
'Index is checked to make sure 1<=Index<=cols.Count
'   defaults to adding column after last
'Index corresponds to the index of Coll after it will have been added

'-----------------------------------------------------------

If Index < 1 Or Index > FieldCount Then
    cols.Add Coll
Else
    cols.Add Coll, Before:=Index
End If

End Sub

Public Sub AddColumn(Optional Index As Integer)

'creates a new, empty "column" at Index
'   Index is checked to make sure 1<=Index<=cols.Count+1
'   if no value supplied, defaults to adding column after last

'---------------------------------------

Dim Coll As New Collection
If Index < 1 Or Index > FieldCount Then
    cols.Add Coll
Else
    cols.Add Coll, Before:=Index
End If

End Sub

Public Sub RemoveColumn(Optional Index As Integer)

'removes field collection at Index
'   Index is checked to make sure 1<=Index<=FieldCount
'   if improper value supplied, defaults to removing last column

'---------------------------------------

If Index < 1 Or Index > FieldCount Then
    cols.Remove FieldCount
Else
    cols.Remove Index
End If

End Sub

Public Sub ReadArray(ByRef arr As Variant, Optional Tilted As Boolean)

'stores Arr columns (or rows if Tilted=True) as collections in cols
'overwrites any existing data in cols

'-------------------------------------------

'read array
Dim col_coll As New Collection
Dim Coll As Collection
Dim i As Long
Dim col_ct As Integer, j As Integer, arr_lbound As Integer, corr As Integer
If Tilted = False Then
    arr_lbound = LBound(arr, 2)
    If arr_lbound = 0 Then corr = 1
    col_ct = UBound(arr, 2) - arr_lbound + 1
    For j = 1 To col_ct
        Set Coll = New Collection
        col_coll.Add Coll
    Next j
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = arr_lbound To UBound(arr, 2)
            col_coll(j + corr).Add arr(i, j)
        Next j
    Next i
Else
    arr_lbound = LBound(arr, 1)
    If arr_lbound = 0 Then corr = 1
    col_ct = UBound(arr, 1) - arr_lbound + 1
    For j = 1 To col_ct
        Set Coll = New Collection
        col_coll.Add Coll
    Next j
    For j = arr_lbound To UBound(arr, 1)
        For i = LBound(arr, 2) To UBound(arr, 2)
            col_coll(j + corr).Add arr(j, i)
        Next i
    Next j
End If

Set cols = col_coll

End Sub

Public Sub AddHeaders(Headers As Variant)

'Inserts headers at beginning of column collections
'procedure reads number of dimensions of Headers
'then tries to find dimension with count equal to cols.Count (if Headers is 2d, columns are checked first)
'terminates if unable

'------------------------------------------------------------------

'convert Headers to 1d array if possible
Dim temp As Variant
Dim dims As Integer, read_dim As Integer, i As Integer, _
cols_ct As Integer
dims = GetArrDimensions(Headers)
cols_ct = cols.Count
If dims = 0 Then
    temp = Split(Headers, ",")
    If UBound(temp, 1) - LBound(temp, 1) + 1 <> cols_ct Then
        ctErr
        Exit Sub
    End If
ElseIf dims = 1 Then
    If UBound(Headers, 1) - LBound(Headers, 1) + 1 = cols_ct Then
        temp = Headers
    Else
        ctErr
        Exit Sub
    End If
ElseIf dims = 2 Then
    If UBound(Headers, 2) - LBound(Headers, 2) + 1 = cols_ct Then
        read_dim = 2
    ElseIf UBound(Headers, 1) - LBound(Headers, 1) + 1 = cols_ct Then
        read_dim = 1
    Else
        ctErr
        Exit Sub
    End If
    ReDim temp(LBound(Headers, read_dim) To UBound(Headers, read_dim))
    If read_dim = 1 Then
        For i = LBound(temp, 1) To UBound(temp, 1)
            temp(i) = Headers(i, 1)
        Next i
    Else
        For i = LBound(temp, 2) To UBound(temp, 2)
            temp(i) = Headers(1, i)
        Next i
    End If
Else
    MsgBox _
        Prompt:="AddHeaders method encountered an error:" & vbNewLine & _
                "Headers has more than two dimensions.", _
        Buttons:=vbOKOnly
    Exit Sub
End If

'loop through temp, adding item at position 1 to each collection in cols
For i = LBound(temp) To UBound(temp)
    cols(i).Add Trim(temp(i)), Before:=1
Next i

End Sub


Public Sub ConvertHeaders(HeaderPairs As Variant)

'changes headers in accordance with passed-in array:
'   "column" collections with first item equal to any odd-numbered element in Headers
'   have item deleted then new one with next element in Headers inserted
'   HeaderPairs is checked to ensure:
'       1) it is an array:
'           if not, procedure attempts to read it as a comma-delimited string (so do not separate values with " " if entering string)
'       2) a dimension has an even number of elements
'           procedure uses first dimension with even elements found (checks columns first)
'       otherwise it terminates

'-------------------------------------------------------------

'convert HeaderPairs to 1d array if possible and ensure even num of elements
Dim temp As Variant
Dim dims As Integer, read_dim As Integer, i As Integer
dims = GetArrDimensions(HeaderPairs)
If dims = 0 Then
    temp = Split(HeaderPairs, ",")
    If UBound(temp, 1) - LBound(temp, 1) + 1 Mod 2 <> 0 Then
        elemErr
        Exit Sub
    End If
ElseIf dims = 1 Then
    If UBound(HeaderPairs, 1) - LBound(HeaderPairs, 1) + 1 Mod 2 = 0 Then
        temp = HeaderPairs
    Else
        elemErr
        Exit Sub
    End If
ElseIf dims = 2 Then
    If UBound(HeaderPairs, 2) - LBound(HeaderPairs, 2) + 1 Mod 2 = 0 Then
        read_dim = 2
    ElseIf UBound(HeaderPairs, 1) - LBound(HeaderPairs, 1) + 1 Mod 2 = 0 Then
        read_dim = 1
    Else
        elemErr
        Exit Sub
    End If
    ReDim temp(LBound(HeaderPairs, read_dim) To UBound(HeaderPairs, read_dim))
    If read_dim = 1 Then
        For i = LBound(temp, 1) To UBound(temp, 1)
            temp(i) = HeaderPairs(i, 1)
        Next i
    Else
        For i = LBound(temp, 2) To UBound(temp, 2)
            temp(i) = HeaderPairs(1, i)
        Next i
    End If
Else
    MsgBox _
        Prompt:="ConvertHeaders method encountered an error:" & vbNewLine & _
                "More than two dimensions found in HeaderPairs.", _
        Buttons:=vbOKOnly
    Exit Sub
End If

'loop through temp, attempting to match odd elements to first items in column collections
Dim Str As String
Dim j As Integer
For i = LBound(temp) To UBound(temp) Step 2
    Str = Trim(temp(i))
    For j = 1 To cols.Count
        If Str = Trim(CStr(cols(j)(1))) Then
            cols(j).Remove (1)
            cols(j).Add Trim(temp(i + 1)), Before:=1
            Exit For
        End If
    Next j
Next i

End Sub

Public Sub RemoveHeaders()

Dim i As Integer
For i = 1 To cols.Count
    cols(i).Remove (1)
Next i

End Sub

Public Sub Update(Index As Integer, Key As Long, NewValue As Variant)

'updates the value in a given column collection at the given key value

'-----------------------------------------------------------------

With cols(Index)
    .Remove (Key)
    If Key <= .Count Then
        .Add NewValue, Before:=Key
    Else
        .Add NewValue
    End If
End With

End Sub

Public Function Index(header As String) As Integer

'returns index of passed-in header
'returns 0 if unable to find

'----------------------------------------

Dim i As Integer
For i = 1 To cols.Count
    If Trim(header) = Trim(CStr(cols(i)(1))) Then
        Index = i
        Exit For
    End If
Next i

End Function

Public Function Key(Str As String, Index As Integer) As Long

'returns "row" (index) in which passed-in string appears in "column" Index
'returns 0 if not found

'-------------------------------------------

Dim coll_ct As Long
coll_ct = cols(Index).Count

If coll_ct > 0 Then
    Dim l As Long
    For l = 1 To coll_ct
        If Trim(CStr(cols(Index)(l))) = Str Then
            Key = l
            Exit Function
        End If
    Next l
End If

End Function

Public Function Subset(Indices As Variant, Optional PreserveColumns As Boolean) As SuperCollection

'returns SC object with only passed-in columns
'columns returned in order they appear in Indices
'a comma-separated string or a 1d or 2d array may be passed to Indices
'   values not corresponding to "columns" are ignored
'   strings are ignored (index corresponding to a header may be found using Index method)
'if PreserveColumns=False, columns selected out will be deleted from object

'-------------------------------------------------------------------

'convert Indices to 1d array if possible
Dim temp As Variant
Dim dims As Integer, read_dim As Integer, i As Integer
dims = GetArrDimensions(Indices)
If dims = 0 Then
    temp = strToArr(CStr(Indices))
ElseIf dims = 1 Then
    temp = Indices
ElseIf dims = 2 Then
    If UBound(Indices, 2) - LBound(Indices, 2) + 1 > UBound(Indices, 1) - LBound(Indices, 1) + 1 Then
        read_dim = 2
    Else
        read_dim = 1
    End If
    ReDim temp(LBound(Indices, read_dim) To UBound(Indices, read_dim))
    If read_dim = 1 Then
        For i = LBound(temp, 1) To UBound(temp, 1)
            temp(i) = Indices(i, 1)
        Next i
    Else
        For i = LBound(temp, 2) To UBound(temp, 2)
            temp(i) = Indices(1, i)
        Next i
    End If
Else
    MsgBox _
        Prompt:="Subset method encountered an error:" & vbNewLine & _
                "More than two dimensions found in Indices.", _
        Buttons:=vbOKOnly
    Exit Function
End If

'compile new column collection to populate SuperCollection to be returned
Dim col_coll As New Collection
Dim Coll As Collection
Dim cur_index As Integer, col_ct As Integer, j As Integer
col_ct = cols.Count
For i = LBound(temp) To UBound(temp)
    cur_index = temp(i)
    If cur_index > 0 And cur_index <= col_ct Then
        Set Coll = New Collection
        col_ct = cols(cur_index).Count
        For j = 1 To col_ct
            If PreserveColumns = True Then
                Coll.Add cols(cur_index)(j)
            Else
                Coll.Add cols(cur_index)(1)
                cols(cur_index).Remove (1)
            End If
        Next j
        col_coll.Add Coll
        col_ct = cols.Count
    End If
Next i

Set Subset = New SuperCollection
Set Subset.ColumnCollection = col_coll

'remove specified cols from SC if user specifies
If PreserveColumns = False Then
    If LBound(temp) <> UBound(temp) Then
        For i = LBound(temp) To UBound(temp)
            cur_index = temp(i)
            If cur_index > 0 And cur_index <= col_ct Then
                cols.Remove (cur_index)
                For j = i + 1 To UBound(temp)
                    temp(j) = temp(j) - 1
                Next j
            End If
        Next i
    Else
        cols.Remove (temp(LBound(temp)))
    End If
End If

End Function

Public Function ToArray(Optional PreserveColumns As Boolean = False, Optional Tilted As Boolean = False) As Variant

'converts object to an array record-by-record
'if PreserveColumns=False, deletes records as it goes to save memory
'   in this case, empty cols will be all left of object after procedure completes
'if Tilted=True, then array with headers in column 1, rather than row 1, is returned

'----------------------------------------------

Dim return_arr As Variant
Dim i As Long
Dim j As Integer

If Tilted = False Then
    
    ReDim return_arr(1 To cols(1).Count, 1 To cols.Count)
    For i = 1 To UBound(return_arr, 1)
        For j = 1 To UBound(return_arr, 2)
            If PreserveColumns = False Then
                return_arr(i, j) = cols(j)(1)
                cols(j).Remove (1)
            Else
                return_arr(i, j) = cols(j)(i)
            End If
        Next j
    Next i
    
Else
    
    ReDim return_arr(1 To cols.Count, 1 To cols(1).Count)
    For i = 1 To UBound(return_arr, 2)
        For j = 1 To UBound(return_arr, 1)
            If PreserveColumns = False Then
                return_arr(j, i) = cols(j)(1)
                cols(j).Remove (1)
            Else
                return_arr(j, i) = cols(j)(i)
            End If
        Next j
    Next i
    
End If

ToArray = return_arr

End Function

Private Sub elemErr()

'displays msgbox explaining ConvertHeaders error: odd number of elements

'--------------------------------------------

MsgBox _
    Prompt:="ConvertHeaders method encountered an error:" & vbNewLine & _
            "HeaderPairs has an odd number of elements.", _
    Buttons:=vbOKOnly

End Sub

Private Sub ctErr()

'displays msgbox explaining AddHeaders error: no dimension with same number of elements as items in cols

'-----------------------------------------------------

MsgBox _
    Prompt:="AddHeaders method encountered an error:" & vbNewLine & _
            "Unable to find dimension with correct number of elements.", _
    Buttons:=vbOKOnly

End Sub

Private Function GetArrDimensions(arr As Variant) As Integer

'returns the number of dimensions of passed-in array

'----------------------------------------

Dim i As Long, dims As Integer
Dim done As Boolean
dims = 1

Do Until done = True
    On Error GoTo dims:
    i = UBound(arr, dims)
    On Error GoTo 0
    If done = True Then Exit Do
    dims = dims + 1
Loop

Exit Function

dims:
done = True
dims = dims - 1
GetArrDimensions = dims

Resume Next

End Function

Private Function strToArr(Str As String, Optional delim As String = ",") As Variant

'performs Split on str using passed in delimiter, then ensures all elements in resulting array are integers
'non-numeric values are deleted

'----------------------------------------------------------------------------

Dim return_arr As Variant
return_arr = Split(Str, delim)

Dim i As Integer, j As Integer, ubound_1 As Integer
i = 0
ubound_1 = UBound(return_arr, 1)
Dim change As Boolean
Do
    change = False
    On Error GoTo non_num:
    return_arr(i) = CInt(return_arr(i))
    On Error GoTo 0
    i = i + 1
Loop Until change = False And i = ubound_1 + 1

strToArr = return_arr

Exit Function

non_num:
For j = i To UBound(return_arr) - 1
    return_arr(j) = return_arr(j + 1)
Next j
ReDim Preserve return_arr(0 To UBound(return_arr) - 1)
ubound_1 = ubound_1 - 1
change = True
i = 0
Resume Next

End Function
