VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tbl_rng As Range
Private tbl_arr As Variant
Private qry_coll As New Collection

Property Set TableRange(Rng As Range)
    
    'sets tbl_rng property and fills tbl_arr
    
    '--------------------------
    
    Set tbl_rng = Rng
    
    tbl_arr = tbl_rng.Value
    
End Property

Property Get TableRange() As Range
    
    Set TableRange = tbl_rng
    
End Property

Property Let TableArray(arr As Variant)
    
    tbl_arr = arr
    
End Property

Property Get TableArray() As Variant
    
    TableArray = tbl_arr
    
End Property

Property Get QueryCollection() As Collection
    
    Set QueryCollection = qry_coll
    
End Property

Public Sub ExecuteQuery(ConnectionString As String, SQL As String, Optional Timeout As Integer = 120, Optional Sproc As Boolean)
    
    'executes passed-in sql str using passed-in connection str
    'saves result as collection of fields (also collections) in qry_coll
    
    'method of saving as collections chosen because less likely to cause overflow with large datasets
    '   largest query successfully saved: 541000x22 after previous failure with arrays (including GetRows method of Recordset class) (yes, it took forever)
    
    '----------------------------------------------------------------------------------------------------------------------
    
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim query As New Collection 'collection of fields returned in recordset
    Dim coll As Collection 'collection of records (only need one object; set dynamically in loop)
    Dim f As Field
    Dim i As Long 'counter
    
    Application.StatusBar = "TableInterface: Opening connection..."
    conn.Open ConnectionString
    
    With cmd
        .ActiveConnection = conn
        .CommandText = SQL
        .CommandTimeout = Timeout
        If Sproc Then
            .CommandType = adCmdStoredProc
        Else
            .CommandType = adCmdText
        End If
    End With
    
    With rs
        
        Application.StatusBar = "TableInterface: Executing query..."
        .Open Source:=cmd
        
        If .State = adStateOpen Then
            
            Application.StatusBar = "TableInterface: Saving query results..."
            
            If Not .EOF Then
                
                'add an empty collection to query collection for each field in recordset
                For i = 1 To .Fields.Count
                    Set coll = New Collection
                    query.Add coll
                Next i
                
                'add field headers at beginning of each collection
                i = 1
                For Each f In .Fields
                    Set coll = query(i)
                    coll.Add f.Name
                    Set coll = Nothing
                    i = i + 1
                Next f
                
                'loop through fields, adding record in each field to corresponding collection in query collection
                Do While Not .EOF
                    i = 1
                    For Each f In .Fields
                        Set coll = query(i)
                        coll.Add f.Value
                        Set coll = Nothing
                        i = i + 1
                    Next f
                    .MoveNext
                Loop
                
            Else
                
                ' if recordset returns no records, then save a 1x1 collections with single item "NULL"
                Set coll = New Collection
                coll.Add "NULL"
                query.Add coll
                
            End If
        
        .Close
        
        End If
        
        Set rs = Nothing
        
    End With
    
    conn.Close
    Set conn = Nothing
    
    qry_coll.Add query
    
    Application.StatusBar = False
    
End Sub

Public Sub EditTableArray(RowIndex As Long, ColumnIndex As Integer, NewValue As Variant)
    
    'updates a single value at passed-in indices in tbl_arr
    
    '-------------------------------------------------
    
    tbl_arr(RowIndex, ColumnIndex) = NewValue
    
End Sub

Public Sub ResizeTable(arr As Variant, Optional KeepColumns As Boolean = True)
    
    'resizes tbl_rng and tbl_arr to fit passed-in arr dimensions
    
    '---------------------------------------------------------
    
    Dim row_ct As Long
    Dim col_ct As Integer
    
    row_ct = UBound(arr, 1) - LBound(arr, 1) + 1
    
    Set tbl_rng = tbl_rng.Resize(row_ct, tbl_rng.Columns.Count)
    
    If KeepColumns = False Then
        col_ct = UBound(arr, 2) - LBound(arr, 1) + 1
        Set tbl_rng = tbl_rng.Resize(row_ct, col_ct)
    End If
    
    tbl_arr = tbl_rng.Value
    
End Sub

Public Sub Fill(Data As Variant)
    
    'fills tbl_arr using passed-in data; matches field headers (and so must exist in data)
    'data is read column-wise, i.e. it is shaped like a table
    'ADD:
    '   param checks:
    
    '----------------
    
    'check that Data is a 2d array
    
    Dim j As Long
    Dim data_col_id As Integer, i As Integer, corr As Integer
    
    If LBound(tbl_arr, 1) = 0 And LBound(Data, 1) = 1 Then corr = 1
    If LBound(tbl_arr, 1) = 1 And LBound(Data, 1) = 0 Then corr = -1
    
    For i = LBound(tbl_arr, 2) To UBound(tbl_arr, 2)
        data_col_id = GetArrColId(CStr(tbl_arr(1, i)), Data)
        If data_col_id > -1 Then
            For j = LBound(tbl_arr, 1) + 1 To UBound(tbl_arr, 1)
                tbl_arr(j, i) = Data(corr + j, data_col_id)
            Next j
        End If
    Next i
    
End Sub

Public Sub PrintTable()
    
    'prints tbl_arr to tbl_rng
    'resizes tbl_rng if dimensions have changed
    
    '-------------------------------
    
    Dim arr_row_ct As Long
    Dim arr_col_ct As Integer
    
    arr_row_ct = UBound(tbl_arr, 1) - LBound(tbl_arr, 1) + 1
    arr_col_ct = UBound(tbl_arr, 2) - LBound(tbl_arr, 2) + 1
    
    Set tbl_rng = tbl_rng.Resize(arr_row_ct, arr_col_ct)
    
    With tbl_rng
        .Value = tbl_arr
        .Columns.AutoFit
    End With
    
End Sub

Public Function GetQueryArray(Index As Integer, Optional Header As Boolean = True) As Variant
    
    'returns a base-1 array of the query results stored at the passed-in index
    'dynamically deletes query results to save memory
    'if Header is False, does not include header row
    
    '----------------------------------------------
    
    Application.StatusBar = "Saving TableInterface query " & Index & " as array..."
    
    Dim return_arr As Variant
    Dim i As Long
    Dim j As Integer, st As Integer
    
    If Header Then
        st = 1
    Else
        st = 2
    End If
    
    ReDim return_arr(1 To qry_coll(Index)(1).Count + 1 - st, 1 To qry_coll(Index).Count)
    
    For i = 1 To UBound(return_arr, 1)
        For j = 1 To UBound(return_arr, 2)
            return_arr(i, j) = qry_coll(Index)(j)(st)
            qry_coll(Index)(j).Remove st
        Next j
    Next i
    
    qry_coll.Remove Index
    
    GetQueryArray = return_arr
    
    Application.StatusBar = False
    
End Function

Private Function GetArrColId(s As String, a As Variant) As Integer
    
    'returns the colid within a passed-in array of a passed-in string
    '   although array must be two-dimensional for function to work, will only search first row
    '   will find match of any type implicitly convertible to string
    'returns -1 if s not found
    'returns -2 if a is not 2d
    'meant to mirror finding a table header
    
    '---------------------------------------------------
    
    'check dimensions of a
    If GetArrDimensions(a) <> 2 Then
        GetArrColId = -2
        Exit Function
    End If
    
    Dim i As Integer
    i = LBound(a, 2)
    GetArrColId = -1
    
    'loop through first row until s is found
    Do Until i > UBound(a, 2)
        If Trim(CStr(a(LBound(a, 1), i))) = s Then
            GetArrColId = i
            Exit Do
        End If
        i = i + 1
    Loop
    
End Function

Private Function GetArrDimensions(arr As Variant) As Integer
    
    'returns the number of dimensions of passed-in array
    
    '----------------------------------------
    
    Dim i As Long, dims As Integer
    dims = 1
    Dim done As Boolean
    
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

