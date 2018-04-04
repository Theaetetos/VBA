VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class providing API between worksheet ranges, VBA arrays, and databases


Option Explicit


' Be sure to check whether these codes will cause collisions with any other custom error codes in project
Private Enum TableArrayErrorCode
    ATTR_MISSING = 513 ' attempting to call methods when one or more attributes are missing
    ARR_DIM = 514 ' attempting to call methods when object attributes do not have matching dimensions
    BAD_ARR = 515 ' attempting to set col_names or is_text using array of more than two dimensions
End Enum


Private recs As Variant
Private col_names As Variant
Private is_text As Variant
Private rng As Range
Private conn_str As String
Dim conn As ADODB.Connection


Property Let Records(RecordsArray As Variant)
    
    recs = to_2d(RecordsArray)
    
End Property

Property Get Records() As Variant
    
    Records = recs
    
End Property

Property Let ColumnNames(ColumnNamesArray As Variant)
    
    ' 1d (2d array can be passed in; will be converted automatically) array of column names
    
    '-----------------------------------------------------------------------------------
    
    col_names = to_1d(ColumnNamesArray) ' mismatch w/ recs allowed here, but not in class methods
    
End Property

Property Get ColumnNames() As Variant
    
    ColumnNames = col_names
    
End Property

Property Let IsText(IsTextArray As Variant)
    
    ' 1d (2d array can be passed in; will be converted automatically) array of Booleans specifying whether or not column is a text type
    
    ' -------------------------------------------------------------------------------------------------------------
    
    is_text = to_1d(IsTextArray) ' mismatch w/ recs arrays allowed here, but not in class methods
    
End Property

Property Get IsText() As Variant
    
    IsText = is_text
    
End Property

Property Set LinkedRange(RangeToLink As Range)
    
    ' Range to read data to or from
    
    ' -----------------------------------
    
    Set rng = RangeToLink
    
End Property

Property Get LinkedRange() As Range
    
    Set LinkedRange = rng
    
End Property

Property Let ConnectionString(ConnStr As String)
    
    ' String to use to open ADODB connection to database
    ' Opening and closing connections is handled by private methods
    
    ' -------------------------------------------------
    
    conn_str = ConnStr
    
End Property

Property Get ConnectionString() As String
    
    ConnectionString = conn_str
    
End Property

Private Function get_arr_dims(ByRef arr As Variant) As Integer
    
    'returns the number of dimensions of passed-in array
    
    '----------------------------------------
    
    Dim i As Long, dimens As Integer
    
    dimens = 1
    
    Do
        On Error GoTo dims:
        i = UBound(arr, dimens)
        On Error GoTo 0
        dimens = dimens + 1
    Loop
    
dims:
    dimens = dimens - 1
    get_arr_dims = dimens
    On Error GoTo 0
    Exit Function
    
End Function

Private Function to_2d(ByRef arr As Variant) As Variant
    
    ' takes as input single value, or 1d or 2d array and returns 2d array (2d array will be returned as is)
    ' a 1d array will simply be given a length-1 first dimension
    ' unravels by rows
    
    ' --------------------------------------------------------------------------------------------------------------------------------
    
    Dim ret As Variant
    Dim row_lbound As Long, row_ubound As Long, col_lbound As Long, col_ubound As Long, counter As Long, i As Long, j As Long
    
    Select Case get_arr_dims(arr)
        Case 0
            ReDim ret(1 To 1, 1 To 1) As Variant
            ret(1, 1) = arr
        Case 1
            ReDim Preserve arr(1 To 1, 1 To UBound(arr) - LBound(arr) + 1)
            ret = arr
        Case 2
            ret = arr
        Case Else
            Err.Raise Number:=vbObjectError + ARR_DIM, _
                      Description:="Attempted to set Records, ColumnNames, or IsText property using array with more than two dimensions"
    End Select
    
    to_2d = ret
    
End Function

Private Function to_1d(ByRef arr As Variant) As Variant
    
    ' takes as input single value, or 1d or 2d array and returns 1d array (1d array will be returned as is)
    ' flattens - a MxN 2d array will be turned into a 1d array of length M*N
    ' unravels by rows
    
    ' --------------------------------------------------------------------------------------------------------------------------------
    
    Dim ret As Variant
    Dim row_lbound As Long, row_ubound As Long, col_lbound As Long, col_ubound As Long, counter As Long, i As Long, j As Long
    
    Select Case get_arr_dims(arr)
        Case 0
            ReDim ret(1 To 1) As Variant
            ret(1) = arr
        Case 1
            ret = arr
        Case 2
            row_lbound = LBound(arr, 1)
            row_ubound = UBound(arr, 1)
            col_lbound = LBound(arr, 2)
            col_ubound = UBound(arr, 2)
            
            ReDim ret(1 To (row_ubound - row_lbound + 1) * (col_ubound - col_lbound + 1)) As Variant
            
            i = row_lbound
            j = col_lbound
            For counter = 1 To UBound(ret)
                ret(counter) = arr(i, j)
                i = i + 1
                If i > row_ubound Then
                    i = row_lbound
                    j = j + 1
                End If
            Next counter
        Case Else
            Err.Raise Number:=vbObjectError + ARR_DIM, _
                      Description:="Attempted to set Records, ColumnNames, or IsText property using array with more than two dimensions"
    End Select
    
    to_1d = ret
    
End Function

Private Sub check_attr(method_name As String, Optional foreign_instance As TableArray = Nothing)
    
    ' ensures all attributes necessary to execute sql statement 1) are set and 2) have matching dimensions (if applicable)
    ' raises ArrayToSQL errors ATTR_MISSING and ARR_DIM respectively if not
    ' optional foreign_instance param is for checking WhereClauseObject param in Update method
    
    ' -----------------------------------
    
    Dim recs_dim As Long, cols_dim As Long, it_dim As Long ' vars for comparing number of elements in class attributes
    
    If Not method_name = "SelectFrom" Then
        If IsEmpty(recs) Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.Records is empty"
        End If
    End If
    
    If IsEmpty(col_names) Then
        Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                  Description:="TableArray.ColumnNames is empty"
    End If
    
    If Not method_name = "SelectFrom" Then
        If IsEmpty(is_text) Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.IsText is empty"
        End If
    End If
    
    If conn_str = "" Then
        Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                  Description:="TableArray.ConnectionString is empty"
    End If
    
    If Not method_name = "SelectFrom" Then
        recs_dim = UBound(recs, 2) - LBound(recs, 2) + 1
        cols_dim = UBound(col_names) - LBound(col_names) + 1
        it_dim = UBound(is_text) - LBound(is_text) + 1
        
        If Not (recs_dim = cols_dim And cols_dim = it_dim) Then
            Err.Raise Number:=vbObjectError + ARR_DIM, _
                      Description:="TableArray attribute dimensions do not match." & vbNewLine & _
                                   "Records: " & CStr(recs_dim) & " elements" & vbNewLine & _
                                   "ColumnNames: " & CStr(cols_dim) & " elements" & vbNewLine & _
                                   "IsText: " & CStr(it_dim) & " elements"
        End If
    End If
    
    If method_name = "Update" Then
        If IsEmpty(foreign_instance.Records) Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.Records (foreign instance) is empty"
        End If
        
        If IsEmpty(foreign_instance.ColumnNames) Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.ColumnNames (foreign instance) is empty"
        End If
        
        If IsEmpty(foreign_instance.IsText) Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.IsText (foreign instance) is empty"
        End If
        
        recs_dim = UBound(foreign_instance.Records, 2) - LBound(foreign_instance.Records, 2) + 1
        cols_dim = UBound(foreign_instance.ColumnNames) - LBound(foreign_instance.ColumnNames) + 1
        it_dim = UBound(foreign_instance.IsText) - LBound(foreign_instance.IsText) + 1
        
        If Not (recs_dim = cols_dim And cols_dim = it_dim) Then
            Err.Raise Number:=vbObjectError + ARR_DIM, _
                      Description:="TableArray (foreign instance) attribute dimensions do not match." & vbNewLine & _
                                   "Records: " & CStr(recs_dim) & " elements" & vbNewLine & _
                                   "ColumnNames: " & CStr(cols_dim) & " elements" & vbNewLine & _
                                   "IsText: " & CStr(it_dim) & " elements"
        End If
    End If
    
End Sub

Private Function wrap_brackets(sql_obj_name As String) As String
    
    ' wraps SQL object name (db, table, or column) in brackets if not already
    
    '------------------------------------------------------------
    
    If Left(sql_obj_name, 1) <> "[" Then
        sql_obj_name = "[" & sql_obj_name
    End If
    
    If Right(sql_obj_name, 1) <> "]" Then
        sql_obj_name = sql_obj_name & "]"
    End If
    
    wrap_brackets = sql_obj_name
    
End Function

Private Sub execute_cmd(sql_str As String, Optional sproc As Boolean)
    
    ' executes sql_str that will not return any records using ConnectionString property
    ' caller responsible for opening instance conn attribute
    
    '--------------------------------------------------------------
    
    Dim cmd As New ADODB.Command
    
    With cmd
        .ActiveConnection = conn
        .CommandText = sql_str
        .CommandTimeout = 0
        If sproc Then
            .CommandType = adCmdStoredProc
        Else
            .CommandType = adCmdText
        End If
        .Execute
    End With
    
End Sub

Private Function to_array(ByRef coll As Collection) As Variant
    
    ' converts coll into array; coll is deleted along the way to preserve memory
    
    ' ---------------------------------------
    
    Dim return_arr As Variant
    Dim i As Long, j As Long
    
    On Error GoTo one_d
    ReDim return_arr(1 To coll(1).Count, 1 To coll.Count) As Variant
    On Error GoTo 0
    
    For i = 1 To UBound(return_arr, 1)
        For j = 1 To UBound(return_arr, 2)
            return_arr(i, j) = coll(j)(1)
            coll(j).Remove (1)
        Next j
    Next i
    
    to_array = return_arr
    
    Exit Function
    
one_d:
    ReDim return_arr(1 To coll.Count)
    
    For i = 1 To UBound(return_arr)
        return_arr(i) = coll(1)
        coll.Remove 1
    Next i
    
    to_array = return_arr
    
End Function

Private Sub execute_qry(sql_str As String, Optional sproc As Boolean)
    
    ' executes query defined by sql_str using ConnectionString property and saves result in recs
    ' caller responsible for opening instance conn attribute
    
    ' ----------------------------------------------------------
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim f As Field
    Dim rec_coll As New Collection, header_coll As New Collection, coll As Collection
    Dim i As Long
    
    With cmd
        .ActiveConnection = conn
        .CommandText = sql_str
        .CommandTimeout = 0
        If sproc Then
            .CommandType = adCmdStoredProc
        Else
            .CommandType = adCmdText
        End If
    End With
    
    With rs
        .Open Source:=cmd
        
        If .State = adStateOpen Then
            If Not .EOF Then
                'add field headers to headers collection and an empty collection to records collection for each field in recordset
                For Each f In .Fields
                    header_coll.Add f.Name
                    Set coll = New Collection
                    rec_coll.Add coll
                Next f
                
                'loop through fields and rows, adding record in each field to corresponding collection in query collection
                Do While Not .EOF
                    i = 1
                    For Each f In .Fields
                        Set coll = rec_coll(i)
                        coll.Add f.Value
                        Set coll = Nothing
                        i = i + 1
                    Next f
                    .MoveNext
                Loop
                
                col_names = to_1d(to_array(header_coll))
                recs = to_array(rec_coll)
            Else
                col_names = to_1d(to_array(header_coll))
                ' if recordset returns no records, then set recs to Empty
                recs = Empty
            End If
        .Close
        End If
        
        Set rs = Nothing
    End With
    
End Sub

Private Function get_arr_col_id(s As String, a As Variant) As Integer
    
    ' returns the colid within a passed-in array of a passed-in string
    ' if a is 2d, will only search first row
    ' returns -1 if s not found
    ' returns -2 if a is not 1d or 2d
    ' meant to mirror finding a table header
    
    '--------------------------------------------------
    
    Dim i As Long, dims As Integer
    
    Select Case get_arr_dims(a)
        Case 1
            i = LBound(a)
    
            get_arr_col_id = -1
            
            Do Until i > UBound(a)
                If Trim(CStr(a(i))) = s Then
                    get_arr_col_id = i
                    Exit Do
                End If
                i = i + 1
            Loop
        Case 2
            i = LBound(a, 2)
    
            get_arr_col_id = -1
            
            Do Until i > UBound(a, 2)
                If Trim(CStr(a(LBound(a, 1), i))) = s Then
                    get_arr_col_id = i
                    Exit Do
                End If
                i = i + 1
            Loop
        Case Else
            get_arr_col_id = -2
    End Select

End Function

Public Sub Insert(TableName As String)
    
    ' Inserts Records into specified table using Records, ColumnNames, IsText, and ConnectionString properties
    ' SQL statement will be concatenated like INSERT INTO TableName (ColumnNames(column)[, ...]) VALUES (Records(row, column)[, ...]);
    ' args:
    '  - TableName: name of table to insert recs into
    
    '----------------------------------------------------------------------------------------------------------------
    
    Dim i As Long, j As Long ' counters for looping through recs
    Dim sql_str_start As String, sql_str_end As String, sql_str As String ' vars for concatenating SQL string
    ' correction to apply to recs' lbound to reference correct items in col_names and is_text during loop
    Dim corr_col_names As Integer, corr_is_text As Integer
    
    check_attr "Insert"
    
    corr_col_names = LBound(col_names) - LBound(recs, 1)
    corr_is_text = LBound(is_text) - LBound(recs, 1)
    
    Set conn = New ADODB.Connection
    conn.Open ConnectionString
    
    For i = LBound(recs, 1) To UBound(recs, 1)
        sql_str_start = "INSERT INTO " & wrap_brackets(TableName) & " ("
        sql_str_end = ") VALUES ("
        
        j = LBound(recs, 2)
        Do While True
            sql_str_start = sql_str_start & wrap_brackets(CStr(col_names(j + corr_col_names)))
            
            If is_text(j + corr_is_text) Then
                sql_str_end = sql_str_end & CStr("'")
                sql_str_end = sql_str_end & CStr(recs(i, j))
                sql_str_end = sql_str_end & CStr("'")
            Else
                ' Empty values must be converted to 0 to avoid SQL errors
                sql_str_end = sql_str_end & CDbl(recs(i, j)) ' THIS WILL NOT WORK IF OTHER DATA TYPES BESIDES TEXT AND NUMERIC
            End If
            
            j = j + 1
            If j > UBound(recs, 2) Then
                GoTo break: ' hack to get around VBA's lack of a break statement
            End If
            
            sql_str_start = sql_str_start & ", "
            sql_str_end = sql_str_end & ", "
        Loop
break:
        sql_str = sql_str_start & sql_str_end & ");"
        
'        Debug.Print sql_str
        execute_cmd sql_str
    Next i
    
    conn.Close
    Set conn = Nothing
    
End Sub

Sub Update(TableName As String, WhereClauseObject As TableArray)
    
    ' Updates records in passed-in table using Records, ColumnNames, IsText, and ConnectionString properties
    ' Update statement is concatenated from method parameters and instance properties, like
    ' ... UPDATE TableName SET ColumnName(column) = Records(row, column) [AND ...];
    ' Where clause is concatenated from properties of WhereClauseObject, like
    ' ... WHERE WhereClauseObject.ColumnNames(column) = WhereClauseObject.Records(row, column) [AND ...];
    ' args:
    '  - TableName: name of table to insert recs into
    '  - WhereClauseObject: instance of TableArray class with properties necessary for concatenating WHERE clause of SQL statement
    
    '----------------------------------------------------------------------------------------------------------
    
    Dim i As Long, j As Long ' counters for looping through recs
    Dim sql_str As String ' var for concatenating SQL string
    ' corrections to apply to recs' lbound to reference correct items in *cols and *is_text during loop
    Dim corr_update_col_names As Integer, corr_update_is_text As Integer
    Dim corr_where_col_names As Integer, corr_where_is_text As Integer
    
    check_attr "Update", WhereClauseObject
    
    corr_update_col_names = LBound(col_names) - LBound(recs, 1)
    corr_update_is_text = LBound(is_text) - LBound(recs, 1)
    
    Set conn = New ADODB.Connection
    conn.Open ConnectionString
    
    With WhereClauseObject
        corr_where_col_names = LBound(.ColumnNames) - LBound(.Records, 1)
        corr_where_is_text = LBound(.IsText) - LBound(.Records, 1)
    End With
    
    For i = LBound(recs, 1) To UBound(recs, 1)
        sql_str = "UPDATE " & wrap_brackets(TableName) & " SET "
        
        j = LBound(recs, 2)
        Do While True
            sql_str = sql_str & wrap_brackets(CStr(col_names(j + corr_update_col_names))) & " = "
            
            If is_text(j + corr_update_is_text) Then
                sql_str = sql_str & CStr("'")
                sql_str = sql_str & CStr(recs(i, j))
                sql_str = sql_str & CStr("'")
            Else
                sql_str = sql_str & CDbl(recs(i, j)) ' THIS WILL NOT WORK IF OTHER DATA TYPES BESIDES TEXT AND NUMERIC
            End If
            
            j = j + 1
            If j > UBound(recs, 2) Then
                GoTo break1: ' hack to get around VBA's lack of a break statement
            End If
            
            sql_str = sql_str & ", "
        Loop
break1:
        sql_str = sql_str & " WHERE "
        
        With WhereClauseObject
            j = LBound(.Records, 2)
            Do While True
                sql_str = sql_str & wrap_brackets(CStr(.ColumnNames(j + corr_where_col_names))) & " = "
                If .IsText(j + corr_where_is_text) Then
                    sql_str = sql_str & CStr("'")
                    sql_str = sql_str & CStr(.Records(i, j))
                    sql_str = sql_str & CStr("'")
                Else
                    sql_str = sql_str & CDbl(.Records(i, j)) ' THIS WILL NOT WORK IF OTHER DATA TYPES BESIDES TEXT AND NUMERIC
                End If
                
                j = j + 1
                If j > UBound(.Records, 2) Then
                    GoTo break2: ' hack to get around VBA's lack of a break statement
                End If
                
                sql_str = sql_str & " AND "
            Loop
break2:
        End With
        
        sql_str = sql_str & ";"
        
'        Debug.Print sql_str
        execute_cmd sql_str
    Next i
    
    conn.Close
    Set conn = Nothing
    
End Sub

Public Sub Delete(TableName As String)
    
    ' Deletes records if exist in passed-in table using Records, ColumnNames, IsText, and ConnectionString properties
    ' For each row in Records, adds a condition in WHERE clause using ColumnNames and Records
    ' SQL statements will be concatenated like DELETE FROM TableName WHERE ColumnNames(column) = Records(row, column) [AND ...];
    ' args:
    '  - TableName: name of table to delete recs from
    
    '---------------------------------------
    
    Dim i As Long, j As Long ' counters for looping through recs
    Dim sql_str As String ' var for concatenating SQL string
    ' correction to apply to recs' lbound to reference correct items in col_names and is_text during loop
    Dim corr_col_names As Integer, corr_is_text As Integer
    
    check_attr "Delete"
    
    corr_col_names = LBound(col_names) - LBound(recs, 1)
    corr_is_text = LBound(is_text) - LBound(recs, 1)
    
    Set conn = New ADODB.Connection
    conn.Open ConnectionString
    
    For i = LBound(recs, 1) To UBound(recs, 1)
        sql_str = "DELETE FROM " & wrap_brackets(TableName) & " WHERE "
        
        j = LBound(recs, 2)
        Do While True
            If recs(i, j) <> Empty Then
                If Right(sql_str, 7) <> " WHERE " Then
                    sql_str = sql_str & " AND "
                End If
                
                sql_str = sql_str & wrap_brackets(CStr(col_names(j + corr_col_names))) & " = "
                
                If is_text(j + corr_is_text) Then
                    sql_str = sql_str & CStr("'")
                    sql_str = sql_str & CStr(recs(i, j))
                    sql_str = sql_str & CStr("'")
                Else
                    sql_str = sql_str & CDbl(recs(i, j)) ' THIS WILL NOT WORK IF OTHER DATA TYPES BESIDES TEXT AND NUMERIC
                End If
            End If
            
            j = j + 1
            If j > UBound(recs, 2) Then
                GoTo break: ' hack to get around VBA's lack of a break statement
            End If
        Loop
break:
        sql_str = sql_str & ";"
        
'        Debug.Print sql_str
        execute_cmd sql_str
    Next i
    
    conn.Close
    Set conn = Nothing
    
End Sub

Public Sub ClearLinkedRange(Optional RangeToLink As Range = Nothing, Optional Headers As Boolean)
    
    ' Clears contents of LinkedRange
    ' If provided, RangeToLink will be set as LinkedRange and cleared in one step
    ' Leaves header row if Headers is False
    ' Raises ATTR_MISSING error if not set
    
    '-----------------------------------
    
    If RangeToLink Is Nothing Then
        If rng Is Nothing Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.LinkedRange property is not set"
        End If
    Else
        Set rng = RangeToLink
    End If
    
    If rng Is Nothing Then
        Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                  Description:="TableArray.LinkedRange property is not set"
    End If
    
    With rng
        If Headers Then
            .ClearContents
        Else
            On Error Resume Next
            .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).ClearContents
            On Error GoTo 0
        End If
    End With
    
End Sub

Public Sub PrintToLinkedRange(Optional PreserveHeaders As Boolean = True)
    
    ' Prints Records to LinkedRange
    ' If PreserveHeaders is True, then matches ColumnNames to LinkedRange header row, leaving out columns that don't match;
    ' if False, clears header row and prints out ColumnNames as headers
    
    ' -------------------------------------------------------------------------------------
    
    Dim rng_headers As Variant, arr As Variant
    Dim i As Long, j As Long, col_idx As Long
    Dim recs_row_lbound As Integer
    
    If PreserveHeaders Then
        ClearLinkedRange
        
        recs_row_lbound = LBound(recs, 1)
        
        ReDim arr(1 To UBound(recs, 1) - recs_row_lbound + 2, 1 To UBound(recs, 2) - LBound(recs, 2) + 1) ' extra row is for headers
        With rng
            rng_headers = to_1d(.Resize(1, .Columns.Count).Value) ' to_1d used to handle case of single column
        End With
        
        For j = 1 To UBound(arr, 2)
            ' rng_headers must be allowed to be smaller than arr in second dimension
            ' case when arr is smaller is handled by controlling loop via UBound(arr, 2)
            If j > UBound(rng_headers) Then
                arr(1, j) = ""
            Else
                arr(1, j) = rng_headers(j)
            End If
        Next j
        
        For j = 1 To UBound(rng_headers)
            col_idx = get_arr_col_id(CStr(rng_headers(j)), col_names)
            If col_idx > -1 Then
                For i = 2 To UBound(arr, 1)
                    arr(i, j) = recs(i - (2 - recs_row_lbound), col_idx)
                Next i
            End If
        Next j
        
        With rng.Resize(UBound(arr, 1), UBound(arr, 2))
            .Value = arr
            .Columns.AutoFit
        End With
    Else
        ClearLinkedRange Headers:=True
        
        With rng
            .Resize(1, UBound(col_names) - LBound(col_names) + 1).Value = col_names
            .Offset(1, 0).Resize(UBound(recs, 1) - LBound(recs, 1) + 1, UBound(recs, 2) - LBound(recs, 2) + 1).Value = recs
            .CurrentRegion.Columns.AutoFit
        End With
    End If
    
End Sub

Public Sub ReadLinkedRange(Optional RangeToLink As Range = Nothing)
    
    ' Stores first row of LinkedRange in ColumnNames and rest in Records (LinkedRange must always have a header row)
    ' If provided, RangeToLink will be set as LinkedRange and read in one step
    ' User must set the IsText property later
    
    ' ------------------------------------------------------------------------------
    
    If RangeToLink Is Nothing Then
        If rng Is Nothing Then
            Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                      Description:="TableArray.LinkedRange property is not set"
        End If
    Else
        Set rng = RangeToLink
    End If
    
    With rng
        col_names = to_1d(.Resize(1, .Columns.Count).Value)
        ' if no records in linked range, then attempting to resize to 0 rows will throw an error
        ' recs is set to Empty in that case
        On Error GoTo no_recs
        recs = .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Value
        On Error GoTo 0
        recs = to_2d(recs)
    End With
    
    Exit Sub
    
no_recs:
    recs = Empty
    
End Sub

Public Sub SelectFrom(TableName As String, Optional WhereClause As String)
    
    ' Executes Select statement using TableName, ColumnNames, and ConnectionString and stores results in Records
    ' User can pass in a WHERE clase that will be added to statement concatenated from ColumnNames
    
    ' --------------------------------------------------------------------------
    
    Dim sql_str As String
    Dim i As Long, corr_where_col_names As Integer, corr_where_is_text As Integer
    
    check_attr "SelectFrom"
    
    sql_str = "SELECT [" & Join(col_names, "], [") & "] FROM " & wrap_brackets(TableName)
    
    If WhereClause <> "" Then
        sql_str = sql_str & WhereClause
    End If
    
    If Right(sql_str, 1) <> ";" Then
        sql_str = sql_str & ";"
    End If
    
'    Debug.Print sql_str
    Set conn = New ADODB.Connection
    conn.Open conn_str
    execute_qry sql_str
    conn.Close
    Set conn = Nothing
    
End Sub

Public Sub ExecuteSQL(SQL As String, ReturnValues As Boolean, StoredProcedure As Boolean)
    
    ' Executes raw SQL command using ConnectionString property
    ' If ReturnValues is true, saves records returned by SQL (set to False for statements like INSERT, etc.)
    ' If StoredProcedure is True, evaluates SQL parameter as name of stored procedure
    ' Provided to allow user full access to database's SQL engine
    
    ' ---------------------------------------------------------------
    
    If conn_str = "" Then
        Err.Raise Number:=vbObjectError + ATTR_MISSING, _
                  Description:="TableArray.ConnectionString is empty"
    End If
    
    Set conn = New ADODB.Connection
    conn.Open conn_str
    If ReturnValues Then
        execute_qry SQL, StoredProcedure
    Else
        execute_cmd SQL, StoredProcedure
    End If
    conn.Close
    Set conn = Nothing
    
End Sub
