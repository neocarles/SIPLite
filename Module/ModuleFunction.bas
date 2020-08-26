Attribute VB_Name = "ModuleFunction"
Public Function Bulatkan(nValue As Double, nDigits As Integer) As Double
    Bulatkan = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
End Function

Public Function TahunK() As String
    TahunK = Format(Now, "yyyy") - 20
End Function

Public Function TahunT() As String
    TahunT = Format(Now, "yyyy") + 20
End Function

Public Function ComboList(objectToIgnore As ComboBox, Rs As ADODB.Recordset, Table As String, field As String, where As String, contents1 As String, contents2 As String)
On Error GoTo ErrComboList
    'rs.Close
    SQL = "select " & field & " from " & Table & " " & where
    Rs.Open SQL, Conn, adOpenDynamic, adLockPessimistic
        While Not Rs.EOF
            objectToIgnore.AddItem Rs.Fields(contents1) & " - " & Rs.Fields(contents2)
            Rs.MoveNext
        Wend
    'rs.Close
Exit Function
ErrComboList:
    MsgBox "Maaf data tidak bisa ditampilkan di combobox." + err.Description, vbCritical
End Function

Public Function ComboListSingle(objectToIgnore As ComboBox, Rs As ADODB.Recordset, Table As String, field As String, where As String, contents1 As String)
On Error GoTo ErrComboList
    'rs.Close
    SQL = "select " & field & " from " & Table & " " & where
    Rs.Open SQL, Conn, adOpenDynamic, adLockPessimistic
        While Not Rs.EOF
            objectToIgnore.AddItem Rs.Fields(contents1)
            Rs.MoveNext
        Wend
    'rs.Close
Exit Function
ErrComboList:
    MsgBox "Maaf data tidak bisa ditampilkan di combobox." + err.Description, vbCritical
End Function

'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
'On Error Resume Next
    Dim RSQuery As New ADODB.Recordset
    '----------------------------- Koneksi ---------------------------------------------
    Set RSQuery = New ADODB.Recordset
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    'Dim rs As New Recordset

    RSQuery.CursorLocation = adUseClient
    RSQuery.Open "SELECT COUNT(id) as TCount FROM " & srcTable & srcCondition, Conn, adOpenStatic, adLockReadOnly
    'RSQuery.Open "SELECT * FROM " & srcTable & srcCondition, Conn, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getRecordCount = Format$(RSQuery![TCount], "#,##0")
    Else
        getRecordCount = RSQuery![TCount]
    End If
    Set RSQuery = Nothing
End Function

'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
    toMoney = Format$(srcCurr, "#,##0")
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String) As Double
    On Error GoTo err
    Dim RSSum As New ADODB.Recordset
    Set RSSum = New ADODB.Recordset

    RSSum.CursorLocation = adUseClient
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " HAVING(" & sCondition & ")"
    If inclField <> "" Then inclField = "," & inclField
    RSSum.Open "SELECT Sum(" & sTable & "." & sField & ") AS fTotal" & inclField & " FROM " & sTable & sCondition, sCN, adOpenStatic, adLockOptimistic
    If RSSum.RecordCount > 0 Then
        RSSum.MoveFirst
        Do While Not RSSum.EOF
            getSumOfFields = getSumOfFields + RSSum.Fields("fTotal")
            RSSum.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If

    Set RSSum = Nothing
    Exit Function
err:
    'Error when incounter a null value
    If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function
