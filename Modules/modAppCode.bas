Attribute VB_Name = "modAppCode"

Public CN                           As New ADODB.Connection
Public SQLstr                       As String
Public DBPathFileName               As String
Public CloseMe                      As Boolean

Public CurrUser                     As USERinfo
Public CurrBusinessInfo             As BusinessInfo
Public CurrInven                    As Inventory

Public Enum FormState
    AddStateMode = 0
    EditStateMode = 1
    adStatePopupMode = 2
End Enum
 
Public Type USERinfo
    UserNAME                        As String
    UserPK                          As String
End Type

Public Type BusinessInfo
    BusinessName                    As String
    BusinessAddress                 As String
End Type

Public Type Inventory
    TotalSupplier                   As String
    TotalProduct                    As String
    TotalAmount                     As String
End Type
 
Public isServer                            As String
Public isUsername                          As String
Public isPASSWORD                          As String
Public isPORT                              As String
Public isDatabase                          As String

  Public Function ConnectDB() As Boolean
    Dim isOpen      As Boolean
    Dim ANS         As VbMsgBoxResult
    
    isServer = "localhost" 'Pagclient Enter the IP Address HEre!
    isUsername = "root"
    isPASSWORD = ""
    isPORT = "3306"
    isDatabase = "DB"
    
    isOpen = False

 DBPathFileName = App.Path & "\DB.mdb"

'    On Error GoTo err
        Do Until isOpen = True
                CN.CursorLocation = adUseClient
                CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password=;"
                 'CN.ConnectionString = "DSN=DB"
                'CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPathFileName & ";Persist Security Info=False;Jet OLEDB:Database Password=;"
            isOpen = True
        Loop
        ConnectDB = isOpen
    Exit Function
err:
    ANS = MsgBox("Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical + vbRetryCancel)
    If ANS = vbCancel Then
        ConnectDB = False
    ElseIf ANS = vbRetry Then
        Resume
    End If
End Function

Public Sub CloseDB()
    'Close the connection
    CN.Close
    Set CN = Nothing
End Sub
Public Sub clearText(ByRef sForm As Form)
    Dim Control As Control
    For Each Control In sForm.Controls
        If (TypeOf Control Is TextBox) Then Control = vbNullString
    Next Control
    Set Control = Nothing
End Sub

Public Sub openRec(ByRef rec As ADODB.Recordset, ByVal Table_Name As String, Optional Condition As String)
Set rec = New ADODB.Recordset
With rec
Set .ActiveConnection = CN
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Source = "SELECT * FROM " & Table_Name & " " & Condition
.Open
End With
End Sub

Public Sub set_rec_getData(ByRef sRecordset As ADODB.Recordset, ByRef sConnection As ADODB.Connection, ByVal sSQL As String)
With sRecordset
    .CursorLocation = adUseClient
    .Open sSQL, sConnection, adOpenKeyset, adLockOptimistic
End With
End Sub

Public Function isRecordExist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, ByRef sEntryField) As Boolean
    Dim RS As New ADODB.Recordset
    RS.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", CN, adOpenStatic, adLockReadOnly
    If RS.RecordCount < 1 Then
        isRecordExist = False
    Else
        MsgBox "The adding of new entry cannot be done because '" & sStr & "' is already" & vbCrLf & "exist in the record.Please check and change it." & vbCrLf & vbCrLf & "Note: Duplication of entries is not allowed in this application.", vbExclamation
        isRecordExist = True
    End If
    Set RS = Nothing
End Function
 
 
Public Sub ToUpper(ByRef sText As TextBox)
    sText = UCase(sText)
 End Sub

Public Function toCurr(ByVal num As Double) As String
toCurr = Format(num, "#,##0.00")
End Function

Public Function rec_found(ByRef sRS As ADODB.Recordset, ByVal sField As String, ByVal sFindText As String) As Boolean
sRS.Requery
sRS.Find sField & " = '" & sFindText & "'"
If sRS.EOF Then
    rec_found = False
Else
    rec_found = True
End If
End Function


Public Function GetTxtVal(ByVal sTxt As String) As Double
    Dim sNew As String
    Dim sC As String
    Dim i As Integer
    'default
    GetTxtVal = 0
    sTxt = Trim(sTxt)
    If Len(sTxt) > 0 Then
        For i = 1 To Len(sTxt)
            sC = Mid(sTxt, i, 1)
            If sC = "-" Or sC = "." Or sC = "1" Or sC = "2" Or sC = "3" Or sC = "4" Or sC = "5" Or sC = "6" Or sC = "7" Or sC = "8" Or sC = "9" Or sC = "0" Then
                sNew = sNew & sC
            End If
        Next
        If Len(sNew) > 0 Then
            GetTxtVal = Val(sNew)
        End If
    End If
End Function
 
 
Public Function GenerateID(ByVal srcNo As String, ByVal src1stStr As String, ByVal src2ndStr As String) As String
    If Len(src2ndStr) <= Len(srcNo) Then
        GenerateID = src1stStr & srcNo
    Else
        GenerateID = src1stStr & Left$(src2ndStr, Len(src2ndStr) - Len(srcNo)) & srcNo
    End If
End Function

Public Sub HighL(ByRef sText)
    On Error Resume Next
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo err
    Dim RS As New Recordset
    Dim RI As Long
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblAutoNumber WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = RS.Fields("NextNumber")
    RS.Fields("NextNumber") = RI + 1
    RS.Update
    
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set RS = Nothing
    Exit Function
err:
        ''Error when incounter a null value
        If err.Number = 94 Then getIndex = 1: Resume Next
End Function

Public Sub loadForm(Frm As Form)
Frm.WindowState = vbMaximized
Frm.Show
Frm.SetFocus
End Sub
 
'Function that return true if the control is empty
Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        is_empty = True
        If UseTagValue = True Then
            MsgBox "The field '" & sText.Tag & "' is required.Please check it!", vbExclamation
        Else
            MsgBox "The field is required.Please check it!", vbExclamation
        End If
        sText.SetFocus
    Else
        is_empty = False
    End If
End Function
 
Public Function changeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeYNValue = "1"
        Case "N": changeYNValue = "0"
        Case "1": changeYNValue = "Y"
        Case "0": changeYNValue = "N"
    End Select
End Function

Public Function isNumber(ByVal sKeyAscii) As Integer
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isNumber = 0
    Else
        isNumber = sKeyAscii
    End If
End Function

Public Function SortLV(ByRef lv As ListView, Optional HeaderIndex As Integer = 0, Optional newSortOrder As ListSortOrderConstants = lvwAscending, Optional AutoOrder As Boolean = True)
    Dim lvHeader As ColumnHeader
    If AutoOrder = True Then
        If lv.SortOrder = lvwAscending Then
           lv.SortOrder = lvwDescending
        Else
           lv.SortOrder = lvwAscending
        End If
    Else
        lv.SortOrder = newSortOrder
    End If
    
    If HeaderIndex > lv.ColumnHeaders.Count - 1 Then
        HeaderIndex = 0
    End If
    
    lv.SortKey = HeaderIndex
    lv.Sorted = True
    lv.Refresh
    
    For Each lvHeader In lv.ColumnHeaders
        lvHeader.Icon = 0
    Next
    
    On Error Resume Next
    lv.ColumnHeaders(HeaderIndex + 1).Icon = lv.SortOrder + 1
End Function

Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function

Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0.00")
End Function

Public Sub FindRec(ByRef sRS As ADODB.Recordset, ByVal sField As String, ByVal isString As Boolean, ByVal sStr As String, ByVal sNum As Long)
    Dim old_pos As Long
    Dim sqlParam As String
    With sRS
        old_pos = .AbsolutePosition
        sRS.Filter = adFilterNone
        sRS.Requery
        .MoveFirst
        If isString = True Then
            sqlParam = sField & " = '" & sStr & "'"
        Else
            sqlParam = sField & " = " & sNum
        End If
        .Find sqlParam
        If .EOF Then .AbsolutePosition = old_pos
    End With
    old_pos = 0
    sqlParam = ""
End Sub


 
Public Function DeCode(vText As String) As String

    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    CurSpc = CurSpc + 1
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        varChr = Mid(vText, CurSpc, 3)
        Select Case varChr
            'lower case
            Case "coe"
                varChr = "a"
            Case "wer"
                varChr = "b"
            Case "ibq"
                varChr = "c"
            Case "am7"
                varChr = "d"
            Case "pm1"
                varChr = "e"
            Case "mop"
                varChr = "f"
            Case "9v4"
                varChr = "g"
            Case "qu6"
                varChr = "h"
            Case "zxc"
                varChr = "i"
            Case "4mp"
                varChr = "j"
            Case "f88"
                varChr = "k"
            Case "qe2"
                varChr = "l"
            Case "vbn"
                varChr = "m"
            Case "qwt"
                varChr = "n"
            Case "pl5"
                varChr = "o"
            Case "13s"
                varChr = "p"
            Case "c%l"
                varChr = "q"
            Case "w$w"
                varChr = "r"
            Case "6a@"
                varChr = "s"
            Case "!2&"
                varChr = "t"
            Case "(=c"
                varChr = "u"
            Case "wvf"
                varChr = "v"
            Case "dp0"
                varChr = "w"
            Case "w$-"
                varChr = "x"
            Case "vn&"
                varChr = "y"
            Case "c*4"
                varChr = "z"
            'numbers
            Case "aq@"
                varChr = "1"
            Case "902"
                varChr = "2"
            Case "2.&"
                varChr = "3"
            Case "/w!"
                varChr = "4"
            Case "|pq"
                varChr = "5"
            Case "ml|"
                varChr = "6"
            Case "t'?"
                varChr = "7"
            Case ">^s"
                varChr = "8"
            Case "<s^"
                varChr = "9"
            Case ";&c"
                varChr = "0"
            'caps
            Case "$)c"
                varChr = "A"
            Case "-gt"
                varChr = "B"
            Case "|p*"
                varChr = "C"
            Case "1" & Chr(34) & "r"
                varChr = "D"
            Case "c>:"
                varChr = "E"
            Case "@+x"
                varChr = "F"
            Case "v^a"
                varChr = "G"
            Case "]eE"
                varChr = "H"
            Case "aP0"
                varChr = "I"
            Case "{=1"
                varChr = "J"
            Case "cWv"
                varChr = "K"
            Case "cDc"
                varChr = "L"
            Case "*,!"
                varChr = "M"
            Case "fW" & Chr(34)
                varChr = "N"
            Case ".?T"
                varChr = "O"
            Case "%<8"
                varChr = "P"
            Case "@:a"
                varChr = "Q"
            Case "&c$"
                varChr = "R"
            Case "WnY"
                varChr = "S"
            Case "{Sh"
                varChr = "T"
            Case "_%M"
                varChr = "U"
            Case "}'$"
                varChr = "V"
            Case "QlU"
                varChr = "W"
            Case "Im^"
                varChr = "X"
            Case "l|P"
                varChr = "Y"
            Case ".>#"
                varChr = "Z"
            'Special characters
            Case "\" & Chr(34) & "]"
                varChr = "!"
            Case "cY,"
                varChr = "@"
            Case "x%B"
                varChr = "#"
            Case "a*v"
                varChr = "$"
            Case "'&T"
                varChr = "%"
            Case ";%R"
                varChr = "^"
            Case "eG_"
                varChr = "&"
            Case "Z/e"
                varChr = "*"
            Case "rG\"
                varChr = "("
            Case "]*F"
                varChr = ")"
            Case "@B*"
                varChr = "_"
            Case "+Hc"
                varChr = "-"
            Case "&|D"
                varChr = "="
            Case "(:#"
                varChr = "+"
            Case "SlW"
                varChr = "["
            Case "'QB"
                varChr = "]"
            Case "{D>"
                varChr = "{"
            Case "+c%"
                varChr = "}"
            Case "(s:"
                varChr = ":"
            Case "^a("
                varChr = ";"
            Case "16."
                varChr = "'"
            Case "s.*"
                varChr = Chr(34)
            Case "&?W"
                varChr = ","
            Case "GPQ"
                varChr = "."
            Case "SK*"
                varChr = "<"
            Case "RL^"
                varChr = ">"
            Case "40C"
                varChr = "/"
            Case "?#9"
                varChr = "?"
            Case "_?/"
                varChr = "\"
            Case "(_@"
                varChr = "|"
            Case "=#B"
                varChr = " "
        End Select
        varFin = varFin & varChr
        CurSpc = CurSpc + 3
        DoEvents
    Loop
    DeCode = varFin
    Exit Function

End Function

Public Function Encode(vText As String)

    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        CurSpc = CurSpc + 1
        varChr = Mid(vText, CurSpc, 1)
        Select Case varChr
            'lower case
            Case "a"
                varChr = "coe"
            Case "b"
                varChr = "wer"
            Case "c"
                varChr = "ibq"
            Case "d"
                varChr = "am7"
            Case "e"
                varChr = "pm1"
            Case "f"
                varChr = "mop"
            Case "g"
                varChr = "9v4"
            Case "h"
                varChr = "qu6"
            Case "i"
                varChr = "zxc"
            Case "j"
                varChr = "4mp"
            Case "k"
                varChr = "f88"
            Case "l"
                varChr = "qe2"
            Case "m"
                varChr = "vbn"
            Case "n"
                varChr = "qwt"
            Case "o"
                varChr = "pl5"
            Case "p"
                varChr = "13s"
            Case "q"
                varChr = "c%l"
            Case "r"
                varChr = "w$w"
            Case "s"
                varChr = "6a@"
            Case "t"
                varChr = "!2&"
            Case "u"
                varChr = "(=c"
            Case "v"
                varChr = "wvf"
            Case "w"
                varChr = "dp0"
            Case "x"
                varChr = "w$-"
            Case "y"
                varChr = "vn&"
            Case "z"
                varChr = "c*4"
            'numbers
            Case "1"
                varChr = "aq@"
            Case "2"
                varChr = "902"
            Case "3"
                varChr = "2.&"
            Case "4"
                varChr = "/w!"
            Case "5"
                varChr = "|pq"
            Case "6"
                varChr = "ml|"
            Case "7"
                varChr = "t'?"
            Case "8"
                varChr = ">^s"
            Case "9"
                varChr = "<s^"
            Case "0"
                varChr = ";&c"
            'caps
            Case "A"
                varChr = "$)c"
            Case "B"
                varChr = "-gt"
            Case "C"
                varChr = "|p*"
            Case "D"
                varChr = "1" & Chr(34) & "r"
            Case "E"
                varChr = "c>:"
            Case "F"
                varChr = "@+x"
            Case "G"
                varChr = "v^a"
            Case "H"
                varChr = "]eE"
            Case "I"
                varChr = "aP0"
            Case "J"
                varChr = "{=1"
            Case "K"
                varChr = "cWv"
            Case "L"
                varChr = "cDc"
            Case "M"
                varChr = "*,!"
            Case "N"
                varChr = "fW" & Chr(34)
            Case "O"
                varChr = ".?T"
            Case "P"
                varChr = "%<8"
            Case "Q"
                varChr = "@:a"
            Case "R"
                varChr = "&c$"
            Case "S"
                varChr = "WnY"
            Case "T"
                varChr = "{Sh"
            Case "U"
                varChr = "_%M"
            Case "V"
                varChr = "}'$"
            Case "W"
                varChr = "QlU"
            Case "X"
                varChr = "Im^"
            Case "Y"
                varChr = "l|P"
            Case "Z"
                varChr = ".>#"
            'Special characters
            Case "!"
                varChr = "\" & Chr(34) & "]"
            Case "@"
                varChr = "cY,"
            Case "#"
                varChr = "x%B"
            Case "$"
                varChr = "a*v"
            Case "%"
                varChr = "'&T"
            Case "^"
                varChr = ";%R"
            Case "&"
                varChr = "eG_"
            Case "*"
                varChr = "Z/e"
            Case "("
                varChr = "rG\"
            Case ")"
                varChr = "]*F"
            Case "_"
                varChr = "@B*"
            Case "-"
                varChr = "+Hc"
            Case "="
                varChr = "&|D"
            Case "+"
                varChr = "(:#"
            Case "["
                varChr = "SlW"
            Case "]"
                varChr = "'QB"
            Case "{"
                varChr = "{D>"
            Case "}"
                varChr = "+c%"
            Case ":"
                varChr = "(s:"
            Case ";"
                varChr = "^a("
            Case "'"
                varChr = "16."
            Case Chr(34)
                varChr = "s.*"
            Case ","
                varChr = "&?W"
            Case "."
                varChr = "GPQ"
            Case "<"
                varChr = "SK*"
            Case ">"
                varChr = "RL^"
            Case "/"
                varChr = "40C"
            Case "?"
                varChr = "?#9"
            Case "\"
                varChr = "_?/"
            Case "|"
                varChr = "(_@"
            Case " "
                varChr = "=#B"
        End Select
        varFin = varFin & varChr
        DoEvents
    Loop
    Encode = varFin
    Exit Function

End Function




