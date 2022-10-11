Attribute VB_Name = "ModFunctions"
Option Explicit

Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)
Public Const CB_SHOWDROPDOWN = &H14F
Declare Function SendMessage Lib "user32" Alias _
                                 "SendMessageA" _
                                 (ByVal hwnd As Long, _
                                  ByVal wMsg As Long, _
                                  ByVal wParam As Long, _
                                  lParam As Any) As Long

Private Const sAS_AutoBackup As String = "AutoBackup"

Public Function SetAutoBackup(ByVal NewValue As Boolean)
    Dim sValue As String
    SaveSetting App.Title, "AppSetting", sAS_AutoBackup, IIf(NewValue, "T", "F")
End Function

Public Function GetAutoBackup() As Boolean
    Dim sValue As String
    'default
    GetAutoBackup = -1
    sValue = GetSetting(App.Title, "AppSetting", sAS_AutoBackup, "T")
    GetAutoBackup = IIf(sValue = "T", True, False)
End Function

'Function used to check if the record exit or not.
Public Function isRecordExist(ByVal sTable As String, ByVal sField As String, ByVal sSTR As String, Optional isString As Boolean) As Boolean
    Dim Rs As New Recordset

    Rs.CursorLocation = adUseClient
    If isString = False Then
        Rs.Open "Select * From " & sTable & " Where " & sField & " = " & sSTR, CN, adOpenStatic, adLockOptimistic
    Else
        Rs.Open "Select * From " & sTable & " Where " & sField & " = '" & sSTR & "'", CN, adOpenStatic, adLockOptimistic
    End If
    If Rs.RecordCount < 1 Then
        isRecordExist = False
    Else
        isRecordExist = True
    End If
    Set Rs = Nothing
End Function

'Function used to check if the Ascii is a number or not (return 0 if number)
Public Function isNumber(ByVal sKeyAscii) As Integer
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isNumber = 0
    Else
        isNumber = sKeyAscii
    End If
End Function

'Function used to left split user fields
Public Function LeftSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then LeftSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = 1 To Len(srcUF)
        If Mid$(srcUF, i, 7) = "*~~~~~*" Then
            Exit For
        Else
            t = t & Mid$(srcUF, i, 1)
        End If
    Next i
    LeftSplitUF = t
    i = 0
    t = ""
End Function

'Function used to right split user fields
Public Function RightSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then RightSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = (InStr(1, srcUF, "*~~~~~*", vbTextCompare) + 7) To Len(srcUF)
        t = t & Mid$(srcUF, i, 1)
    Next i
    RightSplitUF = t
    i = 0
    t = ""
End Function

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

Public Function isCurrency(ByVal sKeyAscii, strCur As String) As Integer
    Dim i As Integer
    Dim intDot As Integer
    intDot = 0
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isCurrency = 0
    Else
        If sKeyAscii = 46 Then
            For i = 1 To Len(strCur)
                If Mid(strCur, i, 1) = "." Then
                    intDot = intDot + 1
                End If
            Next
            If intDot < 1 Then
                isCurrency = sKeyAscii
            Else
                isCurrency = 0
            End If
        Else
            isCurrency = sKeyAscii
        End If
    End If
End Function

'Function used to change the yes/no value
Public Function changeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeYNValue = "1"
        Case "N": changeYNValue = "0"
        Case "1": changeYNValue = "Y"
        Case "0": changeYNValue = "N"
    End Select
End Function

'Function that return true if the control is numeric
Public Function is_numeric(ByRef sText As String) As Boolean
    If IsNumeric(sText) = False Then
        is_numeric = False
        MsgBox "The field required a numeric input.Please check it!", vbExclamation
    Else
        is_numeric = True
    End If
End Function

Public Function Date_To_MMDDYY(ByVal strDate As String) As String
 Date_To_MMDDYY = Mid$(strDate, 1, 2) & Mid$(strDate, 4, 2) & Mid$(strDate, 7, 2)
End Function

'Function that return the value of a certain field
Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim Rs As New Recordset
    
    Rs.CursorLocation = adUseClient
    Rs.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If Rs.RecordCount > 0 Then getValueAt = Rs.Fields(whichField)
    
    Set Rs = Nothing
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

'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    Dim Rs As New Recordset
    
    Rs.CursorLocation = adUseClient
    Rs.Open "SELECT COUNT(PK) as TCount FROM " & srcTable & srcCondition, CN, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getRecordCount = Format$(Rs![TCount], "#,##0")
    Else
        getRecordCount = Rs![TCount]
    End If
    Set Rs = Nothing
End Function

'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0")
End Function

'Function used to determine if the object has been set
Public Function isObjectSet(srcObject As Object) As Boolean
    On Error GoTo Err
    'I use tag because almost all controls have this
    srcObject.Tag = srcObject.Tag
    isObjectSet = True
    
    Exit Function
Err:
    isObjectSet = False
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String) As Double
    On Error GoTo Err
    Dim Rs As New ADODB.Recordset

    Rs.CursorLocation = adUseClient
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " HAVING(" & sCondition & ")"
    If inclField <> "" Then inclField = "," & inclField
    Rs.Open "SELECT Sum(" & sTable & "." & sField & ") AS fTotal" & inclField & " FROM " & sTable & sCondition, sCN, adOpenStatic, adLockOptimistic
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            getSumOfFields = getSumOfFields + Rs.Fields("fTotal")
            Rs.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If
    
    Set Rs = Nothing
    Exit Function
Err:
        'Error when incounter a null value
        If Err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function

' ExportExcel.bas
'
'*****************************************************
' ExportListbox
' Purpose:   Exports listbox information to a MS Excel
'            spreadsheet.
' Inputs:
'   pListview:      Listview control reference
'   pFilename:      Name of Excel file to create.
'   Append:         Appends an existing spreadsheet
'
' Returns:          True if successful
'                   False otherwise.
'*****************************************************

Public Function ExportListview(ByRef pListview As MSComctlLib.listview, _
    ByVal pFilename As String, _
    Optional ByVal WorksheetName As String = "Sheet1", _
    Optional Append As Boolean = False) As Boolean
    
    Dim CN As Object
    Dim CAT As Object
    Dim TBL As Object
    Dim COL As Object
    Dim strConnection As String
    Dim AListItem As MSComctlLib.ListItem
    Dim AColumnHeader As MSComctlLib.ColumnHeader
    Dim Rs As Object
    Dim intLoop As Integer
    Dim intLoop2 As Integer
    
    On Error GoTo ErrHandler
    
    ' Make sure everything is ok with the inputs before
    ' continuing.
    ' pListView
    If pListview.View <> lvwReport Then
        MsgBox "Listview must be in Report mode.", _
            vbCritical + vbOKOnly, "ExportListview"
        GoTo NotSuccessful
    End If
    ' pFilename
    If Trim$(pFilename) = vbNullString Then
        MsgBox "No filename given.", vbCritical + vbOKOnly, _
            "ExportListview"
        GoTo NotSuccessful
    End If
    ' **********
    Set CN = CreateObject("ADODB.Connection")
    
    ' Create a connection to the Excel file using Jet's ISAM
    strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Extended Properties=Excel 8.0;" & _
        "Data Source=" & pFilename
    CN.Open strConnection
    
    ' No need to create a workbook if the spreadsheet already exists
    If Append Then GoTo AlreadyExists
    
    ' Create a Excel Workbook and set the connection to CN
    Set CAT = CreateObject("ADOX.Catalog")
    CAT.ActiveConnection = CN
    
    ' Create a worksheet for the cat
    Set TBL = CreateObject("ADOX.Table")
    TBL.Name = WorksheetName
    
    ' Do the column headers
    
    For Each AColumnHeader In pListview.ColumnHeaders
        Set COL = CreateObject("ADOX.Column")
        COL.Type = 130  ' adWChar
        COL.Name = AColumnHeader.Text
        TBL.Columns.Append COL
        Set COL = Nothing
    Next AColumnHeader
    
    ' Add this worksheet to the workbook
    CAT.Tables.Append TBL
    
AlreadyExists:
    Set Rs = CreateObject("ADODB.Recordset")
    
    ' open the excel file that was just created as a recordset
    ' so we can add records.
    Rs.Open WorksheetName, CN, 1, 3
    
    ' Grab every listitem out of the listview control
    
    For Each AListItem In pListview.ListItems
        ' Listitem and then all subitems
        Rs.AddNew
        Rs.Fields(0) = AListItem.Text
        ' subitems
        For intLoop = 1 To Rs.Fields.Count - 1
            Rs.Fields(intLoop) = AListItem.SubItems(intLoop)
        Next intLoop
        Rs.Update
    Next AListItem
    
        

    ' Mark as success
    ExportListview = True
    GoTo CloseAndNothing
    
NotSuccessful:
    ExportListview = False
    
    ' clear all objects and exit
CloseAndNothing:
    On Error Resume Next
    Rs.Close
    CN.Close
    Set CAT = Nothing
    Set CN = Nothing
    Set TBL = Nothing
    Set COL = Nothing
    Set AListItem = Nothing
    Set AColumnHeader = Nothing
    Set Rs = Nothing
    Exit Function
    
ErrHandler:
    ' simply raise the error to the client
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    GoTo CloseAndNothing
End Function

Public Sub ExportOneTable(ByVal strExcel As String, ByVal strTable As String)

'EXPORTS TABLE IN ACCESS DATABASE TO EXCEL
'REFERENCE TO DAO IS REQUIRED

Dim strExcelFile As String
Dim strWorksheet As String
Dim strDB As String
Dim objDB As Database

'Change Based on your needs, or use
'as parameters to the sub
strExcelFile = App.Path & "\" & strExcel & ""
strWorksheet = "Data"

 'If excel file already exists, you can delete it here
 If Dir(strExcelFile) <> "" Then Kill strExcelFile

CN.Execute _
  "SELECT * INTO [Excel 8.0;DATABASE=" & strExcelFile & _
   "].[" & strWorksheet & "] FROM " & "[" & strTable & "]"
CN.Close
Set CN = Nothing
MsgBox "Export Success..!", vbInformation

End Sub


Public Sub ImportOneTable(ByVal strExcel As String, ByVal strTable As String)

Dim cnXLS As ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim rsAccess As New ADODB.Recordset
Dim i As Integer
' Open Excel Connection
Set cnXLS = New ADODB.Connection
With cnXLS
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=Supplier.xls;" & _
    "Extended Properties=Excel 8.0;"
    .Open
End With

 ' Load ADO Recordset with Excel Sheet1Data
 
oRs.Open "Select * from [Data]", cnXLS, adOpenStatic
MsgBox oRs.RecordCount
 
 
' Load ADO Recordset with Access Data
rsAccess.Open "select * from " & strTable & "", CN, adOpenStatic, adLockOptimistic
MsgBox rsAccess.RecordCount
CN.Execute "Delete from " & strTable & ""
'Synchronize Recordsets and Batch Update
Do While Not (oRs.EOF)
        rsAccess.AddNew
        For i = 0 To 32  '-----11 columns in table 1
        rsAccess.Fields(i).Value = oRs.Fields(i).Value
        Next
        rsAccess.Update
        oRs.MoveNext
 
Loop
MsgBox "Export Success..!", vbInformation

End Sub

Function AutoID(ByVal sTable As String, ByVal sIDField As String, ByVal sKode As String) As String
'Call AutoID("Barang", "kodebarang", "BR")
Dim X As New ADODB.Recordset
Dim z As String
Dim no As String
If X.State = 1 Then X.Close
X.Open "select " & sIDField & " from " & sTable & " order by " & sIDField & "", CN, adOpenDynamic, adLockOptimistic
If Not X.EOF Then
    X.MoveLast
    z = X(sIDField)
    z = Left(z, Len(sKode))
    If sKode = z Then
        no = Val(Right(X(sIDField), 4)) + 1
        no = z & String(4 - Len(no), "0") & no
    Else
        no = sKode & "0001"
    End If
Else
    no = sKode & "0001"
End If
AutoID = no
X.Close
End Function

'Function used to check if the record exit or not.
Public Sub AddComboOne(ByVal sTable As String, ByVal Field1 As String, ByRef cbo As ComboBox, Optional wField As String, Optional sSTR As String)
    Dim Rs As New Recordset
    Dim sSQL As String
    Rs.CursorLocation = adUseClient
    If sSTR <> "" Then
        sSQL = "Select " & Field1 & " From " & sTable & " Where " & wField & " = '" & sSTR & "'"
        Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    Else
        sSQL = "Select " & Field1 & " From " & sTable & ""
        Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    End If
    If Rs.RecordCount > 0 Then
        cbo.Clear
        Rs.MoveFirst
            cbo.AddItem ""
        Do While Not Rs.EOF
            cbo.AddItem Rs.Fields(Field1)
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub


'Function used to check if the record exit or not.
Public Sub AddComboList(ByVal sTable As String, ByVal Field1 As String, ByVal Field2 As String, ByRef cbo As ComboBox, Optional wField As String, Optional sSTR As String)
    Dim Rs As New Recordset
    Rs.CursorLocation = adUseClient
    If sSTR <> "" Then
        Rs.Open "Select " & Field1 & "," & Field2 & " From " & sTable & " Where " & wField & " = '" & sSTR & "'", CN, adOpenStatic, adLockOptimistic
    Else
        Rs.Open "Select " & Field1 & "," & Field2 & " From " & sTable & "", CN, adOpenStatic, adLockOptimistic
    End If
    If Rs.RecordCount > 0 Then
        cbo.Clear
        Rs.MoveFirst
        Do While Not Rs.EOF
            cbo.AddItem Rs.Fields(Field1) & " - " & Rs.Fields(Field2)
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub


'Function used to check if the record exit or not.
Public Sub AddComboField(ByVal sTable As String, ByVal Field1 As String, ByVal Field2 As String, ByRef cbo As ComboBox, Optional wField As String, Optional sSTR As String)
    Dim Rs As New Recordset
    Rs.CursorLocation = adUseClient
    If sSTR <> "" Then
        Rs.Open "Select " & Field1 & "," & Field2 & " From " & sTable & " Where " & wField & " = '" & sSTR & "'", CN, adOpenStatic, adLockOptimistic
    Else
        Rs.Open "Select " & Field1 & "," & Field2 & " From " & sTable & "", CN, adOpenStatic, adLockOptimistic
    End If
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            cbo.AddItem Rs.Fields(Field1) & " - " & Rs.Fields(Field2)
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub

Public Sub AltLVBackground(lv As listview, _
    ByVal BackColorOne As OLE_COLOR, _
    ByVal BackColorTwo As OLE_COLOR, sForm As Form)
'---------------------------------------------------------------------------------
' Purpose   : Alternates row colors in a ListView control
' Method    : Creates a picture box and draws the desired color scheme in it, then
'             loads the drawn image as the listviews picture.
'---------------------------------------------------------------------------------
Dim lH      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
    With lv
        If .View = lvwReport And .ListItems.Count Then
            Set picAlt = sForm.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = .ListItems(1).Height
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lH * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
                Set lv.Picture = .Image
            End With
            Set picAlt = Nothing
            sForm.Controls.Remove "picAlt"
            lv.Parent.ScaleMode = lSM
        End If
    End With
End Sub



Public Sub AddOneCombo(ByVal sTable As String, ByVal Field1 As String, ByRef cbo As ComboBox, Optional sSTR As String)
    Dim Rs As New Recordset
    Rs.CursorLocation = adUseClient
    If sSTR <> "" Then
        Rs.Open "Select " & Field1 & " From " & sTable & " Where " & Field1 & " Like '%" & sSTR & "%'", CN, adOpenStatic, adLockOptimistic
    Else
        Rs.Open "Select Distinct(" & Field1 & ") From " & sTable & "", CN, adOpenStatic, adLockOptimistic
    End If
    If Rs.RecordCount > 0 Then
        cbo.Clear
        Rs.MoveFirst
        Do While Not Rs.EOF
            cbo.AddItem Rs.Fields(Field1)
            Rs.MoveNext
        Loop
    End If
    Set Rs = Nothing
End Sub

Public Function getIP()

Dim WMI     As Object
Dim qryWMI  As Object
Dim Item    As Variant

    Set WMI = GetObject("winmgmts:\\.\root\cimv2")

    Set qryWMI = WMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration " & _
                               "WHERE IPEnabled = True")

    For Each Item In qryWMI
      getIP = Item.IPAddress(0)
    Next

    Set WMI = Nothing
    Set qryWMI = Nothing

End Function



