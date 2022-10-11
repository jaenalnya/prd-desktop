Attribute VB_Name = "ModProcedures"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public p_shift                         As Date

Public Sub GetShift()
    If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
        p_shift = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd hh:mm:ss")), "yyyy-mm-dd hh:mm:ss")
    Else
        p_shift = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
End Sub

Public Sub LoadForm(ByRef srcForm As Form)
    Dim R As RECT
    
    MAIN.b8CW.LoadChildWindow MAIN.hwnd, srcForm.hwnd, srcForm.Name, srcForm.Caption, R.Top, R.Left, R.Right, R.Bottom
    
    srcForm.Show
    srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub

Public Sub ActivateMDIChildForm(ByVal sFormName As String)
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = sFormName Then
            'activate form
            'ResizeMdiChildForm frm
            frm.Visible = True
            frm.Show
            frm.SetFocus
            'set tab active window
            MAIN.b8CW.SetActiveWindow sFormName
            Exit For
        End If
    Next
    Set frm = Nothing
End Sub

'Procedure used to promp unexpected errors
Public Sub Prompt_Err(ByVal sError As ErrObject, ByVal ModuleName As String, ByVal OccurIn As String)
    MsgBox "Error From: " & ModuleName & vbNewLine & _
           "Occur In: " & OccurIn & vbNewLine & _
           "Error Number: " & sError.Number & vbNewLine & _
           "Description: " & sError.Description, vbCritical, "ApplicatiOn Error"
    'Save the error log (The save error log will be display later on in the program)
    Open App.Path & "\Error.log" For Append As #1
        Print #1, Format(Date, "MMM-dd-yyyy") & "~~~~~" & Time & "~~~~~" & sError.Number & "~~~~~" & sError.Description & "~~~~~" & ModuleName & "~~~~~" & OccurIn
    Close #1
End Sub

'Procedure used to delete record with SQL
Public Sub DelRecwSQL(ByVal sTable As String, ByVal sField As String, ByVal sString As String, ByVal isNumber As Boolean, ByVal snum As Long)
    If isNumber = True Then
        CN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " =" & snum
    Else
        CN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " ='" & sString & "'"
    End If
End Sub

'Procedure used to highlight text when focus
Public Sub HLText(ByRef srcText)
    On Error Resume Next
    With srcText
        .SelStart = 0
        .SelLength = Len(srcText.Text)
        .BackColor = &HC0FFFF
    End With
    srcText = UCase(srcText)
End Sub

Public Sub unHLText(ByRef srcText1)
On Error Resume Next
    With srcText1
        .BackColor = &HFFFFFF
    End With
    srcText1 = UCase(srcText1)
End Sub

Public Sub unHLTextTran(ByRef srcText1)
On Error Resume Next
    With srcText1
        .BackColor = &HFFFFC0
    End With
    srcText1 = UCase(srcText1)
End Sub

'Procedure used to clear the text content
Public Sub clearText(ByRef sForm As Form)
    Dim Control As Control
    For Each Control In sForm.Controls
        If (TypeOf Control Is TextBox) Then Control = vbNullString
    Next Control
    Set Control = Nothing
End Sub


'Procedure that will change the value at once
Public Sub ChangeValue(ByRef srcCN As Connection, ByVal srcTable As String, ByVal srcField As String, ByVal srcValue As String, Optional isNumber As Boolean, Optional srcCondition As String)
    If srcCondition <> vbNullString Then srcCondition = " " & srcCondition
    If isNumber = True Then
        srcCN.Execute "UPDATE " & srcTable & " SET " & srcField & " =" & srcValue & " " & srcCondition
    Else
        srcCN.Execute "UPDATE " & srcTable & " SET " & srcField & " ='" & srcValue & "'" & " " & srcCondition
    End If
End Sub

'Procedure used to center form
Public Sub CenterForm(ByRef srcForm1 As Form)
On Error Resume Next
    With srcForm1
    .Move (Screen.Width - srcForm1.Width) / 2, (Screen.Height - srcForm1.Height) / 2
    End With
End Sub
'Procedure used to center object horizontal
Public Sub center_obj_horizontal(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Left = (sParentObj - sMoveObj.Width) / 2
End Sub
'Procedure used to center vertical
Public Sub center_obj_vertical(ByVal sParentObj As Variant, ByRef sMoveObj As Variant)
    sMoveObj.Top = (sParentObj.Height - sMoveObj.Height) / 2
End Sub

Public Function AllowOnlyNumbers(KeyAscii As Integer, obj As Control) As Integer
  If ((KeyAscii <> 8) And (KeyAscii <> vbKeyDelete) And _
  (KeyAscii <> 46)) And ((KeyAscii < 48 Or KeyAscii > 57)) Then
    AllowOnlyNumbers = 0
  Else
    If KeyAscii = 46 Then
      If InStr(obj.Text, ".") Then
        KeyAscii = 0
        Exit Function
      End If
    End If
    AllowOnlyNumbers = KeyAscii
  End If
End Function

'Procedure used to search in listview
Public Sub search_in_listview(ByRef sListView As listview, ByVal sFindText As String)
    Dim tmp_listtview As ListItem
    Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
End Sub

Public Sub sSQL_Insert(ByVal strSQL As String)
Set COMMAND_INSERT = New ADODB.Command
    With COMMAND_INSERT
        .ActiveConnection = CN
        .CommandText = strSQL
        .Execute
    End With
End Sub

Public Sub sSQL_Update(ByVal strSQL As String)
Set COMMAND_UPDATE = New ADODB.Command
    With COMMAND_UPDATE
        .ActiveConnection = CN
        .CommandText = strSQL
        .Execute
    End With
End Sub

Public Sub sSQL_Delete(ByVal strSQL As String)
Set COMMAND_DELETE = New ADODB.Command
    With COMMAND_DELETE
        .ActiveConnection = CN
        .CommandText = strSQL
        .Execute
    End With
End Sub

Public Sub sSQL_Insert_SIP(ByVal strSQL As String)
Set COMMAND_INSERT = New ADODB.Command
    With COMMAND_INSERT
        .ActiveConnection = CN_SIP
        .CommandText = strSQL
        .Execute
    End With
End Sub

Public Sub sSQL_Update_SIP(ByVal strSQL As String)
Set COMMAND_UPDATE = New ADODB.Command
    With COMMAND_UPDATE
        .ActiveConnection = CN_SIP
        .CommandText = strSQL
        .Execute
    End With
End Sub

Public Sub sSQL_Delete_SIP(ByVal strSQL As String)
Set COMMAND_DELETE = New ADODB.Command
    With COMMAND_DELETE
        .ActiveConnection = CN_SIP
        .CommandText = strSQL
        .Execute
    End With
End Sub


Public Sub ButtonList(ByRef ListV As listview, ByRef ButF As CommandButton, _
                      ByRef ButP As CommandButton, ByRef ButN As CommandButton, ByRef ButL As CommandButton)
With MAIN
    'For listview
    Set ListV.SmallIcons = .i16x16
    Set ListV.Icons = .i16x16

    ButF.Picture = .i16x16.ListImages(3).Picture
    ButP.Picture = .i16x16.ListImages(4).Picture
    ButN.Picture = .i16x16.ListImages(5).Picture
    ButL.Picture = .i16x16.ListImages(6).Picture
    
    ButF.DisabledPicture = .i16x16g.ListImages(3).Picture
    ButP.DisabledPicture = .i16x16g.ListImages(4).Picture
    ButN.DisabledPicture = .i16x16g.ListImages(5).Picture
    ButL.DisabledPicture = .i16x16g.ListImages(6).Picture
End With

End Sub

