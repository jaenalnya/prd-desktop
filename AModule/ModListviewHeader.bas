Attribute VB_Name = "ModListviewHeader"

'Listview Consts
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public RecordPage                     As New clsPaging

Public Sub lvSizeColumns(lv As listview)
Dim Counter As Long
    'Resizes Listview Column Headers.
    For Counter = 0 To (lv.ColumnHeaders.Count - 2)
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Counter, _
        ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next Counter
End Sub


Public Sub SetNavigation(ByRef cmdFirst As CommandButton, ByRef cmdPrev As CommandButton, ByRef cmdNext As CommandButton, ByRef cmdLast As CommandButton)
    With RecordPage
        If .PAGE_TOTAL = 1 Then
            cmdFirst.Enabled = False
            cmdPrev.Enabled = False
            cmdNext.Enabled = False
            cmdLast.Enabled = False
        ElseIf .PAGE_CURRENT = 1 Then
            cmdFirst.Enabled = False
            cmdPrev.Enabled = False
            cmdNext.Enabled = True
            cmdLast.Enabled = True
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            cmdFirst.Enabled = True
            cmdPrev.Enabled = True
            cmdNext.Enabled = False
            cmdLast.Enabled = False
        Else
            cmdFirst.Enabled = True
            cmdPrev.Enabled = True
            cmdNext.Enabled = True
            cmdLast.Enabled = True
        End If
    End With
End Sub

