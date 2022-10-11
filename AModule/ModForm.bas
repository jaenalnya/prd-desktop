Attribute VB_Name = "ModForm"
Option Explicit
Public RecordPage                     As New clsPaging

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
