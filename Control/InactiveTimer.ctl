VERSION 5.00
Begin VB.UserControl InactiveTimer 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   InvisibleAtRuntime=   -1  'True
   Picture         =   "InactiveTimer.ctx":0000
   ScaleHeight     =   750
   ScaleWidth      =   885
   ToolboxBitmap   =   "InactiveTimer.ctx":0C42
   Begin VB.Timer tmrInactive 
      Interval        =   1000
      Left            =   480
      Top             =   360
   End
End
Attribute VB_Name = "InactiveTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Long
Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

' We raise this event if the user is inactive for too long.
Public Event UserInactive()

Private m_InactiveInterval As Long
Private Const m_def_InactiveInterval As Long = 1000& * 5 * 60

' Default to 5 minutes.
Private Sub UserControl_InitProperties()
    m_InactiveInterval = m_def_InactiveInterval
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Me.InactiveInterval = PropBag.ReadProperty("InactiveInterval", m_def_InactiveInterval)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Me.Enabled, "True"
    PropBag.WriteProperty "InactiveInterval", Me.InactiveInterval, m_def_InactiveInterval
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = ScaleX(32, vbPixels, vbTwips)
    UserControl.Height = ScaleX(32, vbPixels, vbTwips)
End Sub

Public Property Get InactiveInterval() As Long
    InactiveInterval = m_InactiveInterval
End Property
Public Property Let InactiveInterval(ByVal Value As Long)
    m_InactiveInterval = Value
    PropertyChanged "InactiveInterval"
End Property

' Delegate the Enabled property to the timer.
Public Property Get Enabled() As Boolean
    Enabled = tmrInactive.Enabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
    tmrInactive.Enabled = Value
End Property

' Return the number of seconds
Private Function ElapsedIdleTime() As Long
Dim m_lii As LASTINPUTINFO

    m_lii.cbSize = Len(m_lii)

    If GetLastInputInfo(m_lii) = 0 Then
        Err.Raise vbObjectError + 1001, "InactiveTimer", _
            "Error getting last input information"
    End If

    ElapsedIdleTime = GetTickCount() - m_lii.dwTime
End Function

' See if the user has been idle for too long.
Private Sub tmrInactive_Timer()
    If Not UserControl.Ambient.UserMode Then Exit Sub

    If ElapsedIdleTime() > m_InactiveInterval Then RaiseEvent UserInactive
End Sub
