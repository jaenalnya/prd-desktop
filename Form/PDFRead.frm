VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Document Viewer"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   15960
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2832
      Top             =   48
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   3552
      TabIndex        =   1
      Top             =   144
      Width           =   684
   End
   Begin VB.Frame Frame1 
      Height          =   5025
      Left            =   1440
      TabIndex        =   0
      Top             =   945
      Width           =   8895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Demonstrated by: Julius Enerio
' Requirements: Make sure the Adobe Acrobat 7.0 Control Type Library (AcroPDF.dll) is visible in your toolbox
' Right Click on the toolbox then select Components. Select Adobe Acrobat 7.0 Control Type Library then click OK
' No need to add the control in the form though. Just use the code.
' Read More on how to use the browser control by reading the Interapplication communication API Reference found by searching in on Google

Option Explicit 'Which means all variables must be declared before it can be used in the program

Private m_objPDF As AcroPDFLibCtl.AcroPDF 'Declare an object of type AcroPDF
Private m_strFilePath As String 'Declare a string for the PDF Filename and Path

Private Sub Command1_Click()
    dlg.Filter = "PDF (*.pdf)|*.pdf"
    dlg.ShowOpen
    m_strFilePath = dlg.FileName
    Form_Activate
End Sub

'On Form Load...
Private Sub Form_Load()
   m_strFilePath = App.Path & "\Test.pdf" 'Change this to the path and filename of your PDF File
   Set m_objPDF = Controls.Add("AcroPDF.PDF.1", "x") 'This will add the PDF Browser control to the form on runtime. The "Test" is the control's name
   Set m_objPDF.Container = Frame1 'Attach the PDF Browser control to a container.
   'A Container can be a Frame, PictureBox, or SSTab Control. In this code, I used a Frame.
   
End Sub

'On Form Activate
Private Sub Form_Activate()
   'Load the PDF file specified in m_strFilePath.
   'Make sure to do this before doing any changes to the browser controls view/layout
   m_objPDF.LoadFile m_strFilePath
   
   'Set whether a toolbar will appear in the viewer. True to show, False to Hide.
   m_objPDF.setShowToolbar False
   
   'Sets the Layout Mode for a page view according to the specified string.
   'DontCare — use the current user preference
   'SinglePage — use single page mode (as it would have appeared in pre-Acrobat 3.0 viewers)
   'OneColumn — use one-column continuous mode
   'TwoColumnLeft — use two-column continuous mode with the first page on the left
   'TwoColumnRight — use two-column continuous mode with the first page on the right
   m_objPDF.setLayoutMode "SinglePage"
   
   'Sets the page mode in which a document is to be opened
   'PDDontCare: 0 — leave the view mode as it is
   'PDUseNone: 1 — display without bookmarks or thumbnails
   'PDUseThumbs: 2 — display using thumbnails
   'PDUseBookmarks: 3 — display using bookmarks
   m_objPDF.setPageMode "none"
   
   'Set the Zoom view according to the value specified. ranges from 0 and onwards
   m_objPDF.setZoom 90
   
   'Move and Resize the object in relation to its container/form
   With m_objPDF
      '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
      .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
   End With
   
   'Show the Browser Control
   m_objPDF.Visible = True
End Sub

Private Sub Form_Resize()
   'Move and Resize the object in relation to its container/form
   Frame1.Width = Me.Width - 500
   Frame1.Height = Me.Height - 1000
   With m_objPDF
      .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
   End With
End Sub

'On Form Unload
Private Sub Form_Unload(Cancel As Integer)
 
 'The Browser control will load a blank
 m_objPDF.LoadFile ""
 
 'Set object to nothing
 Set m_objPDF = Nothing
End Sub
