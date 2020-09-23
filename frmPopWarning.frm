VERSION 5.00
Begin VB.Form frmPopWarning 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNotes 
      BackColor       =   &H00FFFFFF&
      Height          =   1060
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   3540
      TabIndex        =   3
      Top             =   350
      Width           =   3600
      Begin VB.ListBox lstNotes 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Height          =   1005
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3540
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   3540
      TabIndex        =   0
      Top             =   0
      Width           =   3570
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   3495
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3540
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3480
      X2              =   3720
      Y1              =   1200
      Y2              =   1560
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   3600
      Picture         =   "frmPopWarning.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Close Warning"
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmPopWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, TopMost

Private Sub Form_Load()
Me.Left = Screen.Width - frmMain.Width / 2 - Me.Width
Me.Top = frmMain.lstNote.Top - Me.Height + 200
lblTitle.Caption = "Today, you have " & frmMain.lstNote.ListCount & " note(s)."
For i = 0 To frmMain.lstNote.ListCount - 1
    lstNotes.AddItem frmMain.lstNote.List(i)
Next i
MakeTransparent Me
TopMost = SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub lstNotes_DblClick()
On Error Resume Next
Dim tmpFilePath As String
Dim OldAHVal As Boolean

OldAHVal = AutohideMenu
If AutohideMenu = True Then frmOptions.Check3.Value = 0
ViewNote = True
tmpFilePath = App.Path & "\SaveNote\" & Trim$(Mid$(lstNotes.List(lstNotes.ListIndex), 3, Len(lstNotes.List(lstNotes.ListIndex))))
frmMain.rtbNote.LoadFile tmpFilePath
If OldAHVal = True Then
    AutohideMenu = True
    Me.Visible = False
    pause (9)
    frmOptions.Check3.Value = 1
End If
Unload Me
End Sub
