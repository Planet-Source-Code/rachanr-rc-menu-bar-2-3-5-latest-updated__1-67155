VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "RC Menu Bar"
   ClientHeight    =   8685
   ClientLeft      =   12870
   ClientTop       =   405
   ClientWidth     =   1935
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   1
      Left            =   1320
      Top             =   4800
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   1
      Left            =   720
      Top             =   4800
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   1
      Left            =   120
      Top             =   4800
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   1
      Left            =   1320
      Top             =   4200
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   1
      Left            =   720
      Top             =   4200
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   1
      Left            =   120
      Top             =   4200
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   1
      Left            =   1320
      Top             =   3600
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   1
      Left            =   720
      Top             =   3600
   End
   Begin VB.Timer Tapp 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   1
      Left            =   120
      Top             =   3600
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   480
      Top             =   8160
   End
   Begin VB.CheckBox chkRefreshNoteList 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.FileListBox NoteFiles 
      Height          =   285
      Left            =   960
      Pattern         =   "*.rcn"
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox txt2Search 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Text            =   "Text to search"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Year"
      Top             =   5640
      Width           =   495
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "Month"
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox cboDay 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Day"
      Top             =   5640
      Width           =   550
   End
   Begin VB.PictureBox picBuffer 
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   1080
   End
   Begin VB.PictureBox picMeter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1560
      ScaleHeight     =   1065
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   840
      Width           =   375
      Begin VB.Label lblMemVal 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   340
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1440
      Top             =   360
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdClearNote 
      Caption         =   "&Clear Note"
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdSaveNote 
      Caption         =   "&Save Note"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7560
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdbSave 
      Left            =   360
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save RC Note File"
   End
   Begin VB.ListBox lstNote 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "frmMain.frx":4E12
      Left            =   0
      List            =   "frmMain.frx":4E14
      TabIndex        =   7
      Top             =   6480
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtbNote 
      Height          =   2775
      Left            =   0
      TabIndex        =   6
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":4E16
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   8520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   8280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   8520
   End
   Begin MSComctlLib.ImageList imgIcoList 
      Left            =   720
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIconList 
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select date for save note:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5450
      Width           =   1830
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   120
      X2              =   1920
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   960
      Y2              =   6120
   End
   Begin VB.Image imgTmpApp 
      Height          =   495
      Left            =   1680
      Picture         =   "frmMain.frx":6FE2
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   11
      Left            =   1320
      MouseIcon       =   "frmMain.frx":72EC
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   10
      Left            =   720
      MouseIcon       =   "frmMain.frx":743E
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   9
      Left            =   120
      MouseIcon       =   "frmMain.frx":7590
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   8
      Left            =   1320
      MouseIcon       =   "frmMain.frx":76E2
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   7
      Left            =   720
      MouseIcon       =   "frmMain.frx":7834
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   6
      Left            =   120
      MouseIcon       =   "frmMain.frx":7986
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   5
      Left            =   1320
      MouseIcon       =   "frmMain.frx":7AD8
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   4
      Left            =   720
      MouseIcon       =   "frmMain.frx":7C2A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   3
      Left            =   120
      MouseIcon       =   "frmMain.frx":7D7C
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   2
      Left            =   1320
      MouseIcon       =   "frmMain.frx":7ECE
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":8020
      Stretch         =   -1  'True
      ToolTipText     =   "Control Panel"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   1
      Left            =   720
      MouseIcon       =   "frmMain.frx":F512
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":F664
      Stretch         =   -1  'True
      ToolTipText     =   "My Documents"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgApp 
      Height          =   495
      Index           =   0
      Left            =   120
      MouseIcon       =   "frmMain.frx":14046
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":14198
      Stretch         =   -1  'True
      ToolTipText     =   "My Computer"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   1800
      Picture         =   "frmMain.frx":15E92
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblFreeMem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Free Memory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblTotalMem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Memory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   2040
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   1920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   2040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   2040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image imgInfo 
      Height          =   255
      Left            =   1560
      MouseIcon       =   "frmMain.frx":17FBC
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":1810E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1200
      MouseIcon       =   "frmMain.frx":18550
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5880
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " NOTE:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label lblNow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hh:mm:ss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblToday 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dd mm yyyy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmMain.frx":186A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000002&
      Caption         =   "     RC Menu Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   9
      Left            =   60
      TabIndex        =   32
      Top             =   4740
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   8
      Left            =   1260
      TabIndex        =   31
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   7
      Left            =   660
      TabIndex        =   30
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   6
      Left            =   60
      TabIndex        =   29
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   5
      Left            =   1260
      TabIndex        =   28
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   4
      Left            =   660
      TabIndex        =   27
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   3
      Left            =   60
      TabIndex        =   26
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   2
      Left            =   1260
      TabIndex        =   25
      Top             =   2940
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   11
      Left            =   1260
      TabIndex        =   34
      Top             =   4740
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   10
      Left            =   660
      TabIndex        =   33
      Top             =   4740
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   1
      Left            =   660
      TabIndex        =   24
      Top             =   2940
      Width           =   615
   End
   Begin VB.Label lblBG 
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   23
      Top             =   2940
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScrHeight As Integer, ScrWidth As Integer, ViewNote As Boolean
Dim tmpInd As Integer, MY_DOCUMENTS As String, fNum As Integer
Dim tmpAppIco(3 To 11) As String, tmpAppName(3 To 11) As String
Dim AppName(3 To 11) As String, LocalLang As String, tmpNoteOptions As Boolean
Dim myID As ITEMIDLIST, i As Integer, nMonth As Integer, rcNotePath As String
Dim tmpDay As Integer, tmpMonth As Integer, tmpYear As Integer
Dim strSaveName As String, strSavePath As String, strSaveFile As String
Dim SaveState As Boolean, tmpAutoHide As Integer, tmpShowAllNote As Integer
Dim HeadDate As String, HdLength As Integer, ScrRes As String
Dim tmpIndex1 As Integer, tmpIndex2 As Integer, tmpNotePath As String
Dim tmpFile(1000) As String, tmpText As String, Button As Integer
Dim tmpTranVal As Integer, tmpTranEnabled As Integer, RunAppName(3 To 11) As String

Private Sub cboDay_Click()
If frmOptions.Check4.Value = 1 Then
    Call LoadList
Else
    If UCase$(LocalLang) = "THAI" Then
        Call LoadDateList(cboDay.Text, Str(nMonth), Str(Val(cboYear.Text) - 543))
    Else
        Call LoadDateList(cboDay.Text, Str(nMonth), cboYear.Text)
    End If
End If
End Sub

Private Sub cboMonth_Click()
If UCase$(LocalLang) = "THAI" Then
    If cboMonth.Text = "Á¡ÃÒ¤Á" Then nMonth = 1
    If cboMonth.Text = "¡ØÁÀÒ¾Ñ¹¸ì" Then nMonth = 2
    If cboMonth.Text = "ÁÕ¹Ò¤Á" Then nMonth = 3
    If cboMonth.Text = "àÁÉÒÂ¹" Then nMonth = 4
    If cboMonth.Text = "¾ÄÉÀÒ¤Á" Then nMonth = 5
    If cboMonth.Text = "ÁÔ¶Ø¹ÒÂ¹" Then nMonth = 6
    If cboMonth.Text = "¡Ã¡®Ò¤Á" Then nMonth = 7
    If cboMonth.Text = "ÊÔ§ËÒ¤Á" Then nMonth = 8
    If cboMonth.Text = "¡Ñ¹ÂÒÂ¹" Then nMonth = 9
    If cboMonth.Text = "µØÅÒ¤Á" Then nMonth = 10
    If cboMonth.Text = "¾ÄÈ¨Ô¡ÒÂ¹" Then nMonth = 11
    If cboMonth.Text = "¸Ñ¹ÇÒ¤Á" Then nMonth = 12
Else
    If cboMonth.Text = "January" Then nMonth = 1
    If cboMonth.Text = "February" Then nMonth = 2
    If cboMonth.Text = "March" Then nMonth = 3
    If cboMonth.Text = "April" Then nMonth = 4
    If cboMonth.Text = "May" Then nMonth = 5
    If cboMonth.Text = "June" Then nMonth = 6
    If cboMonth.Text = "July" Then nMonth = 7
    If cboMonth.Text = "August" Then nMonth = 8
    If cboMonth.Text = "September" Then nMonth = 9
    If cboMonth.Text = "October" Then nMonth = 10
    If cboMonth.Text = "November" Then nMonth = 11
    If cboMonth.Text = "December" Then nMonth = 12
End If
Select Case nMonth
Case Is = 1, 3, 5, 7, 8, 10, 12
    cboDay.Clear
    For i = 1 To 31
        cboDay.AddItem i
    Next i
Case Is = 4, 6, 9, 11
    cboDay.Clear
    For i = 1 To 30
        cboDay.AddItem i
    Next i
Case Is = 2
    cboDay.Clear
    If ((Year(Date) - 2004) Mod 4) = 0 Then
        For i = 1 To 29
            cboDay.AddItem i
        Next i
    Else
        For i = 1 To 28
            cboDay.AddItem i
        Next i
    End If
End Select
cboDay.Text = Day(Date)
cboMonth.toolTipText = cboMonth.Text
If frmOptions.Check4.Value = 1 Then
    Call LoadList
Else
    If UCase$(LocalLang) = "THAI" Then
        Call LoadDateList(cboDay.Text, Str(nMonth), Str(Val(cboYear.Text) - 543))
    Else
        Call LoadDateList(cboDay.Text, Str(nMonth), cboYear.Text)
    End If
End If
End Sub

Private Sub cboYear_Click()
If frmOptions.Check4.Value = 1 Then
    Call LoadList
Else
    If UCase$(LocalLang) = "THAI" Then
        Call LoadDateList(cboDay.Text, Str(nMonth), Str(Val(cboYear.Text) - 543))
    Else
        Call LoadDateList(cboDay.Text, Str(nMonth), cboYear.Text)
    End If
End If
End Sub

Private Sub chkRefreshNoteList_Click()
If frmOptions.Check4.Value = 1 Then Call LoadList Else Call LoadDateList(Str$(Day(Date)), Str$(Month(Date)), Str$(Year(Date)))
End Sub

Private Sub cmdClearNote_Click()
Dim Ans
If ViewNote = False Then
    If SaveState = False Then
        Ans = MsgBox("You are not save your note. Would you like to save your note?" _
            , vbQuestion + vbYesNo, "Save Question")
        If Ans = vbYes Then Call cmdSaveNote_Click
    End If
End If
rtbNote.Text = ""
lstNote.Refresh
rtbNote.SetFocus
SaveState = True
End Sub

Private Sub cmdSaveNote_Click()
On Error Resume Next
strSaveName = Trim$(cboDay.Text) & "_" & Trim$(Str(nMonth)) & "_" & Trim$(cboYear.Text) & "_[" & Trim$(Str(Hour(Time))) & "." & Trim$(Str(Minute(Time))) & "].rcn"
strSavePath = rcNotePath
cdbSave.FLAGS = cdlOFNOverwritePrompt
cdbSave.InitDir = strSavePath
cdbSave.filename = Trim$(strSaveName)
cdbSave.Filter = "RC Note File(s) (*.rcn)|*.rcn"
cdbSave.ShowSave
rtbNote.SaveFile cdbSave.filename
rtbNote.Text = ""
cboDay.Clear
cboMonth.Clear
cboYear.Clear
Call SetDate(Day(Date), Month(Date), Year(Date), GetLanguage)
lstNote.Refresh
If frmOptions.Check4.Value = 1 Then Call LoadList Else Call LoadDateList(Str$(Day(Date)), Str$(Month(Date)), Str$(Year(Date)))
SaveState = True
End Sub

Private Sub cmdSearch_Click()
Me.Enabled = False
If AddedItem = False Then
    AddedItem = True
    txt2Search.AddItem txt2Search.Text
    frmSearch.cboTextSearch.AddItem txt2Search.Text
    If SearchVisibled = False Then
        frmSearch.Visible = True
        SearchVisibled = True
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SetScrRes
LocalLang = GetLanguage
tmpDay = Day(Date)
tmpMonth = Month(Date)
tmpYear = Year(Date)
Call SetDate(tmpDay, tmpMonth, tmpYear, LocalLang)
Call SetPos
MY_DOCUMENTS = GetSFolder(frmMain, MYDOCS, myID)
For tmpInd = 3 To 11
    tmpAppIco(tmpInd) = GetSetting(App.ProductName, "Saved", "AppIcon" & tmpInd)
    tmpAppName(tmpInd) = GetSetting(App.ProductName, "Saved", "AppName" & tmpInd)
    If tmpAppIco(tmpInd) = "" Then
        imgApp(tmpInd).Picture = imgTmpApp.Picture
        imgApp(tmpInd).toolTipText = "No Application!"
        AppName(tmpInd) = ""
        frmOptions.imgAppIcon(tmpInd).Picture = imgTmpApp.Picture
        frmOptions.lblTmpAppPath(tmpInd).Caption = ""
        frmOptions.imgAppIcon(tmpInd).toolTipText = "Disable!"
        frmOptions.lblAppName(tmpInd).Caption = "No Application!"
    Else
        Call FileInfo(tmpAppIco(tmpInd), tmpInd)
        AppName(tmpInd) = tmpAppIco(tmpInd)
        Call FileInfoOpt(tmpAppIco(tmpInd), tmpInd)
        frmOptions.lblTmpAppPath(tmpInd).Caption = tmpAppIco(tmpInd)
        frmOptions.imgAppIcon(tmpInd).toolTipText = frmOptions.lblAppName(tmpInd).Caption
        frmOptions.lblAppName(tmpInd).Caption = tmpAppName(tmpInd)
        imgApp(tmpInd).toolTipText = frmOptions.lblAppName(tmpInd).Caption
    End If
Next tmpInd
StatusBar1.toolTipText = StatusBar1.Panels(1).Text
tmpAutoHide = GetSetting(App.ProductName, "Saved", "Check3")
If tmpAutoHide = 0 Then Timer3.Enabled = False Else Timer3.Enabled = True
AutohideMenu = GetSetting(App.ProductName, "Saved", "Autohidemenu")
tmpNoteOptions = GetSetting(App.ProductName, "Saved", "Option1")
If tmpNoteOptions = False Then
    tmpNoteOptions = GetSetting(App.ProductName, "Saved", "Option2")
    If tmpNoteOptions = False Then
        frmOptions.Option1.Value = True
        frmOptions.txtNotePath.Text = App.Path & "\SaveNote\"
        frmOptions.txtNotePath.Enabled = False
        frmOptions.cmdBrowse(0).Enabled = False
    Else
        frmOptions.Option2.Value = True
        tmpNotePath = GetSetting(App.ProductName, "Saved", "SaveNotePath")
        frmOptions.txtNotePath.Text = tmpNotePath
        frmOptions.txtNotePath.Enabled = True
        frmOptions.cmdBrowse(0).Enabled = True
    End If
Else
    frmOptions.Option1.Value = True
    frmOptions.txtNotePath.Text = App.Path & "\SaveNote\"
End If
rcNotePath = frmOptions.txtNotePath.Text
tmpShowAllNote = GetSetting(App.ProductName, "Saved", "Check4")
If tmpShowAllNote = 1 Then Call LoadList Else Call LoadDateList(Str$(Day(Date)), Str$(Month(Date)), Str$(Year(Date)))
tmpTranEnabled = GetSetting(App.ProductName, "Saved", "Check5")
tmpTranVal = GetSetting(App.ProductName, "Saved", "TranVal")
If tmpTranEnabled = 1 Then
    frmOptions.Check5.Value = 1
    frmOptions.sldTranVal.Value = tmpTranVal
    Transparent Me, (255 - 2.55 * tmpTranVal)
End If
txt2Search.SelLength = Len(txt2Search.Text)
txt2Search.SetFocus

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mmInd As Integer
For mmInd = 0 To 11
    lblBG(mmInd).BackColor = &H8000000F
Next mmInd
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub imgApp_Click(Index As Integer)
If imgApp(Index).Picture <> imgTmpApp.Picture Then
    If imgApp(Index).BorderStyle = 0 Then
        If Index > 2 Then
            imgApp(Index).BorderStyle = 1
            Tapp(Index).Enabled = True
            RunAppName(Index) = tmpAppName(Index)
        End If
        Select Case Index
        Case Is = 0: SHELL MYCOMPUTER, vbNormalFocus
        Case Is = 1: ShellStart MY_DOCUMENTS
        Case Is = 2: SHELL "rundll32 shell32,Control_RunDLL", vbNormalFocus
        Case Else: SHELL frmOptions.lblTmpAppPath(Index).Caption, vbNormalFocus
        End Select
    Else
        If Index > 2 Then Tapp(Index).Enabled = False
        imgApp(Index).BorderStyle = 0
        If Index > 2 Then KillProc tmpAppName(Index)
    End If
Else
    lblOptions_Click
End If
End Sub

Private Sub imgApp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mmInd As Integer
For mmInd = 0 To 11
    lblBG(mmInd).BackColor = &H8000000F
Next mmInd
lblBG(Index).BackColor = &HFFC0FF   '&HC0FFFF
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgInfo_Click()
frmAbout.Show
End Sub

Private Sub lblOptions_Click()
On Error Resume Next
For tmpInd = 3 To 11
    tmpAppIco(tmpInd) = GetSetting(App.ProductName, "Saved", "AppIcon" & tmpInd)
    tmpAppName(tmpInd) = GetSetting(App.ProductName, "Saved", "AppName" & tmpInd)
    If tmpAppIco(tmpInd) = "" Then
        frmOptions.lblTmpAppPath(tmpInd).Caption = ""
        frmOptions.imgAppIcon(tmpInd).Picture = imgTmpApp.Picture
        frmOptions.imgAppIcon(tmpInd).toolTipText = "Disable!"
        frmOptions.lblAppName(tmpInd).Caption = "No Application!"
    Else
        Call FileInfoOpt(tmpAppIco(tmpInd), tmpInd)
        frmOptions.lblTmpAppPath(tmpInd).Caption = tmpAppIco(tmpInd)
        frmOptions.imgAppIcon(tmpInd).toolTipText = frmOptions.lblAppName(tmpInd).Caption
        frmOptions.lblAppName(tmpInd).Caption = tmpAppName(tmpInd)
    End If
Next tmpInd
frmOptions.Show
Me.Enabled = False
End Sub

Private Sub lstNote_Click()
On Error GoTo ErrHnd
Dim tmpFilePath As String
ViewNote = True
tmpFilePath = NoteFiles.Path & "\" & Trim$(Mid$(lstNote.List(lstNote.ListIndex), 3, Len(lstNote.List(lstNote.ListIndex))))
If lstNote.Text <> "Zero note file(s)." Then rtbNote.LoadFile tmpFilePath
lstNote.toolTipText = Trim$(Mid$(lstNote.List(lstNote.ListIndex), 3, Len(lstNote.List(lstNote.ListIndex))))
Exit Sub
ErrHnd:
If Button = 1 Then
    MsgBox "Invalid selected file!", vbInformation + vbOKOnly, "Error read file"
    lstNote.Clear
    Call LoadList
End If
Exit Sub
End Sub

Private Sub rtbNote_KeyPress(KeyAscii As Integer)
ViewNote = False
SaveState = False
End Sub

Private Sub Tapp_Timer(Index As Integer)
Call IsRun(tmpAppName(Index), Index)
End Sub

Private Sub Timer1_Timer()
lblToday.Caption = Format(Date, "Long Date")
lblNow.Caption = Format(Time, "Long Time")
End Sub

Private Sub SetPos()
ScrHeight = Screen.Height
ScrWidth = Screen.Width
Me.Width = ScrWidth / 7
Me.Height = ScrHeight - 450
Me.Left = ScrWidth - Me.Width
Me.Top = 0
lblNow.Width = Me.Width
lblToday.Width = Me.Width
StatusBar1.Panels(1).Text = Winsock1.LocalIP & ":" & Winsock1.LocalHostName
StatusBar1.Panels(1).Width = Me.Width
cmdSaveNote.Top = StatusBar1.Top - cmdSaveNote.Height - 50
cmdSaveNote.Left = Me.Width / 21
cmdSaveNote.Width = 9 * Me.Width / 21
cmdClearNote.Top = cmdSaveNote.Top
cmdClearNote.Left = 11 * Me.Width / 21
cmdClearNote.Width = 9 * Me.Width / 21
lstNote.Top = cmdSaveNote.Top - lstNote.Height - 50
rtbNote.Top = lstNote.Top - rtbNote.Height - 50
Label1.Top = rtbNote.Top - Label1.Height - 50
Label1.Left = 0
lblOptions.Top = rtbNote.Top - lblOptions.Height - 50
lblOptions.Left = Me.Width - lblOptions.Width - 100
imgInfo.Top = 0
imgInfo.Left = Me.Width - imgClose.Width - imgInfo.Width
imgClose.Top = 0
imgClose.Left = Me.Width - imgClose.Width
rtbNote.Width = Me.Width
lstNote.Width = Me.Width
Line1.X1 = 0
Line1.Y1 = lblNow.Top + lblNow.Height + 50
Line1.X2 = Me.Width
Line1.Y2 = Line1.Y1
lblTotalMem.Top = Line1.Y1 + 50
lblFreeMem.Top = lblTotalMem.Top + lblTotalMem.Height + 10
Line2.X1 = 0
Line2.Y1 = Line1.Y1 + lblTotalMem.Height + lblFreeMem.Height + 100
Line2.X2 = Me.Width
Line2.Y2 = Line2.Y1
txt2Search.Width = Me.Width
txt2Search.Top = Line2.Y1 + 50
cmdSearch.Left = Me.Width / 21
cmdSearch.Top = txt2Search.Top + txt2Search.Height + 50
cmdSearch.Width = 19 * Me.Width / 21
Line3.X1 = 0
Line3.Y1 = cmdSearch.Top + cmdSearch.Height + 50
Line3.X2 = Me.Width
Line3.Y2 = Line3.Y1
Line4.X1 = 0
Line4.Y1 = Label1.Top - 100
Line4.X2 = Me.Width
Line4.Y2 = Line4.Y1
Line5.Y1 = 0
Line5.Y2 = ScrHeight
Line6.X1 = 0
Line6.X2 = imgApp(9).Top + imgApp(9).Height + 50
picMeter.Height = Line2.Y1 - Line1.Y1
picMeter.Top = Line1.Y1
cboDay.Top = Line4.Y1 - cboDay.Height - 50
cboMonth.Top = cboDay.Top
cboYear.Top = cboDay.Top
cboDay.Left = 50
cboDay.Width = 5 * Me.Width / 21
cboMonth.Left = cboDay.Left + cboDay.Width
cboMonth.Width = 9 * Me.Width / 21
cboYear.Left = cboMonth.Left + cboMonth.Width
cboYear.Width = 7 * Me.Width / 21
End Sub

Private Sub Timer2_Timer()
Dim FreeMem As Long, TotalMem As Long, PercentMem
Dim valPMem As Single

Call GlobalMemoryStatus(memInfo)
TotalMem = memInfo.dwTotalPhys
FreeMem = memInfo.dwAvailPhys
PercentMem = Format(FreeMem / TotalMem, "0.00%")
lblTotalMem.Caption = TotalMem / 1024 & " KB"
lblFreeMem.Caption = FreeMem / 1024 & " KB" & " (" & PercentMem & ")"

lblMemVal.Top = picMeter.Height * (1 - (FreeMem / TotalMem))
valPMem = Val(Left$(PercentMem, Len(PercentMem) - 1))
If valPMem >= 25# Then
    lblMemVal.BackColor = &HFF00&
    lblFreeMem.ForeColor = &H80000012
ElseIf valPMem < 25# And valPMem > 10# Then
    lblMemVal.BackColor = &HFFFF&
    lblFreeMem.ForeColor = &H80FF&
ElseIf valPMem <= 10# Then
    lblMemVal.BackColor = &HFF&
    lblFreeMem.ForeColor = &HFF&
End If
End Sub

Private Sub Timer3_Timer()
    Dim Pnt As POINTAPI
    Dim showme As Boolean
    GetCursorPos Pnt
      
    If Pnt.X >= Screen.Width / 15 - 5 Then
        frmMain.Left = Screen.Width - frmMain.Width
    End If
    
    If Pnt.X >= Me.Left / 15 And Pnt.X <= (Me.Left + Me.ScaleWidth) / 15 + 5 And _
    Pnt.Y >= Me.Top / 15 And Pnt.Y <= (Me.Top + Me.ScaleHeight) / 15 + 25 Then
        showme = True
    Else
        showme = False
    End If
    
    If showme = True Then
        frmMain.Left = Screen.Width - frmMain.Width
    Else
        frmMain.Left = Screen.Width
    End If
End Sub

Private Sub FileInfo(fPath As String, Index As Integer)
On Error Resume Next
Dim icoNum As Integer
icoNum = GetIconFile(fPath, imgIconList, picBuffer, 32)
If icoNum = 0 Then
   imgApp(Index).Picture = imgIcoList.ListImages(1).Picture
Else
    imgApp(Index).Picture = imgIconList.ListImages(icoNum).Picture
End If
End Sub

Private Sub FileInfoOpt(fPath As String, Index As Integer)
On Error Resume Next
Dim icoNum As Integer
icoNum = GetIconFile(fPath, imgIconList, picBuffer, 32)
If icoNum = 0 Then
   frmOptions.imgAppIcon(Index).Picture = imgIcoList.ListImages(1).Picture
Else
    frmOptions.imgAppIcon(Index).Picture = imgIconList.ListImages(icoNum).Picture
End If
End Sub

Private Sub KillProc(AppName2Kill As String)
'for kill process
Dim TheLoopingProcess
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim exename As String, AppLength As Integer
Dim ProcNum As Long
'---------------------------------
    AppLength = Len(AppName2Kill)
    AppName2Kill = UCase$(AppName2Kill)
    If Right$(AppName2Kill, 4) <> ".EXE" Then AppName2Kill = AppName2Kill & ".EXE"
    snap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&) 'TH32CS_SNAPall, 0) 'get snapshot handle
    proc.dwSize = Len(proc)
    TheLoopingProcess = ProcessFirst(snap, proc)       'first process and return value
    While TheLoopingProcess <> 0      'next process
        exename = UCase$(Left$(proc.szexeFile, AppLength))
        ProcNum = proc.th32ProcessID
        If exename = AppName2Kill Then
            KillProcessById (proc.th32ProcessID)
        End If
        TheLoopingProcess = ProcessNext(snap, proc)
    Wend
    CloseHandle snap
End Sub

Private Sub IsRun(AppName As String, rIndex As Integer)
'for check process whether program running
Dim TheLoopingProcess
Dim proc As PROCESSENTRY32
Dim snap As Long, tmpIndex As Integer, nCount As Integer
Dim exename(1000) As String, AppLength As Integer
Dim ProcNum As Long, IndName As Integer
'---------------------------------
    AppLength = Len(AppName)
    AppName = UCase$(AppName)
    If Right$(AppName, 4) <> ".EXE" Then AppName = AppName & ".EXE"
    snap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&) 'TH32CS_SNAPall, 0) 'get snapshot handle
    proc.dwSize = Len(proc)
    TheLoopingProcess = ProcessFirst(snap, proc)       'first process and return value
    While TheLoopingProcess <> 0      'next process
        exename(IndName) = UCase$(Left$(proc.szexeFile, AppLength))
        ProcNum = proc.th32ProcessID
        TheLoopingProcess = ProcessNext(snap, proc)
        IndName = IndName + 1
    Wend
    For tmpIndex = 0 To IndName
        If exename(tmpIndex) = AppName Then nCount = nCount + 1
    Next tmpIndex
    If nCount <= 0 Then
        imgApp(rIndex).BorderStyle = 0
    Else
        imgApp(rIndex).BorderStyle = 1
    End If
    CloseHandle snap
    IndName = 0
End Sub

Private Sub SetDate(dDay As Integer, dMonth As Integer, dYear As Integer, lLang As String)
If UCase$(lLang) = "THAI" Then
    cboMonth.AddItem "Á¡ÃÒ¤Á"           '1
    cboMonth.AddItem "¡ØÁÀÒ¾Ñ¹¸ì"       '2
    cboMonth.AddItem "ÁÕ¹Ò¤Á"           '3
    cboMonth.AddItem "àÁÉÒÂ¹"           '4
    cboMonth.AddItem "¾ÄÉÀÒ¤Á"          '5
    cboMonth.AddItem "ÁÔ¶Ø¹ÒÂ¹"         '6
    cboMonth.AddItem "¡Ã¡®Ò¤Á"          '7
    cboMonth.AddItem "ÊÔ§ËÒ¤Á"          '8
    cboMonth.AddItem "¡Ñ¹ÂÒÂ¹"          '9
    cboMonth.AddItem "µØÅÒ¤Á"           '10
    cboMonth.AddItem "¾ÄÈ¨Ô¡ÒÂ¹"        '11
    cboMonth.AddItem "¸Ñ¹ÇÒ¤Á"          '12
    For i = 0 To 9
        cboYear.AddItem dYear + i + 543
    Next i
    cboYear.Text = dYear + 543
    cboYear.toolTipText = dYear + 543
    Select Case dMonth
    Case Is = 1
        cboMonth.Text = "Á¡ÃÒ¤Á"
        cboMonth.toolTipText = "Á¡ÃÒ¤Á"
        nMonth = 1
    Case Is = 2
        cboMonth.Text = "¡ØÁÀÒ¾Ñ¹¸ì"
        cboMonth.toolTipText = "¡ØÁÀÒ¾Ñ¹¸ì"
        nMonth = 2
    Case Is = 3
        cboMonth.Text = "ÁÕ¹Ò¤Á"
        cboMonth.toolTipText = "ÁÕ¹Ò¤Á"
        nMonth = 3
    Case Is = 4
        cboMonth.Text = "àÁÉÒÂ¹"
        cboMonth.toolTipText = "àÁÉÒÂ¹"
        nMonth = 4
    Case Is = 5
        cboMonth.Text = "¾ÄÉÀÒ¤Á"
        cboMonth.toolTipText = "¾ÄÉÀÒ¤Á"
        nMonth = 5
    Case Is = 6
        cboMonth.Text = "ÁÔ¶Ø¹ÒÂ¹"
        cboMonth.toolTipText = "ÁÔ¶Ø¹ÒÂ¹"
        nMonth = 6
    Case Is = 7
        cboMonth.Text = "¡Ã¡®Ò¤Á"
        cboMonth.toolTipText = "¡Ã¡®Ò¤Á"
        nMonth = 7
    Case Is = 8
        cboMonth.Text = "ÊÔ§ËÒ¤Á"
        cboMonth.toolTipText = "ÊÔ§ËÒ¤Á"
        nMonth = 8
    Case Is = 9
        cboMonth.Text = "¡Ñ¹ÂÒÂ¹"
        cboMonth.toolTipText = "¡Ñ¹ÂÒÂ¹"
        nMonth = 9
    Case Is = 10
        cboMonth.Text = "µØÅÒ¤Á"
        cboMonth.toolTipText = "µØÅÒ¤Á"
        nMonth = 10
    Case Is = 11
        cboMonth.Text = "¾ÄÈ¨Ô¡ÒÂ¹"
        cboMonth.toolTipText = "¾ÄÈ¨Ô¡ÒÂ¹"
        nMonth = 11
    Case Is = 12
        cboMonth.Text = "¸Ñ¹ÇÒ¤Á"
        cboMonth.toolTipText = "¸Ñ¹ÇÒ¤Á"
        nMonth = 12
    End Select
Else
    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    For i = 0 To 9
        cboYear.AddItem dYear + i
    Next i
    cboYear.Text = dYear
    cboYear.toolTipText = dYear
    Select Case dMonth
    Case Is = 1
        cboMonth.Text = "January"
        cboMonth.toolTipText = "January"
        nMonth = 1
    Case Is = 2
        cboMonth.Text = "February"
        cboMonth.toolTipText = "February"
        nMonth = 2
    Case Is = 3
        cboMonth.Text = "March"
        cboMonth.toolTipText = "March"
        nMonth = 3
    Case Is = 4
        cboMonth.Text = "April"
        cboMonth.toolTipText = "April"
        nMonth = 4
    Case Is = 5
        cboMonth.Text = "May"
        cboMonth.toolTipText = "May"
        nMonth = 5
    Case Is = 6
        cboMonth.Text = "June"
        cboMonth.toolTipText = "June"
        nMonth = 6
    Case Is = 7
        cboMonth.Text = "July"
        cboMonth.toolTipText = "July"
        nMonth = 7
    Case Is = 8
        cboMonth.Text = "August"
        cboMonth.toolTipText = "August"
        nMonth = 8
    Case Is = 9
        cboMonth.Text = "September"
        cboMonth.toolTipText = "September"
        nMonth = 9
    Case Is = 10
        cboMonth.Text = "October"
        cboMonth.toolTipText = "October"
        nMonth = 10
    Case Is = 11
        cboMonth.Text = "November"
        cboMonth.toolTipText = "November"
        nMonth = 11
    Case Is = 12
        cboMonth.Text = "December"
        cboMonth.toolTipText = "December"
        nMonth = 12
    End Select
End If
Select Case dMonth
Case Is = 1, 3, 5, 7, 8, 10, 12
    For i = 1 To 31
        cboDay.AddItem i
    Next i
Case Is = 4, 6, 9, 11
    For i = 1 To 30
        cboDay.AddItem i
    Next i
Case Is = 2
    If ((cboYear - 2004) Mod 4) = 0 Then
        For i = 1 To 29
            cboDay.AddItem i
        Next i
    Else
        For i = 1 To 28
            cboDay.AddItem i
        Next i
    End If
End Select
cboDay.Text = dDay
cboDay.toolTipText = dDay
End Sub

Private Sub Timer4_Timer()
StatusBar1.Panels(1).Text = Winsock1.LocalIP & ":" & Winsock1.LocalHostName
StatusBar1.Panels(1).Width = Me.Width
End Sub

Private Sub txt2Search_KeyPress(KeyAscii As Integer)
AddedItem = False
If KeyAscii = 13 Then Call cmdSearch_Click
End Sub

Private Sub LoadList()
NoteFiles.Refresh
NoteFiles.Path = rcNotePath
lstNote.Clear
If NoteFiles.ListCount > 0 Then
    For fNum = 1 To NoteFiles.ListCount
        lstNote.AddItem fNum & ". " & NoteFiles.List(fNum - 1)
    Next fNum
Else
    tmpText = "No note file(s)."
    lstNote.AddItem tmpText
End If
If lstNote.ListCount > 0 And lstNote.List(0) <> "No note file(s)." Then frmPopWarning.Show
End Sub

Private Sub LoadDateList(tmpDate As String, tmpMonth As String, tmpYear As String)
If UCase$(GetLanguage) = "THAI" Then
    HeadDate = Trim$(tmpDate) & "_" & Trim$(tmpMonth) & "_" & Trim$(Str(Val(tmpYear) + 543))
Else
    HeadDate = Trim$(tmpDate) & "_" & Trim$(tmpMonth) & "_" & Trim$(tmpYear)
End If
    HdLength = Len(Trim$(HeadDate))
    NoteFiles.Refresh
    NoteFiles.Path = rcNotePath
    If NoteFiles.ListCount > 0 Then
        For tmpIndex1 = 0 To (NoteFiles.ListCount - 1)
            If Left$(NoteFiles.List(tmpIndex1), HdLength) = HeadDate Then
                tmpFile(tmpIndex2) = NoteFiles.List(tmpIndex1)
                tmpIndex2 = tmpIndex2 + 1
            End If
        Next tmpIndex1
        lstNote.Clear
        For tmpIndex1 = 0 To (tmpIndex2 - 1)
            lstNote.AddItem tmpIndex1 + 1 & ". " & tmpFile(tmpIndex1)
        Next tmpIndex1
    Else
        tmpText = "No note file(s)."
        lstNote.Clear
        lstNote.AddItem tmpText
    End If
    If lstNote.ListCount = 0 Then
        tmpText = "No note file(s)."
        lstNote.Clear
        lstNote.AddItem tmpText
    End If
    'Clear values
    tmpIndex1 = 0
    tmpIndex2 = 0
    For i = 0 To NoteFiles.ListCount
        tmpFile(i) = ""
    Next i
    If lstNote.ListCount > 0 And lstNote.List(0) <> "No note file(s)." Then frmPopWarning.Show
End Sub

Private Sub SetScrRes()
Dim res As String, ResW As Single, ResH As Single, Ans, App2Run As String
Dim tmpBPP As String, BitPP As Integer, tmpL As Integer, tmpInd As Integer

ScrRes = GetScrResolution
res = Left$(ScrRes, 9)
ResW = Val(Left$(res, 4))
ResH = Val(Right$(res, 4))
tmpL = Len(ScrRes)
Do
    tmpInd = tmpInd + 1
    tmpBPP = Left$(ScrRes, tmpInd)
Loop Until Right$(tmpBPP, 1) = "("
tmpL = Len(tmpBPP)
tmpBPP = Mid$(ScrRes, tmpL + 1, 3)
BitPP = Val(Trim$(tmpBPP))
If ResW < 1024 And ResH < 768 Then
    Ans = MsgBox("Your screen resolution is: " & ScrRes & ". " & _
        "It is too low resolution to run this program. Because " & _
        "this program require a screen resolution at least 1024x768 " & _
        "pixels. Would you like to set your screen resolution to " & _
        "1024x768 pixels with " & BitPP & " bit color for run RC Menu Bar?" & _
        vbCrLf & "If you need to change your screen resolution to 1024x" & _
        "768 pixels with " & BitPP & " bit color, please click 'Yes' " & _
        "button. But if you do not need to change your screen resolution " & _
        "please click 'No' button and  then RC Menu bar will quit autometically.", _
        vbQuestion + vbYesNo, "Question Change Screen Resolution")
    If Ans = vbYes Then
        Call ChangeRes(1024, 768, BitPP)
        MsgBox "Now your screen resolution is 1024x768 pixels with " & BitPP & _
                " bit color. Please restart RC Menu Bar again after program " & _
                "closed and then you will" & _
                "experience the new utility." & vbCrLf & "Thank you for use RC " & _
                "Menu Bar and hope you enjoy with this utility program.", _
                vbInformation + vbOKOnly, "Finished process..."
        Unload Me
    Else
        MsgBox "Sorry for the incomfortable of using RC Menu Bar. Hope you " & _
                "will be back to use RC Menu Bar again later.", _
                vbInformation + vbOKOnly, "Apologize message..."
        Unload Me
    End If
End If
End Sub

