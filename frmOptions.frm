VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options..."
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Save Note Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   45
      Top             =   1920
      Width           =   6735
      Begin VB.CheckBox Check5 
         Caption         =   "Transparent Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   1320
         Width           =   1935
      End
      Begin MSComctlLib.Slider sldTranVal 
         Height          =   375
         Left            =   2280
         TabIndex        =   46
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Max             =   100
         TickStyle       =   1
         TickFrequency   =   5
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse..."
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Customize"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Default"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   3135
      End
      Begin VB.TextBox txtNotePath 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   6375
      End
      Begin VB.Label lblTranPercent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0%"
         Enabled         =   0   'False
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
         Left            =   4580
         TabIndex        =   48
         Top             =   1320
         Width           =   700
      End
   End
   Begin MSComctlLib.ImageList imgIcoList 
      Left            =   840
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":48DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":A4FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Width           =   6735
      Begin VB.PictureBox picOutBox 
         Height          =   3015
         Left            =   240
         ScaleHeight     =   2955
         ScaleWidth      =   6015
         TabIndex        =   25
         Top             =   360
         Width           =   6070
         Begin VB.PictureBox picInBox 
            Height          =   5535
            Left            =   0
            ScaleHeight     =   5475
            ScaleWidth      =   5955
            TabIndex        =   26
            Top             =   0
            Width           =   6015
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   3
               Left            =   4920
               TabIndex        =   9
               Top             =   120
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   4
               Left            =   4920
               TabIndex        =   10
               Top             =   720
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   5
               Left            =   4920
               TabIndex        =   11
               Top             =   1320
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   6
               Left            =   4920
               TabIndex        =   12
               Top             =   1920
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   7
               Left            =   4920
               TabIndex        =   13
               Top             =   2520
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   8
               Left            =   4920
               TabIndex        =   14
               Top             =   3120
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   9
               Left            =   4920
               TabIndex        =   15
               Top             =   3720
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   10
               Left            =   4920
               TabIndex        =   16
               Top             =   4320
               Width           =   1000
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse..."
               Height          =   495
               Index           =   11
               Left            =   4920
               TabIndex        =   17
               Top             =   4920
               Width           =   1000
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   11
               Left            =   0
               TabIndex        =   44
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   10
               Left            =   0
               TabIndex        =   43
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   9
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   8
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   7
               Left            =   0
               TabIndex        =   40
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   6
               Left            =   0
               TabIndex        =   39
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   5
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   4
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblTmpAppPath 
               Height          =   375
               Index           =   3
               Left            =   0
               TabIndex        =   36
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   3
               Left            =   120
               Stretch         =   -1  'True
               Top             =   120
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   3
               Left            =   720
               TabIndex        =   35
               Top             =   120
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   4
               Left            =   120
               Stretch         =   -1  'True
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   4
               Left            =   720
               TabIndex        =   34
               Top             =   720
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   5
               Left            =   120
               Stretch         =   -1  'True
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   5
               Left            =   720
               TabIndex        =   33
               Top             =   1320
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   6
               Left            =   120
               Stretch         =   -1  'True
               Top             =   1920
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   6
               Left            =   720
               TabIndex        =   32
               Top             =   1920
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   7
               Left            =   120
               Stretch         =   -1  'True
               Top             =   2520
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   7
               Left            =   720
               TabIndex        =   31
               Top             =   2520
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   8
               Left            =   120
               Stretch         =   -1  'True
               Top             =   3120
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   8
               Left            =   720
               TabIndex        =   30
               Top             =   3120
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   9
               Left            =   120
               Stretch         =   -1  'True
               Top             =   3720
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   9
               Left            =   720
               TabIndex        =   29
               Top             =   3720
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   10
               Left            =   120
               Stretch         =   -1  'True
               Top             =   4320
               Width           =   495
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   10
               Left            =   720
               TabIndex        =   28
               Top             =   4320
               Width           =   4005
            End
            Begin VB.Label lblAppName 
               BackStyle       =   0  'Transparent
               Caption         =   "Application Name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   11
               Left            =   720
               TabIndex        =   27
               Top             =   4920
               Width           =   4005
            End
            Begin VB.Image imgAppIcon 
               Height          =   495
               Index           =   11
               Left            =   120
               Stretch         =   -1  'True
               Top             =   4920
               Width           =   495
            End
         End
      End
      Begin VB.VScrollBar VScroll 
         Height          =   3015
         LargeChange     =   5
         Left            =   6360
         Max             =   100
         SmallChange     =   5
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.ListBox lstDrive 
      Appearance      =   0  'Flat
      Height          =   1605
      Left            =   4800
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox Check4 
         Caption         =   "Show All Note File(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   3975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Auto hide menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   3975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Start menu when Windows start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Always set menu on top"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   19
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveSetting 
      Caption         =   "&Save Setting"
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   7440
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   3
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComctlLib.ImageList imgIconList 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin MSComDlg.CommonDialog cdbOpenApp 
      Index           =   11
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open for Select Application"
      Filter          =   "Application (*.exe)|*.exe"
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTmpApp 
      Height          =   495
      Left            =   0
      Picture         =   "frmOptions.frx":EBA8
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search drive(s):"
      Height          =   195
      Left            =   4800
      TabIndex        =   21
      Top             =   0
      Width           =   1110
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New Scripting.FileSystemObject
Dim drv As Drive, i As Integer
Dim tmpStr As String, dDrive As String
Dim tmpValue As Boolean, tmpAppPath(3 To 11) As String
Dim SaveChange As Boolean, tmpNotePath As String
Dim CheckVal(1 To 3) As Integer, tmpAppName As String
Dim AppName(3 To 11) As String, AppPath(3 To 11) As String
Dim OldScroll As Integer, tmpAppIco(3 To 11) As String
Dim tmpTranVal As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
    TopMost = SetTopMostWindow(frmMain.hwnd, True)
Else
    TopMost = SetTopMostWindow(frmMain.hwnd, False)
End If
SaveChange = False
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Call AddToRun("RC Menu Bar", App.Path & "\" & App.exename & ".exe")
Else
    Call RemoveFromRun("RC Menu Bar")
End If
SaveChange = False
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    frmMain.Timer3.Enabled = True
    AutohideMenu = True
Else
    frmMain.Timer3.Enabled = False
    AutohideMenu = False
End If
frmMain.Left = Screen.Width - frmMain.Width
frmMain.Top = 0
SaveChange = False
End Sub

Private Sub Check4_Click()
SaveChange = False
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
    sldTranVal.Enabled = True
    lblTranPercent.Enabled = True
ElseIf Check5.Value = 0 Then
    sldTranVal.Enabled = False
    lblTranPercent.Enabled = False
End If
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error Resume Next
'Open browse for folder selection
'Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
'Dim tBrowseInfo As BrowseInfo

Select Case Index
Case Is = 0 'Open browse for folder selection
    szTitle = "Select a folder for save RC Note file:"
    'With tBrowseInfo
    '.hwndOwner = Me.hwnd
    '.lpszTitle = lstrcat(szTitle, "")
    '.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    'End With

    'lpIDList = SHBrowseForFolder(tBrowseInfo)

    'If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        'SHGetPathFromIDList lpIDList, sBuffer
        'sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        sBuffer = BrowseFolder(ssfDESKTOPDIRECTORY, szTitle)
        If Right$(sBuffer, 1) <> "\" Then _
            sBuffer = sBuffer & "\"
        txtNotePath.Text = sBuffer
    'End If
Case Else
    cdbOpenApp(Index).ShowOpen
    i = 0
    If cdbOpenApp(Index).filename <> "" Then
        Do
            i = i + 1
            tmpAppName = Right$(cdbOpenApp(Index).filename, i)
        Loop Until Left$(tmpAppName, 1) = "\"
        lblAppName(Index).Caption = Right$(tmpAppName, i - 1)
        AppPath(Index) = Left$(cdbOpenApp(Index).filename, Len(cdbOpenApp(Index).filename) - i - 1)
        Call FileInfo(cdbOpenApp(Index).filename, Index)
        lblTmpAppPath(Index).Caption = cdbOpenApp(Index).filename
        frmMain.imgApp(Index).Picture = imgAppIcon(Index).Picture
        frmMain.imgApp(Index).Enabled = True
        frmMain.imgApp(Index).ToolTipText = lblAppName(Index).Caption
    Else
        AppName(Index) = "No Application"
        AppPath(Index) = "No Application"
    End If
End Select
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSaveSetting_Click()
Call SaveValue
Unload Me
End Sub

Private Sub Form_Load()
Call LoadSave
i = 0
dDrive = Left$(WinPath, 1)
For Each drv In fso.Drives
    lstDrive.AddItem drv.Path
    tmpStr = GetSetting(App.ProductName, "Saved", drv.Path)
    If tmpStr <> "" Then
        tmpValue = GetSetting(App.ProductName, "Saved", drv.Path)
        lstDrive.Selected(i) = tmpValue
    Else
        If drv.Path = dDrive & ":" Then _
            If lstDrive.Selected(i) = False Then lstDrive.Selected(i) = True
    End If
    i = i + 1
Next
OldScroll = VScroll.Value
SaveChange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Ans
If SaveChange = False Then
    Ans = MsgBox("You are not save your setting." & vbCrLf & _
                "Would you like to save your setting?", vbQuestion + vbYesNo, _
                "Save Question")
    If Ans = vbYes Then
        Call SaveValue
    Else
        Call LoadSave
    End If
End If
frmMain.lstNote.Refresh
If frmMain.chkRefreshNoteList.Value = 1 Then frmMain.chkRefreshNoteList.Value = 0 _
Else frmMain.chkRefreshNoteList.Value = 1
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Public Sub AddToRun(ProgramName As String, FileToRun As String)
    'Add a program to the 'Run at Startup
    '     registry keys
    Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun)
End Sub

Public Sub RemoveFromRun(ProgramName As String)
    'Remove a program from the 'Run at Start
    '     up' registry keys
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName)
End Sub

Private Sub SaveValue()
SaveSetting App.ProductName, "Saved", "Check1", Check1.Value
SaveSetting App.ProductName, "Saved", "Check2", Check2.Value
SaveSetting App.ProductName, "Saved", "Check3", Check3.Value
SaveSetting App.ProductName, "Saved", "Check4", Check4.Value
SaveSetting App.ProductName, "Saved", "Check5", Check5.Value
SaveSetting App.ProductName, "Saved", "TranVal", sldTranVal.Value
SaveSetting App.ProductName, "Saved", "Option1", Option1.Value
SaveSetting App.ProductName, "Saved", "Option2", Option2.Value
SaveSetting App.ProductName, "Saved", "SaveNotePath", txtNotePath.Text
SaveSetting App.ProductName, "Saved", "Autohidemenu", AutohideMenu
For i = 0 To lstDrive.ListCount - 1
    SaveSetting App.ProductName, "Saved", lstDrive.List(i), lstDrive.Selected(i)
Next i
For i = 3 To 11
    SaveSetting App.ProductName, "Saved", "AppIcon" & i, lblTmpAppPath(i).Caption
    SaveSetting App.ProductName, "Saved", "AppName" & i, lblAppName(i).Caption
Next i
SaveChange = True
End Sub

Private Sub lstDrive_Click()
SaveChange = False
End Sub

Private Sub LoadSave()
On Error Resume Next
Check1.Value = GetSetting(App.ProductName, "Saved", "Check1")
Check2.Value = GetSetting(App.ProductName, "Saved", "Check2")
Check3.Value = GetSetting(App.ProductName, "Saved", "Check3")
Check4.Value = GetSetting(App.ProductName, "Saved", "Check4")
Check5.Value = GetSetting(App.ProductName, "Saved", "Check5")
tmpTranVal = GetSetting(App.ProductName, "Saved", "TranVal")
sldTranVal.Value = tmpTranVal
Option1.Value = GetSetting(App.ProductName, "Saved", "Option1")
Option2.Value = GetSetting(App.ProductName, "Saved", "Option2")
tmpNotePath = GetSetting(App.ProductName, "Saved", "SaveNotePath")
If Option1.Value = True Then
    txtNotePath.Text = App.Path & "\SaveNote\"
Else
    If Option2.Value = True Then
        txtNotePath.Text = tmpNotePath
    Else
        txtNotePath.Text = App.Path & "\SaveNote\"
        Option1.Value = True
    End If
End If
For i = 0 To lstDrive.ListCount - 1
    lstDrive.List(i) = GetSetting(App.ProductName, "Saved", lstDrive.List(i))
Next i
For i = 3 To 11
    tmpAppIco(i) = GetSetting(App.ProductName, "Saved", "AppIcon" & i)
    tmpAppPath(i) = tmpAppIco(i)
    Call FileInfo(tmpAppIco(i), i)
    lblTmpAppPath(i).Caption = tmpAppPath(i)
Next i
End Sub

Private Sub Option1_Click()
txtNotePath.Text = App.Path & "\SaveNote\"
txtNotePath.Enabled = False
cmdBrowse(0).Enabled = False
End Sub

Private Sub Option2_Click()
txtNotePath.Enabled = True
cmdBrowse(0).Enabled = True
End Sub

Private Sub sldTranVal_Change()
lblTranPercent.Caption = sldTranVal.Value & "%"
Transparent frmMain, (255 - 2 * sldTranVal.Value)
End Sub

Private Sub VScroll_Change()
On Error Resume Next
If VScroll.Value > OldScroll Then       'move down
    If VScroll.Value Mod 5 = 0 Then picInBox.Top = picInBox.Top - 130
ElseIf VScroll.Value < OldScroll Then   'move up
    If VScroll.Value Mod 5 = 0 Then picInBox.Top = picInBox.Top + 130
End If
OldScroll = VScroll.Value
End Sub

Private Sub FileInfo(fPath As String, Index As Integer)
On Error Resume Next
Dim icoNum As Integer
icoNum = GetIconFile(fPath, imgIconList, picBuffer, 32)
If icoNum = 0 Then
   imgAppIcon(Index).Picture = imgIcoList.ListImages(1).Picture
Else
    imgAppIcon(Index).Picture = imgIconList.ListImages(icoNum).Picture
End If
End Sub

Private Sub VScroll_Scroll()
On Error Resume Next
If VScroll.Value > OldScroll Then       'move down
    picInBox.Top = 0 - 26 * VScroll.Value
    CurrentScroll = picInBox.Top
    'If picInBox.Top <= -2600 Then picInBox.Top = -2600
ElseIf VScroll.Value < OldScroll Then   'move up
    picInBox.Top = -2600 + 26 * (VScroll.Max - VScroll.Value)
    'If picInBox.Top >= 0 Then picInBox.Top = 0
End If
If VScroll.Value >= 100 Then
    OldScroll = 100
ElseIf VScroll.Value <= 0 Then
    OldScroll = 0
Else
    OldScroll = VScroll.Value
End If
End Sub
