VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Searching"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "frmSearchList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMoveToFolder 
      Caption         =   "&Move To Folder..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopyToFolder 
      Caption         =   "&Copy To Folder..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "&Open Folder"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open &File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Enabled Preview Image File"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.Frame FramePreview 
      Caption         =   "Preview Area"
      Height          =   2895
      Left            =   5040
      TabIndex        =   20
      Top             =   600
      Width           =   3015
      Begin VB.PictureBox picTarget 
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   2715
         TabIndex        =   21
         Top             =   240
         Width           =   2775
         Begin RichTextLib.RichTextBox rtbViewText 
            Height          =   2500
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   4419
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmSearchList.frx":27A2
         End
         Begin AniGIFCtrl.AniGIF AniGIF 
            Height          =   2415
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Visible         =   0   'False
            Width           =   2775
            BackColor       =   12632256
            PLaying         =   -1  'True
            Transparent     =   -1  'True
            Speed           =   1
            Stretch         =   1
            AutoSize        =   0   'False
            SequenceString  =   ""
            Sequence        =   0
            HTTPProxy       =   ""
            HTTPUserName    =   ""
            HTTPPassword    =   ""
            MousePointer    =   0
            ExtendWidth     =   4895
            ExtendHeight    =   4260
            Loop            =   0
            AutoRewind      =   0   'False
            Synchronized    =   -1  'True
         End
      End
      Begin VB.PictureBox picTmp 
         Height          =   1815
         Left            =   1200
         ScaleHeight     =   1755
         ScaleWidth      =   1035
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Details"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   7935
      Begin VB.PictureBox picBuffer 
         Height          =   255
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   17
         Top             =   -120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMoreInfo 
         Caption         =   "&More File Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Arch&ive"
         Height          =   195
         Left            =   4800
         TabIndex        =   16
         Top             =   900
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "&Hidden"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   550
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Read-Only"
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ImageList imgIconList 
         Left            =   1440
         Top             =   -480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSearchList.frx":2834
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image imgFileIcon 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblFileInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboTextSearch 
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
      Left            =   120
      TabIndex        =   0
      Text            =   "Text to search"
      Top             =   120
      Width           =   3735
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgResults"
      SmallIcons      =   "imgResults"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Filename"
         Text            =   "Filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "File Path"
         Text            =   "File Path"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "File Size"
         Text            =   "File Size"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.ListBox lstTmpResults 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstTmpDirs 
      Height          =   1815
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblFound 
      Caption         =   "Found: 0 file(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   11
      Top             =   600
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label lblSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Searching..."
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   7935
   End
   Begin VB.Menu MenuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu MenuOpenFile 
         Caption         =   "Open &File"
      End
      Begin VB.Menu MenuOpenFolder 
         Caption         =   "&Open Folder"
      End
      Begin VB.Menu dash00 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCopyTo 
         Caption         =   "Copy To Folder..."
      End
      Begin VB.Menu MenuMoveTo 
         Caption         =   "Move To Folder..."
      End
      Begin VB.Menu dash01 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDel 
         Caption         =   "&Delete To Recycle"
      End
      Begin VB.Menu MenuDelForever 
         Caption         =   "Delete For&ever"
      End
      Begin VB.Menu dash02 
         Caption         =   "-"
      End
      Begin VB.Menu MenuProperties 
         Caption         =   "P&roperties"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, numSDrives As Integer, drv As String
Dim tmpFile As String, tmpFileName(300000) As String, tmpIndex As Integer
Dim tmpFilePath(300000) As String, EnabledPreview As String, tmpPath As String
Dim fso As New FileSystemObject

Private Sub cboTextSearch_KeyPress(KeyAscii As Integer)
AddedItem = False
If KeyAscii = 13 Then Call cmdSearch_Click
End Sub

Private Sub Check4_Click()
SaveSetting App.ProductName, "Saved", "EnabledPreview", Check4.Value
If Check4.Value = 0 Then picTarget.Picture = LoadPicture()
End Sub

Private Sub cmdCopyToFolder_Click()
Call MenuCopyTo_Click
End Sub

Private Sub cmdMoreInfo_Click()
'show windows properties of file
tmpFile = lstResults.SelectedItem.SubItems(1) & lstResults.SelectedItem.Text
Call ShowFileProp(tmpFile, Me)
End Sub

Private Sub cmdMoveToFolder_Click()
Call MenuMoveTo_Click
End Sub

Private Sub cmdOpenFile_Click()
LaunchWithDefaultApp tmpFile
End Sub

Private Sub cmdOpenFolder_Click()
LaunchWithDefaultApp tmpPath
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim start_time As Double
Dim NextDir As String
i = 0
If cmdSearch.Caption = "&Search" Then
    If AddedItem = False Then
        AddedItem = True
        frmMain.txt2Search.AddItem cboTextSearch.Text
        cboTextSearch.AddItem cboTextSearch.Text
    End If
    Me.Caption = "Searching... " & cboTextSearch.Text
    frmMain.txt2Search.Text = cboTextSearch.Text
    lblFileInfo.Caption = ""
    imgFileIcon.Picture = LoadPicture()
    picTarget.Picture = LoadPicture()
    cmdSearch.Caption = "&Cancel"
    lblFound.Caption = "Found: 0 file(s)"
    cboTextSearch.Enabled = False
    'lstResults.Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    cmdMoreInfo.Enabled = False
    start_time = Timer
    'Clears the result listbox
    lstResults.ListItems.Clear
    
    numSDrives = frmOptions.lstDrive.ListCount
    'Calls the FindAllFiles function
    Do While i <> numSDrives
        If frmOptions.lstDrive.Selected(i) = True Then
            drv = frmOptions.lstDrive.List(i)
            FindAllFiles drv, cboTextSearch.Text
        End If
        i = i + 1
    Loop

    Do While lstTmpDirs.ListCount
        'Searches through all the new directories and removes
        'Them from the temp dir listbox
        NextDir = lstTmpDirs.List(0)
        lstTmpDirs.RemoveItem 0
        FindAllFiles NextDir, cboTextSearch.Text
        
        If lstResults.ListItems.Count > 1000000 Then
            'Makes sure there aren't too many results
            'If there are too many, the listbox can't hold them
            lblSearch.Caption = "Too Many Results..."
            lstTmpDirs.Clear
            lstTmpResults.Clear
            cmdSearch.Caption = "&Search"
            cboTextSearch.Enabled = True
            lstResults.Enabled = True
            Frame1.Enabled = True
            Exit Sub
        End If
        lblSearch.Caption = "Searching... " & lstTmpDirs.List(0)
    Loop
End If
    'Changes the status label
    cmdSearch.Caption = "&Search"
    cboTextSearch.Enabled = True
    lstResults.Enabled = True
    Check1.Enabled = True
    Check2.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    cmdMoreInfo.Enabled = True
    lblSearch.Caption = "Found: " & lstResults.ListItems.Count & " file(s) in " & Round(Timer - start_time, 1) & " sec"
    Me.Caption = "Finished search"
    lstTmpDirs.Clear
    lstTmpResults.Clear
End Sub

Private Sub Form_Load()
Me.Show
SearchVisibled = True
EnabledPreview = GetSetting(App.ProductName, "Saved", "EnabledPreview")
Check4.Value = Val(EnabledPreview)
cboTextSearch.Text = frmMain.txt2Search.Text
Me.Caption = "Searching... " & frmMain.txt2Search.Text
If cboTextSearch.Text <> "" Or cboTextSearch.Text <> "Text to search" Then _
    Call cmdSearch_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
SearchVisibled = False
lstResults.ListItems.Clear
cboTextSearch.Text = "Text to search"
frmMain.Enabled = True
frmMain.SetFocus
End Sub

Private Sub FileInfo(fPath As String)
On Error Resume Next
Dim icoNum As Integer
icoNum = GetIconFile(fPath, imgIconList, picBuffer, 32)
If icoNum = 0 Then
   imgFileIcon.Picture = imgIconList.ListImages(1).Picture
Else
    imgFileIcon.Picture = imgIconList.ListImages(icoNum).Picture
End If
End Sub

Public Function FindAllFiles(Directory As String, Optional SearchFor As String)
On Error Resume Next
    Dim Exists As Long
    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    Dim litem As ListItem
    
    SearchFor = LCase$(SearchFor)
    'Sets Exists to equal 1
    'You need this so the loop doesn't automatically exit
    Exists = 1
    
    'Makes sure theres a "\" at the end of the directory
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    'Sets the default search item to *.*
    If SearchFor = vbNullString Then SearchFor = "*.*"
    
    'If the search for text doesn't contain any * or ?
    'Add *'s before and after
    If InStr(1, SearchFor, "?") = 0 And InStr(1, SearchFor, "*") = 0 Then
        SearchFor = "*" & SearchFor & "*"
    End If
    
    'Finds the first file
    hFindFile = FindFirstFile(Directory & SearchFor, FileData)
    
    Do While hFindFile <> -1 And Exists <> 0
        'A loop until all the files have been added
        
        If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        <> vbDirectory Then
            'If the file isn't a directory than add it
            'to the temp listbox
            lstTmpResults.AddItem Directory & ClearNull(FileData.cFileName)
            tmpFileName(tmpIndex) = ClearNull(FileData.cFileName)
            tmpFilePath(tmpIndex) = Directory
        End If
        
        'Finds next file
        Exists = FindNextFile(hFindFile, FileData)
        tmpIndex = tmpIndex + 1
    Loop
    tmpIndex = 0
    Do While lstTmpResults.ListCount
        'Removes everything from the temp listbox (Which is
        'alphabetically sorted, and puts it into the Viewed
        'Listbox
        'This is done so all the files are sorted alphabetically
        Set litem = Me.lstResults.ListItems.Add(, , tmpFileName(tmpIndex))
        litem.ListSubItems.Add , , tmpFilePath(tmpIndex)
        litem.ListSubItems.Add , , GetFileSize(tmpFilePath(tmpIndex) & tmpFileName(tmpIndex))
        lblFound.Caption = "Found: " & lstResults.ListItems.Count & " file(s)"
        lstTmpResults.RemoveItem 0
        If cmdSearch.Caption = "&Search" Then Exit Function
        tmpIndex = tmpIndex + 1
    Loop
    tmpIndex = 0
    'Sets Exists to equal 1
    'You need this so the loop doesn't automatically exit
    Exists = 1
    
    'Find first file, this time includes directories in
    'the search
    hFindFile = FindFirstFile(Directory & "*", FileData)
    
    Do While hFindFile <> -1 And Exists <> 0
        'A loop until all the files have been added
        If (GetAttr(Directory & ClearNull(FileData.cFileName)) And vbDirectory) _
        = vbDirectory And Left(FileData.cFileName, 1) <> "." Then
            'If the file IS a directory and isn't "." or ".."
            'than adds it to the temp dir listbox
            lstTmpDirs.AddItem Directory & ClearNull(FileData.cFileName)
            DoEvents
        End If
        
        'Finds next file
        Exists = FindNextFile(hFindFile, FileData)
    Loop
End Function

Public Function ClearNull(StringToClear As String) As String
    Dim StartOfNulls As Long
    
    'This function clears all the nulls in the string and
    'Returns it, by using Instr to find the first null
    StartOfNulls = InStr(1, StringToClear, Chr(0))
    ClearNull = Left(StringToClear, StartOfNulls - 1)
End Function

Private Function GetTime(FileTimeA As FILETIME) As String
    Dim TmpFileTime As SYSTEMTIME
    
    'Changes the complicated FileTime to a SystemTime
    'Which is much easier to understand
    FileTimeToSystemTime FileTimeA, TmpFileTime
    
    If TmpFileTime.wMinute < 10 Then
        'If the minutes is less than 10, then add an extra 0
        'After the colon ":" and then set GetTime to the time
        GetTime = TmpFileTime.wDay & " " & GetMonth(TmpFileTime.wMonth, GetLanguage) & " " & GetYear(TmpFileTime.wYear, GetLanguage) & _
                    ", " & TmpFileTime.wHour & ":0" & TmpFileTime.wMinute
    Else
        'Set GetTime to the time
        GetTime = TmpFileTime.wDay & " " & GetMonth(TmpFileTime.wMonth, GetLanguage) & " " & GetYear(TmpFileTime.wYear, GetLanguage) & _
                    ", " & TmpFileTime.wHour & ":" & TmpFileTime.wMinute
    End If
End Function

Private Sub lstResults_Click()
On Error GoTo ErrHnd
Dim FileData As WIN32_FIND_DATA
Dim X&, Y&, X1&, Y1&, z1!, picPath As String

tmpFile = lstResults.SelectedItem.SubItems(1) & lstResults.SelectedItem.Text
tmpPath = lstResults.SelectedItem.SubItems(1)
cmdOpenFile.Enabled = True
cmdOpenFolder.Enabled = True
cmdCopyToFolder.Enabled = True
cmdMoveToFolder.Enabled = True
Call FileInfo(tmpFile)
GetAttributes tmpFile
Check1.Value = Atrib.ReadOnly
Check2.Value = Atrib.Hidden
Check3.Value = Atrib.Archive
'Checks to see if the file still exists
'If it does then FileData has all the info on the file
'If not then the Info label is updated and the sub is exited
If FindFirstFile(tmpFile, FileData) = "-1" Then
    lblFileInfo.Caption = "File No Longer Exists"
    Exit Sub
End If
'Changes lblInfo to tell all the file info
'This uses the ClearNull function and GetTime function
lblFileInfo.Caption = "Created On: " & GetTime(FileData.ftCreationTime) & _
    vbCrLf & "Last Accessed: " & GetTime(FileData.ftLastAccessTime) & _
    vbCrLf & "Last Written: " & GetTime(FileData.ftLastWriteTime)
If Check4.Value = 1 Then
  If Right$(UCase$(tmpFile), 3) = "JPG" Or _
        Right$(UCase$(tmpFile), 3) = "BMP" Or _
        Right$(UCase$(tmpFile), 3) = "GIF" Or _
        Right$(UCase$(tmpFile), 3) = "DIB" Or _
        Right$(UCase$(tmpFile), 4) = "JPEG" Or _
        Right$(UCase$(tmpFile), 3) = "JPE" Or _
        Right$(UCase$(tmpFile), 4) = "JFIF" Or _
        Right$(UCase$(tmpFile), 3) = "PNG" Or _
        Right$(UCase$(tmpFile), 3) = "ICO" Or _
        Right$(UCase$(tmpFile), 3) = "CUR" Or _
        Right$(UCase$(tmpFile), 3) = "WMF" Or _
        Right$(UCase$(tmpFile), 3) = "TIF" Or _
        Right$(UCase$(tmpFile), 4) = "TIFF" Then
    rtbViewText.Visible = False
    'Set default stuffs
    picTarget.Cls
    picTarget.AutoRedraw = True
    picTmp.AutoSize = True
    'get target sizing info
    X = picTarget.Width
    Y = picTarget.Height
    'Load the image
    If Right$(UCase$(tmpFile), 3) = "PNG" Then PngPictureLoad tmpFile, picTmp, False _
    Else picTmp.Picture = LoadPicture(tmpFile)
    If Right$(UCase$(tmpFile), 4) = "JPEG" Or _
        Right$(UCase$(tmpFile), 3) = "JPE" Or _
        Right$(UCase$(tmpFile), 4) = "JFIF" Or _
        Right$(UCase$(tmpFile), 3) = "TIF" Or _
        Right$(UCase$(tmpFile), 4) = "TIFF" Then
        SavePicture picTmp.Picture, App.Path & "\tmpPic.bmp"
        picPath = App.Path & "\tmpPic.bmp"
        picTmp.Picture = LoadPicture(picPath)
    End If
    
    'get source sizing info
    X1 = picTmp.Width
    Y1 = picTmp.Height
    'Determine conversion ratio to use
    z1 = IIf(X / X1 * Y1 < Y, X / X1, Y / Y1)
    'Calculate new image size
    X1 = X1 * z1
    Y1 = Y1 * z1
    picTarget.Visible = True
    'Draw Image
    If Right$(UCase$(tmpFile), 3) = "GIF" Then
        AniGIF.Left = (X - X1) / 2
        AniGIF.Top = (Y - Y1) / 2
        AniGIF.Width = X1
        AniGIF.Height = Y1
        AniGIF.Visible = True
        AniGIF.ReadGIF tmpFile
    Else
        AniGIF.StopReadGIF
        AniGIF.Visible = False
        picTarget.PaintPicture picTmp.Picture, (X - X1) / 2, (Y - Y1) / 2, X1, Y1
    End If
  Else 'If Right$(UCase$(tmpFile), 3) = "TXT" Or _
         Right$(UCase$(tmpFile), 3) = "DOC" Then
        rtbViewText.Visible = True
        AniGIF.Visible = False
        rtbViewText.LoadFile tmpFile
  End If
End If
Exit Sub
ErrHnd:
picTarget.Picture = LoadPicture()
AniGIF.Visible = False
rtbViewText.Visible = False
picTarget.Print vbCrLf & vbCrLf & vbCrLf & _
                "       Cannot preview this image." & vbCrLf & vbCrLf & _
                "       But you can double click on" & vbCrLf & _
                "       image file to preview with" & vbCrLf & _
                "       default Windows viewer."
Exit Sub
End Sub

Private Sub lstResults_DblClick()
On Error GoTo ErrHnd
tmpFile = lstResults.SelectedItem.SubItems(1) & lstResults.SelectedItem.Text
LaunchWithDefaultApp (tmpFile)
'SHELL tmpFile, vbNormalFocus
Exit Sub
ErrHnd:
MsgBox lstResults.SelectedItem.Text & " can not open/run.", vbOKOnly + vbInformation _
        , "Cannot Open/Run " & lstResults.SelectedItem.Text
Exit Sub
End Sub

Private Function GetMonth(iMonth As Integer, LocalLang As String) As String
Dim tmpMonthName As String
Select Case UCase$(LocalLang)
Case Is = "THAI"
    Select Case iMonth
    Case Is = 1: GetMonth = "Á¡ÃÒ¤Á"    '1
    Case Is = 2: GetMonth = "¡ØÁÀÒ¾Ñ¹¸ì"       '2
    Case Is = 3: GetMonth = "ÁÕ¹Ò¤Á"           '3
    Case Is = 4: GetMonth = "àÁÉÒÂ¹"           '4
    Case Is = 5: GetMonth = "¾ÄÉÀÒ¤Á"          '5
    Case Is = 6: GetMonth = "ÁÔ¶Ø¹ÒÂ¹"         '6
    Case Is = 7: GetMonth = "¡Ã¡®Ò¤Á"          '7
    Case Is = 8: GetMonth = "ÊÔ§ËÒ¤Á"          '8
    Case Is = 9: GetMonth = "¡Ñ¹ÂÒÂ¹"          '9
    Case Is = 10: GetMonth = "µØÅÒ¤Á"           '10
    Case Is = 11: GetMonth = "¾ÄÈ¨Ô¡ÒÂ¹"        '11
    Case Is = 12: GetMonth = "¸Ñ¹ÇÒ¤Á"          '12
    End Select
Case Else
    Select Case iMonth
    Case Is = 1: GetMonth = "January"
    Case Is = 2: GetMonth = "February"
    Case Is = 3: GetMonth = "March"
    Case Is = 4: GetMonth = "April"
    Case Is = 5: GetMonth = "May"
    Case Is = 6: GetMonth = "June"
    Case Is = 7: GetMonth = "July"
    Case Is = 8: GetMonth = "August"
    Case Is = 9: GetMonth = "September"
    Case Is = 10: GetMonth = "October"
    Case Is = 11: GetMonth = "November"
    Case Is = 12: GetMonth = "December"
    End Select
End Select
End Function

Private Function GetYear(iYear As Integer, LocalLang As String) As String
Dim tmpYear As String
Select Case UCase$(LocalLang)
Case Is = "THAI"
    GetYear = Str(iYear + 543)
Case Else
    GetYear = Str(iYear)
End Select
End Function

Private Sub lstResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmpFile = lstResults.SelectedItem.SubItems(1) & lstResults.SelectedItem.Text
tmpPath = lstResults.SelectedItem.SubItems(1)
If Button = 2 Then PopupMenu MenuPopUp, , , , MenuOpenFile
End Sub

Private Sub MenuCopyTo_Click()
Dim DesFold As String
Dim TargFile As String, TargPath As String

TargFile = lstResults.SelectedItem.Text
TargPath = tmpPath
DesFold = FolderForCopyMove(TargFile, TargPath, True)
CopyFile tmpFile, DesFold
End Sub

Private Sub MenuDel_Click()
DeleteFileEx Me.hwnd, tmpFile, True
lstResults.Refresh
End Sub

Private Sub MenuDelForever_Click()
DeleteFile tmpFile
lstResults.Refresh
End Sub

Private Sub MenuMoveTo_Click()
Dim DesFold As String
Dim TargFile As String, TargPath As String

TargFile = lstResults.SelectedItem.Text
TargPath = tmpPath
DesFold = FolderForCopyMove(TargFile, TargPath, False)
MoveFile tmpFile, DesFold
lstResults.Refresh
End Sub

Private Sub MenuOpenFile_Click()
LaunchWithDefaultApp tmpFile
End Sub

Private Sub MenuOpenFolder_Click()
LaunchWithDefaultApp tmpPath
End Sub

Private Sub MenuProperties_Click()
Call ShowFileProp(tmpFile, Me)
End Sub

Private Function FolderForCopyMove(File As String, Path As String, cCopy As Boolean) As String
On Error Resume Next
'Open browse for folder selection
Dim sBuffer As String
Dim szTitle As String
Dim FullFilePath As String
Dim SelectedFolder As String

If Right$(Path, 1) <> "\" Then
    FullFilePath = Path & "\" & File
Else
    FullFilePath = Path & File
End If
If cCopy = True Then
    szTitle = "Select a folder for copy '" & File & "' to:"
Else
    szTitle = "Select a folder for move '" & File & "' to:"
End If
sBuffer = Space(MAX_PATH)
sBuffer = BrowseFolder(ssfDESKTOPDIRECTORY, szTitle)
If Right$(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
FolderForCopyMove = sBuffer
End Function
