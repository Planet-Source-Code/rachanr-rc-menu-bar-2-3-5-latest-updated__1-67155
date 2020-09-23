Attribute VB_Name = "RCMenuBarModule"
'Check for whether frmSearch visible.
Global SearchVisibled As Boolean
'Check for Autohide Menu
Global AutohideMenu As Boolean

'------------------------------------------------------
'Function for Cut Copy Move Paste and Delete files
'Copy declarations
Private Const FO_MOVE = &H1&
Private Const FO_COPY = &H2&
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4&
Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_CONFIRMMOUSE = &H2&
Private Const FOF_CREATEPROGRESSDLG = &H0&
Private Const FOF_FILESONLY = &H80&
Private Const FOF_MULTIDESTFILES = &H1&
Private Const FOF_NOCONFIRMATION = &H10&
Private Const FOF_NOCONFIRMMKDIR = &H200&
Private Const FOF_RENAMEONCOLLISION = &H8&
Private Const FOF_SILENT = &H4& 'Progress not visible
Private Const FOF_SIMPLEPROGRESS = &H100& 'do not show filenames
Private Const FOF_WANTMAPPINGHANDLE = &H20&

Private Const COPY_ERROR = vbObjectError + 1000
Private Const DEL_ERROR = vbObjectError + 1001
Private Const MOVE_ERROR = vbObjectError + 1002
Private Const RENAME_ERROR = vbObjectError + 1003

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
'--------------------------------------------------------

'Declare for view *.png file to picture box
Private Type GUID
   Data1    As Long
   Data2    As Integer
   Data3    As Integer
   Data4(7) As Byte
End Type

Private Type PICTDESC
   size     As Long
   Type     As Long
   hBmp     As Long
   hPal     As Long
   Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

' GDI+ functions
Private Declare Function GdipLoadImageFromFile Lib "gdiplus.dll" (ByVal filename As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal Token As Long)

' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

'Declare for check add item to search text combobox
Global AddedItem As Boolean

'Function Browse For Folder without "Create New Folder" button
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
'Public Const MAX_PATH = 260

Public Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Function Browse For Folder with "Create New Folder" button
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H40

Public Enum ShellSpecialFolderConstants
    'ssfALTSTARTUP = 29
    'ssfAPPDATA = 26
    'ssfBITBUCKET = 10
    'ssfCOMMONALTSTARTUP = 30
    'ssfCOMMONAPPDATA = 35
    'ssfCOMMONDESKTOPDIR = 25
    'ssfCOMMONFAVORITES = 31
    'ssfCOMMONPROGRAMS = 23
    'ssfCOMMONSTARTMENU = 22
    'ssfCOMMONSTARTUP = 24
    'ssfCONTROLS = 3
    'ssfCOOKIES = 33
    'ssfDESKTOP = 0
    ssfDESKTOPDIRECTORY = 16
    ssfDRIVES = 17
    'ssfFAVORITES = 6
    'ssfFONTS = 20
    'ssfHISTORY = 34
    'ssfINTERNETCACHE = 32
    'ssfLOCALAPPDATA = 28
    'ssfMYPICTURES = 39
    'ssfNETHOOD = 19
    'ssfNETWORK = 18
    'ssfPERSONAL = 5
    'ssfPRINTERS = 4
    'ssfPRINTHOOD = 27
    'ssfPROFILE = 40
    'ssfPROGRAMFILES = 38
    'ssfPROGRAMFILESx86 = 48
    'ssfPROGRAMS = 2
    'ssfRECENT = 8
    'ssfSENDTO = 9
    'ssfSTARTMENU = 11
    'ssfSTARTUP = 7
    'ssfSYSTEM = 37
    'ssfSYSTEMx86 = 41
    'ssfTEMPLATES = 21
    'ssfWINDOWS = 36
End Enum

'----------------------------------------------------------------
'
'              Show fileproperties in VB 5.0
'
'              written by D. Rijmenants 2004
'         Comments or suggestions are most welcome at
'                mail: dr.defcom@telenet.be
'
'----------------------------------------------------------------
Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Type SHELLEXECUTEINFO
       cbSize As Long
       fMask As Long
       hwnd As Long
       lpVerb As String
       lpFile As String
       lpParameters As String
       lpDirectory As String
       nShow As Long
       hInstApp As Long
       lpIDList As Long
       lpClass As String
       hkeyClass As Long
       dwHotKey As Long
       hIcon As Long
       hProcess As Long
End Type

'Declare for get attributes
Public Type ATTRIBUTES
    'Alias As Integer
    Archive As Integer
    'System As Integer
    ReadOnly As Integer
    'Volume As Integer
    'Directory As Integer
    Hidden As Integer
    'Normal As Integer
End Type
Global Atrib As ATTRIBUTES

'Find File declarations and types
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 260
        cAlternate As String * 14
End Type

'Convert Time Declare and type
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

'Get Country and Language
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Private Const LOCALE_SENGCOUNTRY = &H1002       '  English name of country

'--------------------------------------------------------------
'Declare for extract icon from file
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal FLAGS&) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim FileInfo As typSHFILEINFO
'-------------------------------------------------------

'Shell Execute function
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As Any, _
    ByVal lpDirectory As Any, _
    ByVal nShowCmd As Long) As Long

'Get Mouse Cursor position on screen
Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

'Declare for kill process
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Public Const MAX_PATH As Integer = 260

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    End Type

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
'Format message declaration
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1

'My computer constant
Public Const MYCOMPUTER As String = "explorer ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"

'Functions for get special folder
'Use this function and SHGetPathFromIDList function to get path of special folder
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'type for SHGetSpecialFolderLocation
Public Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type
'type for SHGetSpecialFolderLocation
Public Type ITEMIDLIST
    shellID As SHITEMID
End Type
'Public Const DESKTOP = &H0
'Public Const PROGRAMS = &H2
Public Const CONTROLPS = &H3
Public Const MYDOCS = &H5
'Public Const FAVORITES = &H6
'Public Const STARTUP = &H7
'Public Const RECENT = &H8
'Public Const SENDTO = &H9
Public Const MYCOMP = &H11                   'My Computer virtual folder
'Public Const STARTMENU = &HB
'Public Const NETHOOD = &H13
'Public Const FONTS = &H14
'Public Const SHELLNEW = &H15
'Public Const TEMPINETFILES = &H20
'Public Const COOKIES = &H21
'Public Const HISTORY = &H22

'Get system directory
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Function to get Windows folder
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim m_WinPath As String

'Get Memory status
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Public memInfo As MEMORYSTATUS

'Function about Registry for Add to Start Up
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    'Const ERROR_SUCCESS = 0&
    Const REG_SZ = 1 ' Unicode nul terminated String
    Const REG_DWORD = 4 ' 32-bit number
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

'***********************************************************************************
'***********************declarations for LaunchWithDefaultApp***********************
'***********************************************************************************
#If Win32 Then
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
#Else

Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, _
    ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
    ByVal lpszDir$, ByVal fsShowCmd%) As Integer

Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If
Private Const SW_SHOWNORMAL = 1
'End declaration for launchwithdefaultapp

'TOP MOST DECLARATION
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS_TM = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

'---Find & Set new SCREEN RESOLUTION------
Public Const ScrRatio = 13.8
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const ENUM_CURRENT_SETTINGS = &HFFFF - 1
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Get screen resolution function
Public Function GetScrResolution() As String
    Dim curDPS As DEVMODE
    Dim colors As String
    Dim SMR As Long
    
    SMR = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, curDPS)
    
    If SMR = 0 Then
        GetScrResolution = "Error evaluating the current screen resolution!"
    Else
        Select Case curDPS.dmBitsPerPel
            Case 4:      colors = "16 Color (4 bits)"
            Case 8:      colors = "256 Color (8 bits)"
            Case 16:     colors = "High Color (16 bits)"
            Case 24:     colors = "True Color (24 bits)"
            Case 32:     colors = "True Color (32 bits)"
            Case 64:     colors = "True Color (64 bits)"
            Case 128:     colors = "True Color (128 bits)"
        End Select
        GetScrResolution = Format(curDPS.dmPelsWidth, "@@@@") + "x" + _
                      Format(curDPS.dmPelsHeight, "@@@@") + "  " + _
                      Format(colors, "@@@@@@@@@@@@@@@@@@@@  ") + _
                      Format(curDPS.dmDisplayFrequency, "@@@ Hz")
    End If
End Function

'Set new screen resolution
Public Function ChangeRes(Width As Single, Height As Single, BPP As Integer) As Integer
    On Error GoTo ERROR_HANDLER
    Dim DevM As DEVMODE, i As Integer, ReturnVal As Boolean, _
        RetValue, OldWidth As Single, OldHeight As Single, _
        OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
    i = 0
    Do
        ReturnVal = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until (ReturnVal = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = Width
    DevM.dmPelsHeight = Height
    DevM.dmBitsPerPel = BPP
    Call ChangeDisplaySettings(DevM, 1)
    RetValue = MsgBox("Do You Wish To Keep Your Screen Resolution To " _
                & Width & "x" & Height & " - " & BPP & " bit per pixels?", _
                vbQuestion + vbOKCancel, "Change Resolution Confirm...")
    If RetValue = vbCancel Then
        DevM.dmPelsWidth = OldWidth
        DevM.dmPelsHeight = OldHeight
        DevM.dmBitsPerPel = OldBPP
        Call ChangeDisplaySettings(DevM, 1)
        MsgBox "Old Resolution(" & OldWidth & " x " & OldHeight & ", " & _
                OldBPP & " Bit) Successfully Restored!", vbInformation + _
                vbOKOnly, "Resolution Confirm..."
        ChangeRes = 0
    Else
        ChangeRes = 1
    End If
    Exit Function
ERROR_HANDLER:
    ChangeRes = 0
End Function
'TOP MOST FUNCTION
Public Function SetTopMostWindow(hwnd As Long, TopMost As Boolean) _
   As Long

   If TopMost = True Then
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS_TM)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS_TM)
      SetTopMostWindow = False
   End If
End Function

'Check file exists
Public Function FileExists(strFile As String) As Boolean
    If Dir(strFile, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
    Else FileExists = False
End Function

'Check folder exists
Function FolderExists(Path As String) As Boolean
    If Dir(Path, vbDirectory) <> "" Then FolderExists = True _
    Else FolderExists = False
End Function


'*** Launch With Default App ********************
'* Iputs: VData - The name of the file          *
'************************************************
Public Function LaunchWithDefaultApp(VData As String) As Long
      Dim Scr_hDC As Long
      Dim Dire As String
      Dire = Left(VData, 3)
      Scr_hDC = GetDesktopWindow()
      StartDoc = ShellExecute(Scr_hDC, "Open", VData, "", Dire, SW_SHOWNORMAL)
End Function

'Add and Remove key to/from registry
Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strData As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyHand As Long
    Dim R As Long
    R = RegCreateKey(hKey, strPath, keyHand)
    R = RegSetValueEx(keyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    R = RegCloseKey(keyHand)
End Sub

Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyHand As Long
    R = RegOpenKey(hKey, strPath, keyHand)
    R = RegDeleteValue(keyHand, strValue)
    R = RegCloseKey(keyHand)
End Function

Public Function ReadKey(Value As String) As String

Dim b As Object
Dim R
On Error Resume Next
Set b = CreateObject("wscript.shell")
R = b.regread(Value)
ReadKey = R
End Function

Public Function WinPath() As String
    'This function retrieves the Windows path.
    If m_WinPath = "" Then
        m_WinPath = String(1024, 0)
        GetWindowsDirectory m_WinPath, Len(m_WinPath)
        m_WinPath = Left(m_WinPath, InStr(m_WinPath, Chr(0)) - 1)
        If Right(m_WinPath, 1) <> "\" Then m_WinPath = m_WinPath & "\"
    End If
    WinPath = m_WinPath
End Function

Public Function SystemDir() As String
Dim result
Dim SystemDirectory As String
SystemDirectory = Space(144)
result = GetSystemDirectory(SystemDirectory, 144)
If result = 0 Then
    MsgBox "Cannot Get the Windows System Directory", vbCritical, "Warning"
Else
    SystemDir = Trim(SystemDirectory) & "\"
End If
End Function

Public Function GetSFolder(frmForm As Form, Folder As Long, myID As ITEMIDLIST) As String
Dim fPath As String * 256
Dim rval As Long

rval = SHGetSpecialFolderLocation(frmForm.hwnd, Folder, myID)
If rval = 0 Then
    rval = SHGetPathFromIDList(ByVal myID.shellID.SHItem, ByVal fPath)
    If rval Then
        GetSFolder = Left(fPath, InStr(fPath, Chr(0)) - 1)
    End If
End If
End Function

Public Sub KillProcessById(p_lngProcessId As Long)
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    
    If lngReturn = 0 Then
        RetrieveError
    End If
End Sub

'PROCESS

Public Sub RetrieveError()
  Dim strBuffer As String
    strBuffer = Space(200)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strBuffer, 200, ByVal 0&
    MsgBox strBuffer
End Sub

Function SetMouse(x, y)
    SetCursorPos x, y
End Function

Public Sub pause(ByVal nSecond As Single)
   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim dummy As Integer

      dummy = DoEvents()
      ' if we cross midnight, back up one day
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
      If fStop = True Then Exit Do
   Loop

End Sub

Public Sub ShellStart(Path_File As String)
    Dim Xx As Long
    
    Xx = ShellEx(0, "open", Path_File, "", "", 1)
End Sub

'Extract icon from any file function
Public Function GetIconFile(fName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(fName, 0&, FileInfo, Len(FileInfo), FLAGS Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(fName, 0&, FileInfo, Len(FileInfo), FLAGS Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      GetIconFile = IconIndex
    End If
End Function

Public Function GetLanguage() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE, Buffer, 99)
   GetLanguage = LPSTRToVBString(Buffer)
End Function

Private Function LPSTRToVBString$(ByVal s$)
   Dim nullpos&
   nullpos& = InStr(s$, Chr$(0))
   If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
   Else
      LPSTRToVBString = ""
   End If
End Function

Public Function GetFileSize(filename) As String
Dim s As Double
Dim st As String
s = FileLen(filename) / 1024
s = s / 1024
st = " Mb"
 If s > 1 Then
 GoTo ok
 Else
 s = s * 1024
 st = " Kb"
 End If
 If s > 1 Then
 GoTo ok
 Else
 s = s * 1024
 st = " bytes"
 End If
 
ok:
 GetFileSize = Round(s, 2) & st
fin:
End Function

'*** Get and Set attributes of o file ****************
Public Function GetAttributes(fName As String) As String
Dim Tmp As VbFileAttribute
Tmp = GetAttr(fName)
With Atrib
    .Archive = 0
    .Hidden = 0
    .ReadOnly = 0
End With

If Tmp >= vbArchive Then ' 32
    GetAttributes = GetAttributes & " Archive"
    Tmp = Tmp - vbArchive
    Atrib.Archive = 1
End If

If Tmp >= vbHidden Then '2
    GetAttributes = GetAttributes & " Hidden"
    Tmp = Tmp - vbHidden
    Atrib.Hidden = 1
End If

If Tmp >= vbReadOnly Then '1
    GetAttributes = GetAttributes & " Read Only"
    Tmp = Tmp - vbReadOnly
    Atrib.ReadOnly = 1
End If
End Function

Public Function ShowFileProp(ByVal filename As String, aForm As Form) As Long
'open the file properties for the filename
'if return <=32 error occured
Dim SEI As SHELLEXECUTEINFO
Dim R As Long
If filename = "" Then
    ShowFileProp = 0
    Exit Function
    End If
With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = aForm.hwnd
    .lpVerb = "properties"
    .lpFile = filename
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
End With
R = ShellExecuteEX(SEI)
ShowFileProp = SEI.hInstApp
End Function

'Load Png to Picture Control
Sub PngPictureLoad(PathFilename As String, PictureControl As PictureBox, AutoResize As Boolean)
   Dim Token    As Long
    Token = InitGDIPlus
    If AutoResize = False Then
     PictureControl = LoadPictureGDIPlus(PathFilename)
    Else
     PictureControl = LoadPictureGDIPlus(PathFilename, PictureControl.ScaleWidth / Screen.TwipsPerPixelX, PictureControl.ScaleHeight / Screen.TwipsPerPixelY)
    End If
    FreeGDIPlus Token
End Sub

' Initialises GDI Plus
Public Function InitGDIPlus() As Long
    Dim Token    As Long
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(PicFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long
        
    ' Load the image
    If GdipLoadImageFromFile(StrPtr(PicFile), Img) <> 0 Then
        Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    
    ' Initialise the hDC
    InitDC hDC, hBitmap, BackColor, Width, Height

    ' Resize the picture
    gdipResize Img, hDC, Width, Height, RetainRatio
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
End Function

' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub

' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub

' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As GUID
    Dim Pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.Data1 = &H20400
    IID_IDispatch.Data4(0) = &HC0
    IID_IDispatch.Data4(7) = &H46
        
    ' Fill Pic with necessary parts
    Pic.size = Len(Pic)        ' Length of structure
    Pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    Pic.hBmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect Pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function

Public Function BrowseFolder(OpenAt As ShellSpecialFolderConstants, strTitle As String) As String
    Dim ShellApplication As Object
    Dim Folder As Object
    Set ShellApplication = CreateObject("Shell.Application")
    On Error Resume Next
    Set Folder = ShellApplication.BrowseForFolder(0, strTitle, BIF_EDITBOX Or BIF_NEWDIALOGSTYLE, CInt(OpenAt))
    BrowseFolder = Folder.Items.Item.Path
    On Error GoTo 0

    If Left(BrowseFolder, 2) = "::" Or InStr(1, BrowseFolder, "\") = 0 Then
        BrowseFolder = vbNullString
    End If
End Function

'Delete to Recycle
Public Function DeleteFileEx(lHwnd As Long, sFilePathName As String, bToRecycleBin As Boolean, Optional bConfirm As Boolean = True) As Boolean
    ' Original code from Ruturaj
    ' Email : mailme_friends@yahoo.com
    '=======================================
    ' ReturnType : Boolean
    '=======================================
    ' Arguments : [1] lHwnd= hWnd Property o
    '     f calling Form.
    ' [2] sFilePathName= Full path of File w
    '     hich is to be deleted.
    ' [3] bToRecycleBin= Set to True if file
    '     is to be moved to Recycle Bin.
    ' [4] bConfirm = Optional. Default is Tr
    '     ue. Set to False if you
    ' don' want OS Confirmation prompts befo
    '     re
    ' performing Delete Action.
    '=======================================
    ' Purpose: This function can do followin
    '     g things ...
    ' [1] Show default system confirmation P
    '     rompt to move File to Recycle Bin
    ' [2] Move directly the selected File to
    '     Recycle Bin without any Prompt
    ' [3] Show default system confirmation P
    '     rompt to remove File forever.
    ' [4] Delete File forever without any pr
    '     ompt. (Same lile Kill function)
    '---------------------------------------
    On Error GoTo DeleteFileEx_Error
    Dim TSHStruct As SHFILEOPSTRUCT
    Dim lResult As Long
    'See if File exists ...

    If Dir(sFilePathName, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
        'Fill the necessary Structure elements b
        '     y specified values ...

        With TSHStruct
            .hwnd = lHwnd
            .pFrom = sFilePathName
            .wFunc = FO_DELETE
            'Flag settings ... the heart of this fun
            '     ction !

            If bToRecycleBin = True Then
                If bConfirm = True Then
                    .fFlags = FOF_ALLOWUNDO
                Else
                    .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
                End If
            ElseIf bToRecycleBin = False Then
                If bConfirm = False Then
                    .fFlags = FOF_NOCONFIRMATION
                End If
            End If
        End With
        'It's Show-Time !
        lResult = SHFileOperation(TSHStruct)
        'SHFileOperation returns Zero if success
        '     ful or non-zero if failed.
        If lResult > 0 Then
            DeleteFileEx = False
        Else
            DeleteFileEx = True
        End If
    Else
        DeleteFileEx = False
    End If
    'This will avoid empty error window to a
    '     ppear.
    Exit Function
DeleteFileEx_Error:
    'Show the Error Message with Error Number and its Description.
    MsgBox DEL_ERROR & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, App.ProductName & "Error Delete File"
    Exit Function
End Function

Public Sub CopyFile(filename As String, ToDir As String)
    Dim FileStruct As SHFILEOPSTRUCT
    Dim x As Long
    Dim P As Boolean
    Dim strNoConfirm As Integer, strNoConfirmMakeDir As Integer, strRenameOnCollision As Integer
    Dim strSilent As Integer, strSimpleProgress As Integer
    If NoConfirm = True Then
        strNoConfirm = FOF_NOCONFIRMATION
    Else
        strNoConfirm = 0
    End If
    If NoConfirmMakeDir = True Then
        strNoConfirmMakeDir = FOF_NOCONFIRMMKDIR
    Else
        strNoConfirmMakeDir = 0
    End If
    If RenameOnCollision = True Then
        strRenameOnCollision = FOF_RENAMEONCOLLISION
    Else
        strRenameOnCollision = 0
    End If
    If Silent = True Then
        strSilent = FOF_SILENT
    Else
        strSilent = 0
    End If
    If SimpleProgress = True Then
        strSimpleProgress = FOF_SIMPLEPROGRESS
    Else
        strSimpleProgress = 0
    End If
    
    P = FileExists(filename)
    If P = True Then
        FileStruct.pFrom = filename
        FileStruct.pTo = ToDir
        FileStruct.fFlags = strNoConfirm + strNoConfirmMakeDir + strRenameOnCollision + strSilent + strSimpleProgress
                
        FileStruct.wFunc = FO_COPY
        x = SHFileOperation(FileStruct)
    Else
        Err.Raise COPY_ERROR, App.ProductName & "Error Copy File", Err.Description
        Exit Sub
    End If
End Sub

'Delete Forever
Public Sub DeleteFile(filename As String)
    Dim FileStruct As SHFILEOPSTRUCT
    Dim x As Long
    Dim P As Boolean
    Dim strNoConfirm As Integer, strNoConfirmMakeDir As Integer, strRenameOnCollision As Integer
    Dim strSilent As Integer, strSimpleProgress As Integer
    If NoConfirm = True Then
        strNoConfirm = FOF_NOCONFIRMATION
    Else
        strNoConfirm = 0
    End If
    If Silent = True Then
        strSilent = FOF_SILENT
    Else
        strSilent = 0
    End If
    If SimpleProgress = True Then
        strSimpleProgress = FOF_SIMPLEPROGRESS
    Else
        strSimpleProgress = 0
    End If
    
    P = FileExists(filename)
    If P = True Then
        FileStruct.pFrom = filename
        FileStruct.fFlags = strNoConfirm + strSilent + strSimpleProgress
        FileStruct.wFunc = FO_DELETE
        x = SHFileOperation(FileStruct)
    Else
        Err.Raise DEL_ERROR, App.ProductName & "Error Delete File", Err.Description
        Exit Sub
    End If
End Sub

Public Sub MoveFile(filename As String, DestName As String)
    Dim FileStruct As SHFILEOPSTRUCT
    Dim P As Boolean
    Dim x As Long
    Dim strNoConfirm As Integer, strNoConfirmMakeDir As Integer, strRenameOnCollision As Integer
    Dim strSilent As Integer, strSimpleProgress As Integer
    If NoConfirm = True Then
        strNoConfirm = FOF_NOCONFIRMATION
    Else
        strNoConfirm = 0
    End If
    If NoConfirmMakeDir = True Then
        strNoConfirmMakeDir = FOF_NOCONFIRMMKDIR
    Else
        strNoConfirmMakeDir = 0
    End If
    If RenameOnCollision = True Then
        strRenameOnCollision = FOF_RENAMEONCOLLISION
    Else
        strRenameOnCollision = 0
    End If
    If Silent = True Then
        strSilent = FOF_SILENT
    Else
        strSilent = 0
    End If
    If SimpleProgress = True Then
        strSimpleProgress = FOF_SIMPLEPROGRESS
    Else
        strSimpleProgress = 0
    End If
    
    P = FileExists(filename)
    If P = True Then
        FileStruct.pFrom = filename
        FileStruct.pTo = DestName
     
        FileStruct.fFlags = strNoConfirm + strNoConfirmMakeDir + strRenameOnCollision + strSilent + strSimpleProgress
        FileStruct.wFunc = FO_MOVE
        x = SHFileOperation(FileStruct)
    Else
        Err.Raise MOVE_ERROR, App.ProductName & "Error Move File", Err.Description
        Exit Sub
    End If

End Sub
