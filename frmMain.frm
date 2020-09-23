VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1365
   ClientLeft      =   3870
   ClientTop       =   4470
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   91
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.CommandButton cmdDiskIO 
      Caption         =   "Save As Icon ..."
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   1350
   End
   Begin VB.CommandButton cmdDiskIO 
      Caption         =   "Open Bitmap ..."
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   1350
   End
   Begin VB.PictureBox picBitmap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1500
      MousePointer    =   2  'Cross
      Picture         =   "frmMain.frx":080A
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   540
      Width           =   750
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1500
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1500
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label1 
      Caption         =   "Transparent Color"
      Height          =   315
      Left            =   45
      TabIndex        =   6
      Top             =   135
      Width           =   1350
   End
   Begin VB.Label lblTransparency 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1485
      TabIndex        =   5
      Top             =   60
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH                      As Long = 260&
Private Const OFN_ALLOWMULTISELECT          As Long = &H200
Private Const OFN_CREATEPROMPT              As Long = &H2000
Private Const OFN_ENABLEHOOK                As Long = &H20
Private Const OFN_ENABLETEMPLATE            As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE      As Long = &H80
Private Const OFN_EXPLORER                  As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT        As Long = &H400
Private Const OFN_FILEMUSTEXIST             As Long = &H1000
Private Const OFN_HIDEREADONLY              As Long = &H4
Private Const OFN_LONGNAMES                 As Long = &H200000
Private Const OFN_NOCHANGEDIR               As Long = &H8
Private Const OFN_NODEREFERENCELINKS        As Long = &H100000
Private Const OFN_NOLONGNAMES               As Long = &H40000
Private Const OFN_NONETWORKBUTTON           As Long = &H20000
Private Const OFN_NOREADONLYRETURN          As Long = &H8000& 'see comments
Private Const OFN_NOTESTFILECREATE          As Long = &H10000
Private Const OFN_NOVALIDATE                As Long = &H100
Private Const OFN_OVERWRITEPROMPT           As Long = &H2
Private Const OFN_PATHMUSTEXIST             As Long = &H800
Private Const OFN_READONLY                  As Long = &H1
Private Const OFN_SHAREAWARE                As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH          As Long = 2
Private Const OFN_SHAREWARN                 As Long = 0
Private Const OFN_SHARENOWARN               As Long = 1
Private Const OFN_SHOWHELP                  As Long = &H10
Private Const OFS_MAXPATHNAME               As Long = 260
Private Const PICTYPE_BITMAP                As Long = 1
Private Const PICTYPE_ICON                  As Long = 3

Private Const DEFAULT_OPEN_FLAGS = OFN_EXPLORER _
                                   Or OFN_LONGNAMES _
                                   Or OFN_CREATEPROMPT _
                                   Or OFN_NODEREFERENCELINKS

Private Const DEFAULT_SAVE_FLAGS = OFN_EXPLORER _
                                   Or OFN_LONGNAMES _
                                   Or OFN_OVERWRITEPROMPT _
                                   Or OFN_HIDEREADONLY

Private Type IconInfo
    fIcon               As Long
    xHotspot            As Long
    yHotspot            As Long
    hBMMask             As Long
    hBMColor            As Long
End Type

Private Type Guid
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type

Private Type PictureInfo
    cbSizeofStruct      As Long
    picType             As Long
    hImage              As Long
    xExt                As Long
    yExt                As Long
End Type

Private Type OPENFILENAME
    nStructSize       As Long
    hWndOwner         As Long
    hInstance         As Long
    sFilter           As String
    sCustomFilter     As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    sFile             As String
    nMaxFile          As Long
    sFileTitle        As String
    nMaxTitle         As Long
    sInitialDir       As String
    sDialogTitle      As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    sDefFileExt       As String
    nCustData         As Long
    fnHook            As Long
    sTemplateName     As String
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (icoinfo As IconInfo) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictureInfo As PictureInfo, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private mintHeight                  As Integer
Private mintWidth                   As Integer
Private mlngTransparent             As Long

Private Sub picBitmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTransparency.BackColor = picBitmap.Point(X, Y)
End Sub

Private Sub picBitmap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngTransparent = picBitmap.Point(X, Y)
End Sub

Private Sub picBitmap_Resize()
    ScaleImage
End Sub

Private Sub cmdDiskIO_Click(Index As Integer)
    On Error Resume Next
    Dim strFile                 As String
    Dim udtFileIO               As OPENFILENAME
    Dim strPath                 As String
    Static sstrFileName         As String
    With udtFileIO
        .nStructSize = Len(udtFileIO)
        .hWndOwner = hWnd
        .sDialogTitle = cmdDiskIO(Index).Caption
        .nFilterIndex = 1
        If Index = 0 Then
            .sFilter = "Bitmaps (*.bmp)" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar
            .sDefFileExt = "bmp" & vbNullChar & vbNullChar
            .sFileTitle = "*.bmp" & vbNullChar & Space$(512) & vbNullChar & vbNullChar
            .sFile = "*.bmp" & Space$(1024) & vbNullChar & vbNullChar
            .sInitialDir = GetSetting(App.EXEName, "FileIO", "InitialBitmapDir", "C:\") & vbNullChar & vbNullChar
            .nMaxFile = Len(.sFile)
            .nMaxTitle = Len(udtFileIO.sFileTitle)
            .flags = DEFAULT_OPEN_FLAGS
            If GetOpenFileName(udtFileIO) = 0 Then Exit Sub
            cmdDiskIO(1).Enabled = True
        Else
            .sFilter = "Icon (*.ico)" & vbNullChar & "*.ico" & vbNullChar & vbNullChar
            .sDefFileExt = "ico" & vbNullChar & vbNullChar
            .sFileTitle = "*.ico" & vbNullChar & Space$(512) & vbNullChar & vbNullChar
            .sFile = "*.ico" & Space$(1024) & vbNullChar & vbNullChar
            .sInitialDir = GetSetting(App.EXEName, "FileIO", "InitialIconDir", "C:\") & vbNullChar & vbNullChar
            .nMaxFile = Len(.sFile)
            .nMaxTitle = Len(udtFileIO.sFileTitle)
            .flags = DEFAULT_SAVE_FLAGS
            If GetSaveFileName(udtFileIO) = 0 Then Exit Sub
            cmdDiskIO(1).Enabled = False
        End If
    End With
    strFile = Trim$(Replace(udtFileIO.sFileTitle, vbNullChar, vbNullString))
    strPath = Trim$(Replace(udtFileIO.sFile, vbNullChar, vbNullString))
    If Index = 0 Then
        picBitmap.Picture = LoadPicture(strPath)
        ScaleImage
        SaveSetting App.EXEName, "FileIO", "InitialBitmapDir", Left$(strPath, Len(strPath) - Len(strFile) - 1)
    Else
        SaveIcon strPath
        SaveSetting App.EXEName, "FileIO", "InitialIconDir", Left$(strPath, Len(strPath) - Len(strFile) - 1)
    End If
End Sub

Private Sub Form_Load()
    ScaleImage
End Sub

Private Sub ScaleImage()
    mintHeight = picBitmap.Height
    mintWidth = picBitmap.Width
    mintHeight = mintHeight - (mintHeight Mod 16)
    mintWidth = mintWidth - (mintWidth Mod 16)
    picImage.Move picBitmap.Left, picBitmap.Top, picBitmap.Width, picBitmap.Height
    picMask.Move picBitmap.Left, picBitmap.Top, picBitmap.Width, picBitmap.Height
End Sub

Private Sub SaveIcon(ByVal strPath As String)
    On Error Resume Next
    Dim lngCurrentColor         As Long
    Dim intCurrentX             As Long
    Dim intCurrentY             As Long
    Dim udtGuid                 As Guid
    Dim udtIconInfo             As IconInfo
    Dim lngNewBitmap            As Long
    Dim lngNewBitmapDC          As Long
    Dim lngNewMask              As Long
    Dim lngNewMaskDC            As Long
    Dim objPicture              As IPicture
    Dim udtPictureInfo          As PictureInfo
    Dim lngPreviousBitmap       As Long
    Dim lngPreviousMask         As Long
    picImage.Picture = LoadPicture()
    picMask.Picture = LoadPicture()
    picImage.Line (0, 0)-(mintWidth, mintHeight), vbBlack, BF
    picMask.Line (0, 0)-(mintWidth, mintHeight), vbWhite, BF
    For intCurrentY = 0 To mintHeight
        For intCurrentX = 0 To mintWidth
            lngCurrentColor = picBitmap.Point(intCurrentX, intCurrentY)
            If Not lngCurrentColor = mlngTransparent Then
                 picImage.PSet (intCurrentX, intCurrentY), lngCurrentColor
                 picMask.PSet (intCurrentX, intCurrentY), vbBlack
            End If
        Next
    Next
    lngNewMaskDC = CreateCompatibleDC(hDC)
    lngNewMask = CreateCompatibleBitmap(lngNewMaskDC, mintWidth, mintHeight)
    lngPreviousMask = SelectObject(lngNewMaskDC, lngNewMask)
    BitBlt lngNewMaskDC, 0, 0, mintWidth, mintHeight, picMask.hDC, 0, 0, vbSrcCopy
    SelectObject lngNewMaskDC, lngPreviousMask
    
    lngNewBitmapDC = CreateCompatibleDC(0)
    lngNewBitmap = CreateCompatibleBitmap(picBitmap.hDC, mintWidth, mintHeight)
    lngPreviousBitmap = SelectObject(lngNewBitmapDC, lngNewBitmap)
    BitBlt lngNewBitmapDC, 0, 0, mintWidth, mintHeight, picImage.hDC, 0, 0, vbSrcCopy
    SelectObject lngNewBitmapDC, lngPreviousBitmap
    With udtGuid
         .Data1 = &H20400
         .Data4(0) = &HC0
         .Data4(7) = &H46
    End With
    With udtIconInfo
        .fIcon = 1
        .hBMColor = lngNewBitmap
        .hBMMask = lngNewMask
        .xHotspot = mintWidth * 0.5
        .yHotspot = mintHeight * 0.5
    End With
    With udtPictureInfo
        .cbSizeofStruct = Len(udtPictureInfo)
        .picType = PICTYPE_ICON
        .hImage = CreateIconIndirect(udtIconInfo)
    End With
    OleCreatePictureIndirect udtPictureInfo, udtGuid, 1, objPicture
    Set Icon = objPicture
    SavePicture objPicture, strPath
    DestroyIcon udtPictureInfo.hImage
    DeleteObject lngNewBitmap
    DeleteObject lngNewMask
    DeleteDC lngNewBitmapDC
    DeleteDC lngNewMaskDC
    DeleteObject lngPreviousBitmap
    DeleteObject lngPreviousMask
    Set objPicture = Nothing
End Sub
