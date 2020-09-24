VERSION 5.00
Begin VB.Form Skin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   Picture         =   "fMain.frx":0000
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   1320
   End
   Begin VB.CommandButton max 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8910
      TabIndex        =   8
      Top             =   120
      Width           =   360
   End
   Begin VB.CommandButton min 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cloz 
      Caption         =   "Command1"
      Height          =   255
      Left            =   9270
      TabIndex        =   7
      Top             =   120
      Width           =   630
   End
   Begin VB.PictureBox BRp 
      Height          =   375
      Left            =   10320
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   7560
      Width           =   375
   End
   Begin VB.PictureBox blp 
      Height          =   255
      Left            =   0
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   7680
      Width           =   255
   End
   Begin VB.PictureBox tlp 
      Height          =   255
      Left            =   0
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Rightp 
      Height          =   4575
      Left            =   10440
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4515
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Leftp 
      Height          =   4575
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4515
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Topp 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   0
      Width           =   2775
   End
   Begin VB.PictureBox Bottomp 
      Height          =   255
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   195
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   7680
      Width           =   3255
   End
End
Attribute VB_Name = "Skin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const HTCAPTION = 2

Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const LWA_COLORKEY As Long = &H1

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_NORMAL As Long = 1
Private Const SW_SHOWMINIMIZED As Long = 2
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_MAXIMIZE As Long = 3
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const SW_SHOW As Long = 5
Private Const SW_MINIMIZE As Long = 6
Private Const SW_SHOWMINNOACTIVE As Long = 7
Private Const SW_SHOWNA As Long = 8
Private Const SW_RESTORE As Long = 9
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SW_FORCEMINIMIZE As Long = 11
Private Const SW_MAX As Long = 11

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cX As Long
    cY As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function EndDialog Lib "user32.dll" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hDCSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
Private Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim mDC As Long  ' Memory hDC
Dim mainBitmap As Long ' Memory Bitmap
Dim blendFunc32bpp As BLENDFUNCTION
Dim Token As Long ' Needed to close GDI+
Dim oldBitmap As Long

Public mfrm As Form

Private Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
  Dim R As RECT
Public mhwnd As Long
Public cLeft As New c32bppDIB, cRight As New c32bppDIB
Public TL As New c32bppDIB, TR As New c32bppDIB, cTop As New c32bppDIB
Public BL As New c32bppDIB, BR As New c32bppDIB, cBottom As New c32bppDIB
Public mx As Integer
Public my As Integer
Public Trans As Boolean
Dim clozO As Boolean
Dim maxO As Boolean
Dim minO As Boolean

Private Sub blp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If

End Sub

Private Sub Bottomp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
my = Y
Timer1 = False
End If
End Sub

Private Sub Bottomp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Bottomp.Top = Bottomp.Top + (Y - my) / 15
Height = Bottomp.Top * 15 + Bottomp.Height * 15
resign
End If
End Sub

Private Sub Bottomp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If
End Sub

Private Sub BRp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
my = Y
mx = X
Timer1 = False
End If
End Sub

Private Sub BRp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
BRp.Move BRp.Left + (X - mx) / 15, BRp.Top + (Y - my) / 15
Move Left, Top, BRp.Left * 15 + BRp.Width * 15, BRp.Top * 15 + BRp.Height * 15
resign
End If
End Sub

Private Sub Blp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
my = Y
mx = X
Timer1 = False
End If
End Sub

Private Sub Blp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
blp.Move blp.Left + (X - mx) / 15, blp.Top + (Y - my) / 15
Move Left + (blp.Left * 15 + blp.Width * 15), Top, Width - (blp.Left * 15 + blp.Width * 15), blp.Top * 15 + blp.Height * 15
blp.Left = 0
resign
End If
End Sub

Private Sub BRp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If
End Sub

Private Sub cloz_Click()
On Error Resume Next
Timer1 = False
SetParent mhwnd, hwnd
Unload Me
Unload mfrm
Set cLeft = Nothing
Set cRight = Nothing
Set TL = Nothing
Set TR = Nothing
Set cTop = Nothing
Set BL = Nothing
Set BR = Nothing
Set cBottom = Nothing


End Sub


Private Sub Form_Initialize()
  ' Start up GDI+
  Dim GpInput As GdiplusStartupInput
  GpInput.GdiplusVersion = 1
  If GdiplusStartup(Token, GpInput) <> 0 Then
    MsgBox "Error loading GDI+!", vbCritical
    Unload Me
  End If
  MakeTrans ""
End Sub

Private Sub Form_Load()
  Dim curWinLong As Long
  Dim tmpst As Long
  curWinLong = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
  SetWindowLong Me.hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
    GetWindowRect mhwnd, R
    Me.Width = R.Right - R.Left
    Me.Height = R.Bottom - R.Top
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Width <= 2985 Then Width = 2986
If Height <= 1380 Then Height = 1381
Leftp.Height = ScaleHeight
Rightp.Height = ScaleHeight
Rightp.Left = ScaleWidth - Rightp.Width
Topp.Width = ScaleWidth
Bottomp.Width = ScaleWidth
Bottomp.Top = ScaleHeight - Bottomp.Height
blp.Top = ScaleHeight - blp.Height
BRp.Left = ScaleWidth - BRp.Width
BRp.Top = ScaleHeight - BRp.Height
cloz.Left = ScaleWidth - 25 - cloz.Width
max.Left = ScaleWidth - 25 - cloz.Width - max.Width
min.Left = ScaleWidth - 25 - cloz.Width - max.Width - min.Width
MakeTrans ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cleanup everything
    Call GdiplusShutdown(Token)
    SelectObject mDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC mDC
End Sub

Private Function MakeTrans(posi As String) As Boolean
  Dim tempBI As BITMAPINFO
  Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
  Dim lngHeight As Long, lngWidth As Long
  Dim img As Long
  Dim graphics As Long
  Dim winSize As Size
  Dim srcPoint As POINTAPI
  
  With tempBI.bmiHeader
     .biSize = Len(tempBI.bmiHeader)
     .biBitCount = 32    ' Each pixel is 32 bit's wide
     .biHeight = Me.ScaleHeight  ' Height of the form
     .biWidth = Me.ScaleWidth    ' Width of the form
     .biPlanes = 1   ' Always set to 1
     .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
  End With
  mDC = CreateCompatibleDC(Me.hDC)
  mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
  oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
  
        TL.LoadPicture_File App.Path & "\Skin\T-L.png"
        BL.LoadPicture_File App.Path & "\Skin\B-L.png"
        BR.LoadPicture_File App.Path & "\Skin\B-R.png"
        cLeft.LoadPicture_File App.Path & "\Skin\left.png"
        cRight.LoadPicture_File App.Path & "\Skin\right.png"
        cTop.LoadPicture_File App.Path & "\Skin\Top.png"
        cBottom.LoadPicture_File App.Path & "\Skin\Bottom.png"
        
If posi = "Close" Then
        TR.LoadPicture_File App.Path & "\Skin\Close.png"
ElseIf posi = "Max" Then
        TR.LoadPicture_File App.Path & "\Skin\Max.png"
ElseIf posi = "Mini" Then
        TR.LoadPicture_File App.Path & "\Skin\Min.png"
Else
        TR.LoadPicture_File App.Path & "\Skin\T-R.png"
End If

  TL.Render mDC, 0, 0, 34, 53
  TR.Render mDC, Me.ScaleWidth - 124, 0, 124, 50
  BL.Render mDC, 0, Me.ScaleHeight - 29, 26, 29
  BR.Render mDC, Me.ScaleWidth - 32, Me.ScaleHeight - 33, 32, 33
    
    cLeft.Render mDC, 0, 52, 17, Me.ScaleHeight - 51 - 30
    cRight.Render mDC, Me.ScaleWidth - 21, 50, 21, Me.ScaleHeight - 50 - 33
    cTop.Render mDC, 34, 0, Me.ScaleWidth - 124 - 34, 41
    cBottom.Render mDC, 26, Me.ScaleHeight - 21, Me.ScaleWidth - 26 - 32, 21
    

  ' Needed for updateLayeredWindow call
  srcPoint.X = 0
  srcPoint.Y = 0
  winSize.cX = Me.ScaleWidth
  winSize.cY = Me.ScaleHeight
   
  With blendFunc32bpp
     .AlphaFormat = AC_SRC_ALPHA ' 32 bit
     .BlendFlags = 0
     .BlendOp = AC_SRC_OVER
     .SourceConstantAlpha = 255
  End With
   
  Call GdipDisposeImage(img)
  Call GdipDeleteGraphics(graphics)
  Call UpdateLayeredWindow(Me.hwnd, Me.hDC, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function


Private Sub leftp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mx = X
Timer1 = False
End If
End Sub

Private Sub leftp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Leftp.Left = Leftp.Left + (X - mx) / 15
Width = Width - (Leftp.Left * 15 + Leftp.Width * 15)
Left = Left + (Leftp.Left * 15 + Leftp.Width * 15)
Leftp.Left = 0
resign
End If
End Sub

Private Sub Leftp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If

End Sub

Private Sub max_Click()
If Me.WindowState = 2 Then
Me.WindowState = 0
Else
Me.WindowState = 2
End If
resign
TBS
End Sub

Private Sub min_Click()
mfrm.WindowState = 1
Me.Hide
End Sub

Private Sub Rightp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Me.Left = mfrm.Left - 17 * 15
Me.Top = mfrm.Top - 41 * 15
Me.Width = mfrm.Width - (-21 - 17) * 15
Me.Height = mfrm.Height - (-41 - 21) * 15
If mfrm.WindowState <> 1 Then
If Me.Visible = False Then
Me.Show
TBS
End If
End If
Dim pt As POINTAPI
Dim phwnd As Long
GetCursorPos pt
phwnd = WindowFromPoint(pt.X, pt.Y)
If phwnd = cloz.hwnd Then
If clozO = False Then
MakeTrans "Close"
clozO = True
Else
clozO = False
End If
ElseIf phwnd = max.hwnd Then
If maxO = False Then
MakeTrans "Max"
maxO = True
Else
maxO = False
End If
ElseIf phwnd = min.hwnd Then
If minO = False Then
MakeTrans "Mini"
minO = True
Else
minO = False
End If
Else
If maxO = True Or minO = True Or clozO = True Then
MakeTrans ""
End If
End If
End Sub

Private Sub tlp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
my = Y
mx = X
Timer1 = False
End If
End Sub

Private Sub tlp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
tlp.Move tlp.Left + (X - mx) / 15, tlp.Top + (Y - my) / 15
Move Left + (tlp.Left * 15 + tlp.Width * 15), Top + (tlp.Top * 15 + tlp.Height * 15), Width - (tlp.Left * 15 + tlp.Width * 15), Height - (tlp.Top * 15 + tlp.Height * 15)
tlp.Left = 0
tlp.Top = 0
resign
End If
End Sub

Private Sub rightp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
mx = X
Timer1 = False
End If
End Sub

Private Sub rightp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Rightp.Left = Rightp.Left + (X - mx) / 15
Width = Rightp.Left * 15 + Rightp.Width * 15
resign
End If
End Sub

Private Sub tlp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If

End Sub

Private Sub Topp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = False
mx = X
my = Y
If Trans = True Then
MakeTransparent mfrm.hwnd, 175
End If
End If
End Sub

Private Sub Topp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Me.Move Left + X - mx, Top + Y - my
resign
End If
End Sub

Private Sub Topp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Timer1 = True
End If
If Trans = True Then
MakeOpaque mfrm.hwnd
End If
End Sub
Sub resign()
mfrm.Left = Me.Left + 17 * 15
mfrm.Top = Me.Top + 41 * 15
mfrm.Width = Me.Width + (-17 - 21) * 15
mfrm.Height = Me.Height + (-41 - 21) * 15
End Sub
Function TBS()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, &H40)
End Function
