Attribute VB_Name = "mMain"
Option Explicit

Public Const LB_FINDSTRING = &H18F

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hDC As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hDC As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, Image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, BITMAP As Long) As GpStatus
Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
End Type

Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)


