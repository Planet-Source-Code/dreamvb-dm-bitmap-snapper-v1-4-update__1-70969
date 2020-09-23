Attribute VB_Name = "ModTools"
Option Explicit

Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal Hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetWindowDC Lib "user32.dll" (ByVal Hwnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Public Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PreviewPic
    wHdc As Long
    dWidth As Long
    dHeight As Long
End Type

Public Type TRGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Public wRect As RECT
Public pt As POINTAPI
Public GrabSrc As PreviewPic
Public CanUnload As Boolean

Function EnumChildWindow(ByVal Hwnd As Long, ByVal lParm As Long) As Long
Dim cName As String
Dim Ret As Long

    cName = Space(128)
    Ret = GetClassName(Hwnd, cName, 128)
    cName = Left(cName, InStr(cName, Chr(0)) - 1)
    
    If Len(cName) > 0 Then
        frmmain.cboCls.AddItem cName
        frmmain.cboCls.ItemData(frmmain.cboCls.ListCount - 1) = Hwnd
    End If
    
    cName = vbNullString


    EnumChildWindow = 1
End Function

Public Sub Long2Rgb(lColor As Long, RgbType As TRGB)
Dim Tmp As Long
    Tmp = lColor
    'Convert Long To RGB
    With RgbType
        .Red = (Tmp Mod &H100)
        Tmp = (Tmp \ &H100)
        .Green = (Tmp Mod &H100)
        Tmp = (Tmp \ &H100)
        .Blue = (Tmp Mod &H100)
    End With
End Sub

Public Function Dec2Web(hDecCol As TRGB) As String
Dim StrHex As String
Dim oColor As Long

    oColor = RGB(hDecCol.Red, hDecCol.Green, hDecCol.Blue)
    'Convert a long color to a HTML color
    StrHex = Hex(oColor)
    StrHex = StrHex & String(6 - Len(StrHex), "0")
    
    Dec2Web = "#" & Right(StrHex, 2) & Mid(StrHex, 3, 2) & Left(StrHex, 2)
    StrHex = vbNullString
    
End Function

Public Function RgbToHex(r, g, b) As String
Dim WebColor As OLE_COLOR
    WebColor = b + 256 * (g + 256 * r)
    'Format Hex to 6 places
    RgbToHex = Right$("000000" & Hex$(WebColor), 6)
End Function

Public Function FixPath(lPath As String)
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Sub CenterForm(Frm As Form)
    'Center a Form
    Frm.Left = (Screen.Width - Frm.Width) \ 2
    Frm.Top = (Screen.Height - Frm.Height) \ 2
End Sub
