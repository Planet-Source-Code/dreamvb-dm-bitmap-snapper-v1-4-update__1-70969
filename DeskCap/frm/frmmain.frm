VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM Bitmap Snapper"
   ClientHeight    =   5160
   ClientLeft      =   -45
   ClientTop       =   180
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   Begin Project1.dFlatButton cmdPaste 
      Height          =   390
      Left            =   2175
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Paste"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":0000
   End
   Begin Project1.dFlatButton cmdEdit 
      Height          =   390
      Left            =   5565
      TabIndex        =   32
      ToolTipText     =   "Edit Image"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":0352
   End
   Begin VB.PictureBox pGrip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   6975
      Picture         =   "frmmain.frx":06A4
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   30
      Top             =   195
      Visible         =   0   'False
      Width           =   165
   End
   Begin Project1.Line3D Line3D4 
      Height          =   30
      Left            =   0
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   53
   End
   Begin Project1.Line3D Line3D3 
      Height          =   30
      Left            =   0
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1335
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboCls 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   975
      Width           =   2760
   End
   Begin VB.ComboBox cboWidth 
      Height          =   315
      Left            =   3375
      Style           =   2  'Dropdown List
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   555
      Width           =   1215
   End
   Begin Project1.dFlatButton cmdMove 
      Height          =   390
      Index           =   3
      Left            =   3915
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Move Up"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":077E
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboStep 
      Height          =   315
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   555
      Width           =   1215
   End
   Begin Project1.dFlatButton cmdZoom 
      Height          =   390
      Left            =   3450
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Zoom"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":0AD0
   End
   Begin Project1.dFlatButton cmdOpen 
      Height          =   390
      Left            =   870
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Open"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":0BE2
   End
   Begin Project1.dFlatButton cmdBorderCol 
      Height          =   390
      Left            =   3030
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Border Color"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":0F34
   End
   Begin Project1.dFlatButton cmdGrab 
      Height          =   390
      Left            =   15
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Grab Desktop"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":1046
   End
   Begin VB.Timer TmrGetWin 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   7140
      Top             =   2235
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7095
      Top             =   1755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox psBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   516
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4830
      Width           =   7740
      Begin VB.Label lblResSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2235
         TabIndex        =   31
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   90
         Width           =   45
      End
   End
   Begin VB.PictureBox pHolder 
      BackColor       =   &H8000000C&
      Height          =   3345
      Left            =   45
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1395
      Width           =   6795
      Begin VB.PictureBox pSpacer 
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   4800
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   570
         Width           =   270
      End
      Begin VB.HScrollBar hBar 
         Height          =   270
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1380
         Width           =   450
      End
      Begin VB.VScrollBar vBar 
         Height          =   465
         Left            =   4800
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox SrcDc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   112
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1680
         Begin VB.PictureBox PicPoint 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   105
            Left            =   525
            MousePointer    =   8  'Size NW SE
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   34
            Top             =   465
            Width           =   105
         End
         Begin VB.Image ImgMove 
            Height          =   480
            Left            =   0
            MousePointer    =   5  'Size
            Top             =   0
            Width           =   480
         End
         Begin VB.Shape shpMove 
            BorderColor     =   &H00000000&
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin Project1.dFlatButton cmdCpy 
      Height          =   390
      Left            =   1740
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Copy"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":1158
   End
   Begin Project1.dFlatButton cmdWin 
      Height          =   390
      Left            =   435
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Grab Window"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmmain.frx":14AA
      Caption         =   ""
      Picture         =   "frmmain.frx":160C
   End
   Begin Project1.dFlatButton cmdColPicker 
      Height          =   390
      Left            =   2610
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Color Picker"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":195E
   End
   Begin Project1.dFlatButton cmdSave 
      Height          =   390
      Left            =   1305
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Save"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":1A70
   End
   Begin Project1.Line3D Line3D2 
      Height          =   30
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   885
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   53
   End
   Begin Project1.dFlatButton cmdMove 
      Height          =   390
      Index           =   1
      Left            =   4320
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Move Down"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":1DC2
   End
   Begin Project1.dFlatButton cmdMove 
      Height          =   390
      Index           =   4
      Left            =   4725
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Move Left"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":2114
   End
   Begin Project1.dFlatButton cmdMove 
      Height          =   390
      Index           =   2
      Left            =   5130
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Move Right"
      Top             =   60
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "frmmain.frx":2466
   End
   Begin VB.Label lblCpy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Classname"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4125
      MouseIcon       =   "frmmain.frx":27B8
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label lblCls 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window Class:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   45
      TabIndex        =   25
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblGWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grab Width:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2340
      TabIndex        =   23
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Step:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   30
      TabIndex        =   16
      Top             =   600
      Width           =   915
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuTake 
         Caption         =   "&Grab Desktop..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Save Preview"
      End
      Begin VB.Menu mnu1Screen 
         Caption         =   "Save Screen Capture"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucpy 
         Caption         =   "&Copy Preview"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnu100 
            Caption         =   "Original"
         End
         Begin VB.Menu mnu200 
            Caption         =   "200%"
         End
         Begin VB.Menu mnu300 
            Caption         =   "300%"
         End
         Begin VB.Menu mnu400 
            Caption         =   "400%"
         End
         Begin VB.Menu mnu500 
            Caption         =   "500%"
         End
         Begin VB.Menu mnu600 
            Caption         =   "600%"
         End
      End
   End
   Begin VB.Menu mnuPrv 
      Caption         =   "Preview"
      Begin VB.Menu mnuShow 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuEdit1 
         Caption         =   "Edit Image"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucon 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnublank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnu1 
      Caption         =   "#1"
      Visible         =   0   'False
      Begin VB.Menu mnuSave1 
         Caption         =   "Preview"
      End
      Begin VB.Menu mnuSave2 
         Caption         =   "Screen Capture"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CanMove As Boolean
Private MoveStep As Integer
Private mZoomFactor As Integer
Private mCurHwnd As Long
Private mPrevHwnd As Long

Private WndCaption As String
Private mShowPreview As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cboCls_Click()
Dim mCurHwnd As Long
    'Get Windows Hwnd
    mCurHwnd = cboCls.ItemData(cboCls.ListIndex)
    'Take screenshot.
    Call TakeScreenShot(mCurHwnd, mZoomFactor)
    Call ShowPreview
    SrcDc.SetFocus
End Sub

Private Sub cboWidth_Click()
On Error Resume Next
    shpMove.BorderWidth = Val(cboWidth.Text)
    'Save Border width
    SaveSetting "DmBitmapSnap", "Cfg", "GrabWidth", cboWidth.ListIndex
    SrcDc.SetFocus
    'Show preview
    Call ShowPreview
End Sub

Private Sub cmdBorderCol_Click()
On Error GoTo ErrFlag:
    'Set border color of the shape.
    With CD1
        .CancelError = True
        .ShowColor
        'Sets the shape border color
        shpMove.BorderColor = .Color
        'Save value to regedit
        SaveSetting "DmBitmapSnap", "Cfg", "bColor", .Color
    End With
    
    Exit Sub
ErrFlag:
    If (Err.Number = cdlCancel) Then Err.Clear
    
End Sub

Private Sub cmdEdit_Click()
    Call mnuEdit1_Click
End Sub

Private Sub cmdMove_Click(Index As Integer)
    Call SrcDc_KeyDown((41 - Index), 0)
End Sub

Private Sub cmdMove_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CanMove = False
End Sub

Private Sub ZoomImage(ByVal Factor As Integer)
    Call TakeScreenShot(mCurHwnd, Factor)
    Call ShowPreview
    SrcDc.SetFocus
End Sub

Private Sub CleanExit()
    If (TmrGetWin.Enabled) Then
        TmrGetWin.Enabled = False
    End If
    
    Set SrcDc.Picture = Nothing
    cboStep.Clear
    cboWidth.Clear
    Unload frmmain
    Unload frmColPicker
End Sub

Private Sub SavePix(PicBox As PictureBox, Optional ByVal Title As String = "")
On Error GoTo SaveErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = "Bitmap Files(*.bmp)|*.bmp|"
        .ShowSave
        'Save Preview
        SavePicture PicBox.Image, .FileName
        .FileName = vbNullString
    End With
    
    Exit Sub
SaveErr:
    If (Err.Number = cdlCancel) Then Err.Clear

End Sub

Private Sub Mover(ByVal theleft As Long, ByVal TheTop As Long)
    If (GrabSrc.dWidth < 16) Then GrabSrc.dWidth = 16
    If (GrabSrc.dHeight < 16) Then GrabSrc.dHeight = 16
    
    ImgMove.Move theleft, TheTop, GrabSrc.dWidth, GrabSrc.dHeight
    shpMove.Move theleft, TheTop, GrabSrc.dWidth, GrabSrc.dHeight
    PicPoint.Move (ImgMove.Left + ImgMove.Width), (ImgMove.Top + ImgMove.Height), 7, 7
End Sub

Private Sub FixPosition()
    'This stops the controls been moved, of the area
    With ImgMove

        If (hBar.Value = 0) Then
            If (.Left + GrabSrc.dWidth) >= SrcDc.ScaleWidth Then
                .Left = (SrcDc.ScaleWidth - GrabSrc.dWidth)
            End If
        End If
        
        If (vBar.Value = 0) Then
            If (.Top + GrabSrc.dHeight) >= SrcDc.ScaleHeight Then
                .Top = (SrcDc.ScaleHeight - GrabSrc.dHeight)
            End If
        End If
        
        If (.Left - hBar.Value) <= 0 Then
            .Left = hBar.Value
        End If
        
        If (.Top - vBar.Value) <= 0 Then
            .Top = vBar.Value
        End If
        
        If ((.Top - vBar.Value) + GrabSrc.dHeight >= hBar.Top) Then
            .Top = (hBar.Top - GrabSrc.dHeight) + vBar.Value
        End If
        
        If ((.Left - hBar.Value) + GrabSrc.dWidth) > vBar.Left Then
            .Left = (vBar.Left - GrabSrc.dWidth) + hBar.Value
        End If
        
        If (.Left + GrabSrc.dWidth) >= SrcDc.ScaleWidth Then
            .Left = (SrcDc.ScaleWidth - GrabSrc.dWidth)
        End If
        
        If (.Top + GrabSrc.dHeight) >= SrcDc.ScaleHeight Then
            .Top = (SrcDc.ScaleHeight - GrabSrc.dHeight)
        End If
    End With
    
End Sub

Private Sub ShowPreview()
    'Move and resize preview
    With frmPreview
        .pPreview.Height = GrabSrc.dHeight
        .pPreview.Width = GrabSrc.dWidth
        
        .pPreview.Left = (.ScaleWidth - GrabSrc.dWidth) \ 2
        .pPreview.Top = (.ScaleHeight - GrabSrc.dHeight) \ 2
        .Caption = "Preview " & GrabSrc.dWidth & "x" & GrabSrc.dHeight
    End With
    
    'Show Preview.
    With frmPreview.pPreview
        .Cls
        BitBlt .hDC, 0, 0, GrabSrc.dWidth, GrabSrc.dHeight, SrcDc.hDC, ImgMove.Left, ImgMove.Top, vbSrcCopy
        .Refresh
    End With
End Sub

Private Sub SetScrollbars()
    vBar.Max = (SrcDc.ScaleHeight - pHolder.ScaleHeight) + 18
    hBar.Max = (SrcDc.Width - pHolder.ScaleWidth) + 18
    'Enable/Disable scrollbars.
    hBar.Enabled = (hBar.Max > 0)
    vBar.Enabled = (vBar.Max > 0)
End Sub

Private Sub TakeScreenShot(ByVal TheHwnd As Long, ByVal ZoomFactor As Integer)
Dim wHdc As Long
Dim Ret As Long
    
    'Take Screen Grab
    frmmain.Visible = False
    frmPreview.Visible = False
    DoEvents
    Sleep 100
 
    If (TheHwnd = 0) Then
        MsgBox "Unable to take screenshot.", vbInformation, frmmain.Caption
        Exit Sub
    Else
        'Get Window DC
        wHdc = GetWindowDC(TheHwnd)
        'Get Window Size
        Ret = GetWindowRect(TheHwnd, wRect)
    End If
    
    With SrcDc
        .Cls
        .Width = (wRect.Right - wRect.Left) * ZoomFactor
        .Height = (wRect.Bottom - wRect.Top) * ZoomFactor
        lblResSize.Caption = .Width & " x " & .Height & " pixels"
        StretchBlt .hDC, 0, 0, .Width, .Height, wHdc, 0, 0, (.Width \ ZoomFactor), (.Height \ ZoomFactor), vbSrcCopy
        'Set scrollbars
        Call SetScrollbars
        .Refresh
    End With

    frmmain.Visible = True
    frmPreview.Visible = mShowPreview
           
End Sub

Private Sub cboStep_Click()
On Error Resume Next
    SaveSetting "DmBitmapSnap", "Cfg", "MoveStep", cboStep.ListIndex
    MoveStep = (cboStep.ListIndex + 1)
    SrcDc.SetFocus
End Sub

Private Sub cmdColPicker_Click()
    GrabSrc.dHeight = frmPreview.pPreview.Height
    GrabSrc.dWidth = frmPreview.pPreview.Width
    GrabSrc.wHdc = frmPreview.pPreview.hDC
    'Show color picker form
    frmColPicker.Show vbModal, frmmain
End Sub

Private Sub cmdCpy_Click()
    'Copy preview to clipboad
    Clipboard.Clear
    Clipboard.SetData frmPreview.pPreview.Image, vbCFBitmap
End Sub

Private Sub cmdGrab_Click()
    mCurHwnd = GetDesktopWindow()
    Call TakeScreenShot(mCurHwnd, mZoomFactor)
    Call ShowPreview
    SrcDc.SetFocus
End Sub

Private Sub cmdOpen_Click()
On Error GoTo OpenErr:

    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Picture Files(*.bmp;*.jpg;*.jpeg;*.gif)|*.bmp;*.jpg;*.gif;*.jpeg|"
        .ShowOpen
        Set SrcDc.Picture = Nothing
        vBar.Value = 0
        hBar.Value = 0
        SrcDc.Picture = LoadPicture(.FileName)
        'Set scrollbars
        Call SetScrollbars
        .FileName = vbNullString
    End With
    
    Exit Sub

OpenErr:
    If (Err.Number = cdlCancel) Then Err.Clear
    
End Sub

Private Sub cmdPaste_Click()
    'Paste image from clipboard
    SrcDc.Picture = Clipboard.GetData(vbCFBitmap)
    Call SetScrollbars
End Sub

Private Sub cmdSave_Click()
    PopupMenu mnu1, , 85, 30
End Sub

Private Sub cmdWin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        TmrGetWin.Enabled = True
    End If
End Sub

Private Sub cmdWin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (TmrGetWin.Enabled) And (Button = vbLeftButton) Then
        cmdWin.MousePointer = vbCustom
    End If
End Sub

Private Sub cmdWin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ret As Long

    If (Button = vbLeftButton) Then
        TmrGetWin.Enabled = False
        cmdWin.MousePointer = vbNormal
        If Len(WndCaption) > 0 Then
            'Clear comboxbox.
            cboCls.Clear
            cboCls.AddItem WndCaption
            cboCls.ItemData(0) = mCurHwnd
            'Fill combo box with parent child windows.
            Ret = EnumChildWindows(mCurHwnd, AddressOf EnumChildWindow, 0)
            '
            If (cboCls.ListCount > 0) Then
                cboCls.ListIndex = 0
            End If
            
            cboCls.Enabled = (cboCls.ListCount) > 0
        End If
    End If
    
    WndCaption = vbNullString
End Sub

Private Sub cmdZoom_Click()
    PopupMenu mnuZoom, , 200, 30
End Sub


Private Sub Form_Load()
Dim I As Integer
On Error Resume Next
    'Center this form
    Call CenterForm(frmmain)
    'Hwnd of Desktop
    mCurHwnd = GetDesktopWindow
    'Set zoom factor
    mZoomFactor = 1
    
    'Add some Move steps
    For I = 1 To 16
        cboStep.AddItem "Step " & I
    Next I
    
    'Add Grabber border widths
    cboWidth.AddItem "1"
    cboWidth.AddItem "2"
    cboWidth.AddItem "3"
    cboWidth.AddItem "4"
    
    'Set combo index's
    cboStep.ListIndex = Val(GetSetting("DmBitmapSnap", "Cfg", "MoveStep", 0))
    cboWidth.ListIndex = Val(GetSetting("DmBitmapSnap", "Cfg", "GrabWidth", 0))
    'Read the shapes border color value.
    shpMove.BorderColor = Val(GetSetting("DmBitmapSnap", "Cfg", "bColor", 0))
    'Read grabber width and height
    GrabSrc.dWidth = Val(GetSetting("DmBitmapSnap", "Cfg", "gWidth", 32))
    GrabSrc.dHeight = Val(GetSetting("DmBitmapSnap", "Cfg", "gHeight", 32))
    
    'Position the grabber
    Call Mover(0, 0)
    Call cmdGrab_Click
    Call mnuShow_Click
    frmPreview.Show , frmmain
    'Align preview form.
    frmPreview.Left = (frmmain.Left + frmmain.Width) + Screen.TwipsPerPixelX
    frmPreview.Top = frmmain.Top
    '
    PicPoint.Line (0, 0)-(PicPoint.ScaleWidth - 8, PicPoint.ScaleHeight - 8), vbWhite, B
    PicPoint.Refresh
    '
    Call Form_Resize
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCpy.FontUnderline = False
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Line3D1.Width = frmmain.ScaleWidth
    Line3D2.Width = Line3D1.Width
    Line3D3.Width = Line3D1.Width
    Line3D4.Width = Line3D1.Width
    
    pHolder.Width = (frmmain.ScaleWidth - pHolder.Left)
    pHolder.Height = (frmmain.ScaleHeight - psBar.Height) - pHolder.Top - 1
    
    'Position scrollbars
    vBar.Left = (pHolder.ScaleWidth - vBar.Width)
    vBar.Height = (pHolder.ScaleHeight) - 18
    
    hBar.Top = (pHolder.ScaleHeight - hBar.Height)
    hBar.Width = (pHolder.ScaleWidth - vBar.Width)
    
    'Position Spacer
    pSpacer.Left = (pHolder.ScaleWidth - pSpacer.Width)
    pSpacer.Top = (pHolder.ScaleHeight - pSpacer.Height)
    'Setup scrollbars
    Call SetScrollbars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save Grabber width and height
    SaveSetting "DmBitmapSnap", "Cfg", "gWidth", GrabSrc.dWidth
    SaveSetting "DmBitmapSnap", "Cfg", "gHeight", GrabSrc.dHeight
        
    CanUnload = True
    Set frmColPicker = Nothing
    Set frmmain = Nothing
End Sub

Private Sub hBar_Change()
    SrcDc.Left = -hBar.Value
    
    If (ImgMove.Left - hBar.Value) <= 0 Then
        Mover hBar.Value, ImgMove.Top
    End If

    If ((ImgMove.Left - hBar.Value) + GrabSrc.dWidth) > vBar.Left Then
        Call Mover((vBar.Left - GrabSrc.dWidth) + hBar.Value, ImgMove.Top)
    End If
    
End Sub

Private Sub hBar_Scroll()
    Call hBar_Change
End Sub

Private Sub ImgMove_Click()
    SrcDc.SetFocus
End Sub

Private Sub ImgMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        wRect.Left = X
        wRect.Top = Y
        CanMove = True
    End If
End Sub

Private Sub ImgMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) And (CanMove) Then
        With ImgMove
            .Left = (.Left + (X - wRect.Left) \ Screen.TwipsPerPixelX)
            .Top = (.Top + (Y - wRect.Top) \ Screen.TwipsPerPixelY)
            'Fix Grabber position.
            Call FixPosition
            'Position mover
            Call Mover(.Left, .Top)
        End With
    End If
End Sub

Private Sub ImgMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Call ShowPreview
        CanMove = False
    End If
End Sub

Private Sub lblCpy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        lblCpy.ForeColor = vbRed
        Clipboard.Clear
        Clipboard.SetText cboCls.Text, vbCFText
    End If
End Sub

Private Sub lblCpy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCpy.FontUnderline = True
End Sub

Private Sub lblCpy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCpy.ForeColor = vbBlue
End Sub

Private Sub mnu100_Click()
    Call ZoomImage(1)
End Sub

Private Sub mnu1Screen_Click()
    Call mnuSave2_Click
End Sub

Private Sub mnu200_Click()
    Call ZoomImage(2)
End Sub

Private Sub mnu300_Click()
    Call ZoomImage(3)
End Sub

Private Sub mnu400_Click()
    Call ZoomImage(4)
End Sub

Private Sub mnu500_Click()
    Call ZoomImage(5)
End Sub

Private Sub mnu600_Click()
    Call ZoomImage(6)
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & "Ver 1.4" & vbCrLf & vbTab & "By Ben Jones" _
    & vbCrLf & vbTab & vbTab & "Please vote if you like this code.", vbInformation, "About"
End Sub

Private Sub mnucon_Click()
Dim Ret As Long

    Ret = ShellExecute(frmmain.Hwnd, "open", _
    FixPath(App.Path) & "Help.rtf", vbNullString, vbNullString, 3)
    
    If (Ret = 2) Then
        MsgBox "Help file was not found", vbExclamation, "File Not Found"
    End If
    
End Sub

Private Sub mnucpy_Click()
    Call cmdCpy_Click
End Sub

Private Sub mnuEdit1_Click()
Dim TmpFile As String
Dim Ret As Long
    'Open the preview image in the systems image editing program.
    TmpFile = FixPath(App.Path) & "preview.bmp"
    SavePicture frmPreview.pPreview.Image, TmpFile
    Ret = ShellExecute(frmmain.Hwnd, "edit", TmpFile, vbNullString, vbNullString, 3)
End Sub

Private Sub mnuExit_Click()
    Call CleanExit
End Sub

Private Sub mnuPreview_Click()
    Call mnuSave1_Click
End Sub

Private Sub mnuSave1_Click()
    Call SavePix(frmPreview.pPreview, "Save Preview")
End Sub

Private Sub mnuSave2_Click()
    Call SavePix(SrcDc, "Save Screen Capture")
End Sub

Private Sub mnuShow_Click()
    mShowPreview = (Not mShowPreview)
    frmPreview.Visible = mShowPreview
    
    If (mShowPreview) Then
        mnuShow.Caption = "Hide"
    Else
        mnuShow.Caption = "Show"
    End If
    
End Sub

Private Sub mnuTake_Click()
    Call cmdGrab_Click
End Sub

Private Sub PicPoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        CanMove = True
        wRect.Left = X
        wRect.Top = Y
    End If
End Sub

Private Sub PicPoint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton And CanMove) Then
        'Position Resize grip
        PicPoint.Left = (PicPoint.Left - (wRect.Left - X) \ Screen.TwipsPerPixelX)
        PicPoint.Top = (PicPoint.Top - (wRect.Top - Y) \ Screen.TwipsPerPixelY)
        'Store grabber width and height
        wRect.Right = (PicPoint.Left - shpMove.Left)
        wRect.Bottom = (PicPoint.Top - shpMove.Top)
        'Make sure grabber does not go below 16x16
        If (wRect.Right < 16) Then wRect.Right = 16
        If (wRect.Bottom < 16) Then wRect.Bottom = 16
        'Position grabber ressize grip
        PicPoint.Move (ImgMove.Left + ImgMove.Width), (ImgMove.Top + ImgMove.Height), 7, 7
        'Resize Shape
        shpMove.Width = wRect.Right
        shpMove.Height = wRect.Bottom
        'Resize Image
        ImgMove.Width = shpMove.Width
        ImgMove.Height = shpMove.Height
                'Update preview caption
        frmPreview.Caption = "Preview " & wRect.Right & "x" & wRect.Bottom
    End If
End Sub

Private Sub PicPoint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (CanMove) Then
        CanMove = False
        'Resize the grabber tool
        GrabSrc.dHeight = shpMove.Height
        GrabSrc.dWidth = shpMove.Width
        'Display preview.
        Call ShowPreview
        SrcDc.SetFocus
    End If
End Sub

Private Sub psBar_Resize()
Dim Xpos As Integer
Dim YPos As Integer

    'Position Grip Image
    Xpos = (psBar.ScaleWidth - 11)
    YPos = (psBar.ScaleHeight - 11)
    'Clear DC
    psBar.Cls
    TransparentBlt psBar.hDC, Xpos, YPos, 11, 11, pGrip.hDC, 0, 0, 11, 11, vbMagenta
    psBar.Refresh
    'Position Grabed image size
    lblResSize.Left = (psBar.ScaleWidth - lblResSize.Width) - 30
    
End Sub

Private Sub SrcDc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = vbKeyUp And Shift) Then
        GrabSrc.dHeight = (GrabSrc.dHeight - 1)
        Call Mover(ImgMove.Left, ImgMove.Top)
        'Update preview
        Call ShowPreview
        KeyCode = 0
    End If
    
    If (KeyCode = vbKeyDown And Shift) Then
        GrabSrc.dHeight = (GrabSrc.dHeight + 1)
        Call Mover(ImgMove.Left, ImgMove.Top)
        'Update preview
        Call ShowPreview
        KeyCode = 0
    End If
    
    If (KeyCode = vbKeyLeft And Shift) Then
        GrabSrc.dWidth = (GrabSrc.dWidth - 1)
        Call Mover(ImgMove.Left, ImgMove.Top)
        'Update preview
        Call ShowPreview
        KeyCode = 0
    End If
    
    If (KeyCode = vbKeyRight And Shift) Then
        GrabSrc.dWidth = (GrabSrc.dWidth + 1)
        Call Mover(ImgMove.Left, ImgMove.Top)
        'Update preview
        Call ShowPreview
        KeyCode = 0
    End If
    
    'Movement Positions
    Select Case KeyCode
        Case vbKeyUp
            ImgMove.Top = (ImgMove.Top - MoveStep)
        Case vbKeyDown
            ImgMove.Top = (ImgMove.Top + MoveStep)
        Case vbKeyLeft
            ImgMove.Left = (ImgMove.Left - MoveStep)
        Case vbKeyRight
            ImgMove.Left = (ImgMove.Left + MoveStep)
        Case vbKeyEnd
            vBar.Value = vBar.Max
        Case vbKeyHome
            vBar.Value = 0
    End Select
    
    'Fix Grabber position.
    Call FixPosition
    
    If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or _
    (KeyCode = vbKeyHome) Or (KeyCode = vbKeyEnd) Then
        'Position the grabber
        shpMove.Move ImgMove.Left, ImgMove.Top, GrabSrc.dWidth, GrabSrc.dHeight
        'Position Resize grip
        PicPoint.Move (ImgMove.Left + GrabSrc.dWidth), (ImgMove.Top + GrabSrc.dHeight), 7, 7
        'Show Preview
        Call ShowPreview
    End If

    SrcDc.SetFocus
End Sub

Private Sub SrcDc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Call Mover((X - GrabSrc.dWidth \ 2), (Y - GrabSrc.dHeight \ 2))
        '
        Call ShowPreview
    End If
End Sub

Private Sub SrcDc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SrcDc.MousePointer = vbDefault
End Sub

Private Sub SrcDc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then SrcDc.MousePointer = vbSizeAll
End Sub

Private Sub TmrGetWin_Timer()
Dim Ret As Long
Dim wLen As Long
Dim wDc As Long

    'Get cursor pos
    Ret = GetCursorPos(pt)
    'Get Hwnd from cursor pos
    mCurHwnd = WindowFromPoint(pt.X, pt.Y)
    
    If (mCurHwnd) Then
        'Get Window textlength
        wLen = GetWindowTextLength(mCurHwnd) + 1
        'Check text length
        If (wLen > 0) Then
            WndCaption = Space(wLen)
            Ret = GetWindowText(mCurHwnd, WndCaption, wLen)
            'Update window display info
            lblCaption.Caption = Left(WndCaption, InStr(WndCaption, Chr(0)) - 1)
            'Get window rect
            Ret = GetWindowRect(mCurHwnd, wRect)
            'Update Grabed image size label
            wDc = GetWindowDC(mCurHwnd)
            
            'InvertRect wDc, wRect
            
            lblResSize.Caption = (wRect.Right - wRect.Left) & " x " & (wRect.Bottom - wRect.Top) & " pixels"
            '
            mPrevHwnd = mCurHwnd
        End If
    End If
End Sub

Private Sub vBar_Change()
    SrcDc.Top = -vBar.Value
    
    If (ImgMove.Top - vBar.Value) <= 0 Then
        Mover ImgMove.Left, (vBar.Value + 3)
    End If
    
    If ((ImgMove.Top - vBar.Value) + GrabSrc.dHeight >= hBar.Top) Then
        Mover ImgMove.Left, (hBar.Top - GrabSrc.dHeight) + vBar.Value
    End If
    
End Sub

Private Sub vBar_Scroll()
    vBar_Change
End Sub
