VERSION 5.00
Begin VB.Form frmColPicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Picker"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.dFlatButton cmdClose 
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   3870
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Close"
   End
   Begin Project1.dFlatButton cmdCopy 
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   3420
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Copy Color Value"
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      Height          =   675
      Left            =   4710
      MousePointer    =   99  'Custom
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   114
      TabIndex        =   13
      Top             =   1515
      Width           =   1770
   End
   Begin VB.PictureBox RgbC 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   4650
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   3030
      Width           =   300
   End
   Begin VB.HScrollBar hsBar 
      Height          =   270
      Index           =   2
      Left            =   5010
      Max             =   255
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1440
   End
   Begin VB.PictureBox RgbC 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   4650
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   2715
      Width           =   300
   End
   Begin VB.HScrollBar hsBar 
      Height          =   270
      Index           =   1
      Left            =   5010
      Max             =   255
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2685
      Width           =   1440
   End
   Begin VB.PictureBox RgbC 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   4650
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   2400
      Width           =   300
   End
   Begin VB.HScrollBar hsBar 
      Height          =   270
      Index           =   0
      Left            =   5010
      Max             =   255
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1455
   End
   Begin VB.PictureBox pColor 
      Height          =   330
      Left            =   4710
      ScaleHeight     =   270
      ScaleWidth      =   1710
      TabIndex        =   6
      Top             =   945
      Width           =   1770
   End
   Begin VB.ComboBox cboType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4710
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   345
      Width           =   1770
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0,0,0"
      Top             =   4620
      Width           =   3540
   End
   Begin VB.PictureBox pZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4500
      Left            =   75
      MouseIcon       =   "frmColPicker.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   0
      Top             =   45
      Width           =   4500
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Mixer:"
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
      Index           =   2
      Left            =   4710
      TabIndex        =   14
      Top             =   1305
      Width           =   855
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Preview:"
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
      Index           =   1
      Left            =   4710
      TabIndex        =   5
      Top             =   735
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Format:"
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
      Index           =   0
      Left            =   4710
      TabIndex        =   4
      Top             =   75
      Width           =   960
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Value:"
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
      Index           =   3
      Left            =   90
      TabIndex        =   1
      Top             =   4665
      Width           =   885
   End
End
Attribute VB_Name = "frmColPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CanUpdate As Boolean

Private Sub GradBar(oRgb As TRGB)
Dim rVal As Integer
Dim gVal As Integer
Dim bVal As Long
Dim X As Integer
    'Draw a Grident bar of the color
    rVal = oRgb.Red
    gVal = oRgb.Green
    bVal = oRgb.Blue
    
    For X = 0 To (p1.ScaleWidth - 1)
        rVal = (rVal + 1)
        If (rVal > 255) Then rVal = 255
        gVal = (gVal + 1)
        If (gVal > 255) Then gVal = 255
        bVal = (bVal + 1)
        If (bVal > 255) Then bVal = 255
        
        p1.Line (X, 0)-(X, p1.ScaleHeight), RGB(rVal, gVal, bVal)
    Next X
    
    p1.Refresh
End Sub

Private Function GetColorVal(RgbData As TRGB) As String
    'Return color value.
    Select Case cboType.ListIndex
        Case 0
            GetColorVal = RgbData.Red & "," & RgbData.Green & "," & RgbData.Blue
        Case 1
            GetColorVal = "&H" & Hex(RGB(RgbData.Red, RgbData.Green, RgbData.Blue)) & "&"
        Case 2
            GetColorVal = Dec2Web(RgbData)
        Case 3
            GetColorVal = "#" & RgbToHex(RgbData.Red, RgbData.Green, RgbData.Blue)
        Case 4
            GetColorVal = "Color(" & RgbData.Red & "," & RgbData.Green & "," & RgbData.Blue & ")"
        Case 5
            GetColorVal = "ColorTranslator.FromOle(" & RGB(RgbData.Red, RgbData.Green, RgbData.Blue) & ")"
    End Select
    
End Function

Private Sub cboType_Click()
Dim Tmp As TRGB
    Call Long2Rgb(pColor.BackColor, Tmp)
    txtValue.Text = GetColorVal(Tmp)
End Sub

Private Sub cmdClose_Click()
    Unload frmColPicker
End Sub

Private Sub cmdCopy_Click()
    Clipboard.SetText txtValue.Text
End Sub

Private Sub Form_Load()
Dim Ret As Long

    CanUpdate = True
    Set p1.MouseIcon = pZoom.MouseIcon
    'Zoom the image for previewing.
    Ret = StretchBlt(pZoom.hDC, 0, 0, pZoom.Width, pZoom.Height, GrabSrc.wHdc, _
    0, 0, GrabSrc.dWidth, GrabSrc.dHeight, vbSrcCopy)
    pZoom.Refresh
    
    'Add some color types
    cboType.AddItem "RGB"
    cboType.AddItem "VBHex"
    cboType.AddItem "HTML"
    cboType.AddItem "WebSafe"
    cboType.AddItem "Java"
    cboType.AddItem "DotNetOLE"
    cboType.ListIndex = 0
    
    pZoom_MouseDown vbLeftButton, 0, 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmColPicker = Nothing
End Sub

Private Sub hsBar_Change(Index As Integer)
Dim Tmp As TRGB
    pColor.BackColor = RGB(hsBar(0).Value, hsBar(1).Value, hsBar(2).Value)
    'Convert Long To Rgb
    Call Long2Rgb(pColor.BackColor, Tmp)
    txtValue.Text = GetColorVal(Tmp)
    'Draw Grident bar
    If (CanUpdate) Then
        Call GradBar(Tmp)
    End If
End Sub

Private Sub hsBar_Scroll(Index As Integer)
    Call hsBar_Change(Index)
End Sub

Private Sub ImgCursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call pZoom_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Tmp As TRGB
Dim oCol As Long

    If (Button = 1) Then
        oCol = GetPixel(p1.hDC, X, Y)
        CanUpdate = False 'Don't update this bar
        If (oCol <> -1) Then
            Call Long2Rgb(p1.Point(X, Y), Tmp)
            'Update scrollbar values.
            hsBar(0).Value = Tmp.Red
            hsBar(1).Value = Tmp.Green
            hsBar(2).Value = Tmp.Blue
        End If
    End If
End Sub

Private Sub p1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CanUpdate = True
End Sub

Private Sub pZoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call pZoom_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oCol As Long
Dim Tmp As TRGB
    If (Button = vbLeftButton) Then
        oCol = GetPixel(pZoom.hDC, X, Y)
        If (oCol <> -1) Then
            'Convert Long to RGB
            Call Long2Rgb(oCol, Tmp)
            'Update color in preview box.
            pColor.BackColor = RGB(Tmp.Red, Tmp.Green, Tmp.Blue)
            'Update textbox with color value.
            txtValue.Text = GetColorVal(Tmp)
            'Update scrollbars.
            hsBar(0).Value = Tmp.Red
            hsBar(1).Value = Tmp.Green
            hsBar(2).Value = Tmp.Blue
            'Draw Grident bar
            Call GradBar(Tmp)
        End If
    End If
End Sub

