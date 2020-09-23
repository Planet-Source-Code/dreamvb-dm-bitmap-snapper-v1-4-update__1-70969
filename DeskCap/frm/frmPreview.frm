VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   1515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbase 
      Height          =   1020
      Left            =   0
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin VB.PictureBox pPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   255
         ScaleHeight     =   540
         ScaleWidth      =   630
         TabIndex        =   1
         Top             =   270
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    pbase.Width = frmPreview.ScaleWidth
    pbase.Height = frmPreview.ScaleHeight
    
    pPreview.Left = (pbase.ScaleWidth - pPreview.Width) \ 2
    pPreview.Top = (pbase.ScaleHeight - pPreview.Height) \ 2
    
End Sub
