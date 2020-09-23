VERSION 5.00
Begin VB.Form frmShadow 
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   1620
   ClientTop       =   1215
   ClientWidth     =   6660
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picShadow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   0
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "frmShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    picShadow.Move 0, 0

End Sub

Private Sub Form_Resize()

    picShadow.Width = Me.ScaleWidth
    picShadow.Height = Me.ScaleHeight
End Sub

