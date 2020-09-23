VERSION 5.00
Begin VB.Form TestFrm 
   Caption         =   "Form1"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin ShadowOCXDemo.Shadow Shadow1 
      Left            =   240
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      MoveShadow      =   -1  'True
   End
End
Attribute VB_Name = "TestFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shadow1.showShadow
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shadow1.closeShadow
End Sub
