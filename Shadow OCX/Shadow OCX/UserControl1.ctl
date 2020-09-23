VERSION 5.00
Begin VB.UserControl Shadow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "UserControl1.ctx":000E
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   1560
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   3120
      Picture         =   "UserControl1.ctx":0320
      Top             =   1320
      Width           =   570
   End
End
Attribute VB_Name = "Shadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"

Option Explicit
'Default Property Values:
Const m_def_xOffsetPos = 90
Const m_def_yOffsetPos = 90
Const m_def_MoveShadow = "False"
'Property Variables:
Dim m_xOffsetPos As Integer
Dim m_yOffsetPos As Integer
Dim m_MoveShadow As Boolean
Dim l_hwnd As Long

Public Function closeShadow()

    If MoveShadow Then
        Call UnHook(l_hwnd)
    End If
    Unload frmShadow
    Set frmShadow = Nothing

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,Flase
Public Property Get MoveShadow() As Boolean

    MoveShadow = m_MoveShadow

End Property

Public Property Let MoveShadow(ByVal New_MoveShadow As Boolean)

    m_MoveShadow = New_MoveShadow
    PropertyChanged "MoveShadow"

End Property

Public Function showShadow()

    If MoveShadow Then
        Call Hook(l_hwnd)
    End If

    I_xOffsetPos = m_xOffsetPos
    I_yOffsetPos = m_yOffsetPos

    Init frmShadow, UserControl.Parent, UserControl.Parent.Left, UserControl.Parent.Top, UserControl.Parent.Width, UserControl.Parent.Height, Not Ambient.UserMode

End Function

Private Sub Timer1_Timer()

    If Not Ambient.UserMode Then
        Timer1.Enabled = False
        Exit Sub '>---> Bottom
      Else 'NOT NOT...
        Timer1.Enabled = True
    End If

    If UserControl.Parent.WindowState = vbMinimized Then
        frmShadow.Visible = False
      Else 'NOT USERCONTROL.PARENT.WINDOWSTATE...
        frmShadow.Visible = True
        UserControl.Parent.ZOrder
    End If

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_MoveShadow = m_def_MoveShadow
    m_xOffsetPos = m_def_xOffsetPos
    m_yOffsetPos = m_def_yOffsetPos

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    If Not Ambient.UserMode Then
        Timer1.Enabled = False
      Else 'NOT NOT...
        Timer1.Enabled = True
    End If
    
    m_MoveShadow = PropBag.ReadProperty("MoveShadow", m_def_MoveShadow)
    m_xOffsetPos = PropBag.ReadProperty("xOffsetPos", m_def_xOffsetPos)
    m_yOffsetPos = PropBag.ReadProperty("yOffsetPos", m_def_yOffsetPos)
    
    l_hwnd = UserControl.Parent.hwnd

End Sub

Private Sub UserControl_Resize()

    Height = Image1.Height
    Width = Image1.Width
    Image1.Move 0, 0

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MoveShadow", m_MoveShadow, m_def_MoveShadow)
    Call PropBag.WriteProperty("xOffsetPos", m_xOffsetPos, m_def_xOffsetPos)
    Call PropBag.WriteProperty("yOffsetPos", m_yOffsetPos, m_def_yOffsetPos)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get xOffsetPos() As Integer

    xOffsetPos = m_xOffsetPos

End Property

Public Property Let xOffsetPos(ByVal New_xOffsetPos As Integer)

    m_xOffsetPos = New_xOffsetPos
    PropertyChanged "xOffsetPos"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get yOffsetPos() As Integer

    yOffsetPos = m_yOffsetPos

End Property

Public Property Let yOffsetPos(ByVal New_yOffsetPos As Integer)

    m_yOffsetPos = New_yOffsetPos
    PropertyChanged "yOffsetPos"
End Property
