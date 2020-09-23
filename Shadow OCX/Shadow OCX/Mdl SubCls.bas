Attribute VB_Name = "MdlShadow"

Option Explicit
'*#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#*#
'=====================================================================================
'Module                 : MdlShadow
'Purpose                : Uses the Subclassing technique to Subclass the controls' parent form
'                       : so that the the shadow form can be moved when the parent form is
'                       : is dragged. It also creates the Dark Screen Shot on the
'                       : shadow form's back
'Author                 : Unknown
'Modified               : Praveen Menon
'Last Modified          : 12th April 2002
'=====================================================================================
'*#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#**#*#

'Used for creating the ScreenShot on the back of the shadow form
Private Declare Function CreateDC Lib "Gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function BitBlt Lib "Gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32.dll" (ByVal hdc As Long) As Long
'=====================================================================================

'Used for SubClassing the form to know when the form is dragged
Public Declare Function CallWindowProc Lib "user32" _
                        Alias "CallWindowProcA" _
                        (ByVal lpPrevWndFunc As Long, _
                        ByVal hwnd As Long, _
                        ByVal Msg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
                        Alias "SetWindowLongA" _
                        (ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public Const WM_MYMSG = &H232
Public defWndProc As Long
Private gl_frm As Form
'=====================================================================================

Public I_xOffsetPos As Integer
Public I_yOffsetPos As Integer

Private Sub CaptureScreen(frm As Form, Pic As PictureBox)

  Dim hDCscr As Long

    Pic.Cls                                         'Clear the Drawing board
    hDCscr = CreateDC("DISPLAY", "", "", 0)         'Get the Device Context of the Screen
    'Draw what u've got to the ShadowForm
    'The vbSrcErase Constant for the dwRop Argument makes
    'sure it looks like a shadow when it's drawn
    BitBlt Pic.hdc, 0, 0, frm.Width, frm.Height, _
           hDCscr, frm.Left / Screen.TwipsPerPixelX, _
           frm.Top / Screen.TwipsPerPixelX, vbSrcErase
            
    DeleteDC hDCscr                                 'Clean up the mess

End Sub

Public Sub Hook(hwnd As Long)

    If defWndProc = 0 Then

        defWndProc = SetWindowLong(hwnd, _
                     GWL_WNDPROC, _
                     AddressOf WindowProc)
    End If
                                 
End Sub

Public Sub Init(ShadowFrm As Form, frm As Form, iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer, showMe As Boolean)

  Dim IsLoad As Boolean

    On Error Resume Next

      With ShadowFrm
          .Visible = False
          .ScaleMode = 3
          .Left = iLeft + I_xOffsetPos    ' to cope for the extended part of the shadow form
          .Top = iTop + I_yOffsetPos      ' to cope for the extended part of the shadow form
          .Width = iWidth
          .Height = iHeight

          CaptureScreen ShadowFrm, ShadowFrm.picShadow        ' Get the Screen SHot on the Shadowform's Back

          .Visible = showMe               'you can't give this value true by default
          'bcos the shadow is shown even when the control
          'is at the design state
      End With 'SHADOWFRM
      frm.ZOrder                          'Make sure the Shadow form is not on top
      Set gl_frm = frm

End Sub

Public Sub UnHook(hwnd As Long)

    If defWndProc > 0 Then
    
        Call SetWindowLong(hwnd, GWL_WNDPROC, defWndProc)
        defWndProc = 0

    End If
    
End Sub

Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    Select Case uMsg
      Case WM_MYMSG
        
        If gl_frm.WindowState = vbMinimized Then frmShadow.Visible = False: Exit Function ':( Expand Structure or consider reversing Condition
        
        DoEvents
        frmShadow.Hide
        DoEvents
        
        Init frmShadow, gl_frm, gl_frm.Left, gl_frm.Top, gl_frm.Width, gl_frm.Height, False
            
      Case Else
        
        WindowProc = CallWindowProc(defWndProc, _
                     hwnd, _
                     uMsg, _
                     wParam, _
                     lParam)
    End Select
    
End Function
