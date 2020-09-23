VERSION 5.00
Begin VB.Form frmMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1215
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4365
      ScaleHeight     =   420
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   10000
      Width           =   465
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2700
      Top             =   495
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   1260
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   45
      Width           =   2760
   End
   Begin VB.Label lblYesNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Index           =   0
      Left            =   315
      TabIndex        =   2
      Top             =   720
      Width           =   870
   End
   Begin VB.Label lblYesNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Index           =   1
      Left            =   1485
      TabIndex        =   1
      Top             =   720
      Width           =   870
   End
   Begin VB.Label lblOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   900
      TabIndex        =   0
      Top             =   720
      Width           =   870
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
 
Enum enButtons
    None = 0
    Ok = 1
    YesNo = 2
End Enum


Dim messageBoxRet       As Long
Dim m_buttons           As enButtons
Dim m_TransparentColor  As Long
  
Function Message(callingForm As Form, _
                  strMessage As String, _
                  Optional strCaption As String, _
                  Optional mForecolor As Long, _
                  Optional Buttons As enButtons, _
                  Optional AutoHideSeconds As Long, _
                  Optional Picture As StdPicture, _
                  Optional TransparentColor As Long)
 
   '  assign the captions for title and message
   lblCaption = strCaption
   lblMsg = strMessage
   lblMsg.Forecolor = mForecolor
   m_buttons = Buttons
   m_TransparentColor = TransparentColor
   Set Picture1.Picture = Picture
   '
   Call PositionFormsControls
   '
   '  center this form to the calling form
   Me.Move PositionMeLeft(callingForm), PositionMeTop(callingForm)
   '
   Call SetSelfCloseOptions(Buttons, AutoHideSeconds)
   '
   '  set the appropriate buttons
   Call ShowCorrectButtons(Buttons)
   '
   'based upon the buttons used we can now set
   'forms modality and show this form
   Me.Show DeterminedModality(Buttons), callingForm
   '
   '  return whether yes or no was pressed
   Message = messageBoxRet
   
End Function
 
Private Sub PositionFormsControls()
  '
  'we need to position the controls so that
  'this form matches the size of lblMessage
  'and that the buttons are in the right place
  With lblMsg
     .Move 100, 300
      Width = .Width + 100
      Height = .Height + 800
  End With
  
  '  position caption label
  lblCaption.Move 0, 0, Width, 230
  '  position OK button
  With lblOk
     .Move (Width * 0.5) - (.Width * 0.5), (Height - 300)
  End With
  '  position Yes button
  With lblYesNo(0)
     .Move (Width * 0.5) - (.Width + 50), (Height - 300)
  End With
  '  position No button
  With lblYesNo(1)
     .Move (Width * 0.5) + 50, (Height - 300)
  End With
  
End Sub

Private Function PositionMeLeft(callingForm As Form) As Long

  Dim halfOfMe As Long, halfOfYou As Long, leftPoint As Long
 
  With callingForm
     '  store coods for positioning this left
     halfOfMe = (Width * 0.5)
     halfOfYou = .Left + (.Width * 0.5)
     PositionMeLeft = (halfOfYou - halfOfMe)
  End With
  
End Function

Private Function PositionMeTop(callingForm As Form) As Long

  Dim halfOfMe As Long, halfOfYou As Long, toppoint As Long
  
  With callingForm
     '  store coods for positioning this top
     halfOfMe = (Height * 0.5)
     halfOfYou = .Top + (.Height * 0.5)
     PositionMeTop = (halfOfYou - halfOfMe)
  End With
  
End Function

Private Sub SetSelfCloseOptions(Buttons As enButtons, AutoHideSeconds As Long)
   '
   '  are we autoclosing this ?
   '  autoclose is NOT a valid option
   '  if the messagebox is a yesNo messagebox
   If Buttons <> YesNo Then
      If Buttons = None Then
         '  however, if there are not buttons,
         '  this HAS to be self closing
         If AutoHideSeconds < 1 Then
           Timer1.Interval = 3000
           Timer1.Enabled = True
         End If
      Else
        '  user chose OK button
        If AutoHideSeconds > 0 Then
           Timer1.Interval = (AutoHideSeconds * 1000)
           Timer1.Enabled = True
        End If
      End If
   End If
   
End Sub

Private Function DeterminedModality(Buttons As enButtons) As Long
  '
  'this function will determine when and when not
  'to make this message box modal or not, obviously
  'with no buttons it cant be modal (and better be
  'self closing)
  DeterminedModality = 1
  
  If Buttons = None Then
      DeterminedModality = 0
  End If
  
End Function


Private Sub ShowCorrectButtons(Buttons As enButtons)
  '
  '  show and position specified buttons
  '
  'set all hidden as default
  lblOk.Visible = False
  lblYesNo(0).Visible = False
  lblYesNo(1).Visible = False
  '
  If Buttons = None Then
     
  ElseIf Buttons = Ok Then
      lblOk.Visible = True
  ElseIf Buttons = YesNo Then
      lblYesNo(0).Visible = True
      lblYesNo(1).Visible = True
  End If
  
End Sub

Private Sub MoveItem(item_hwnd As Long)
    '
    'for moving this form around
    '
    On Error Resume Next
    Const WM_NCLBUTTONDOWN As Long = &HA1
    Const HTCAPTION As Long = 2
    ReleaseCapture
    SendMessage item_hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    
End Sub



Private Sub Form_Paint()

  Dim lcnt As Long
  '
  'paint faint horizontal lines for balance and attractiveness
  '
  For lcnt = 250 To Height Step 50
     Me.Line (20, lcnt)-(Width - 20, lcnt), RGB(230, 250, 250)
  Next lcnt
  '  dark line surrounding message part of box
  Me.Line (0, 230)-(Width - 20, Height - 20), RGB(160, 160, 175), B
  '
  '  If there is a picture to paint then paint it
  With Picture1
     If Not (.Picture Is Nothing) Then
         Dim pixwid  As Long, pixhei As Long
         Dim toppoint As Long, titlebarHeight As Long
         
         '  get api friendly measurements of the source pic
         pixwid = (.Width / Screen.TwipsPerPixelX)
         pixhei = (.Height / Screen.TwipsPerPixelY)
         titlebarHeight = (250 / Screen.TwipsPerPixelX)
         toppoint = _
             ((Height * 0.5) / Screen.TwipsPerPixelX) - _
             ((.Height * 0.5) / Screen.TwipsPerPixelY)
         ' we dont want the top position of the picture to
         '  be any higher up than the bottom of our titlebar
         If toppoint <= titlebarHeight Then
            toppoint = titlebarHeight
         End If
         '  paint the pic
         TransparentBlt hdc, 5, toppoint, pixwid, pixhei, _
            .hdc, 0, 0, pixwid, pixhei, m_TransparentColor
     End If
  End With
End Sub
 
 
Private Sub Form_Terminate()
  
  Timer1 = False
  
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  '
  'move this
  '
  Call MoveItem(hwnd)
  
End Sub

Private Sub lblMsg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  '
  'move this
  '
  Call MoveItem(hwnd)
  
End Sub
 
Private Sub lblOk_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

 Unload Me

End Sub

Private Sub lblYesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
 ' return the result
 If Index = 0 Then 'yes was pressed
    messageBoxRet = vbYes
 ElseIf Index = 1 Then 'no was pressed
    messageBoxRet = vbNo
 End If
 
 Unload Me

End Sub

Private Sub Timer1_Timer()
  
  Timer1.Interval = 0
  Unload Me

End Sub

 
Private Sub Form_KeyPress(KeyAscii As Integer)
 '
 'If the user presses one of the shortcut
 'keys for one of the buttons
 '
 '  Enter key closes all boxes
 '  If the YesNo buttons are visible then
 '  Yes is the default
 If m_buttons = YesNo Then
   If KeyAscii = 13 Or KeyAscii = 121 Then
      Call lblYesNo_MouseUp(0, 1, 0, 0, 0)
   ElseIf KeyAscii = 110 Then
      Call lblYesNo_MouseUp(1, 1, 0, 0, 0)
   End If
 Else
    Unload Me
 End If
 
End Sub

