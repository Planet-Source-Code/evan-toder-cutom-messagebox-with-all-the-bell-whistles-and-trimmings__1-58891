VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Test"
      Height          =   285
      Left            =   225
      TabIndex        =   2
      Top             =   1125
      Width           =   1230
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   2700
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "You can draw any picture on the custom message box. any size. AND you can specify a color that is transparent or not painted"
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   270
      TabIndex        =   1
      Top             =   135
      Width           =   4065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Debug.Print frmMessage.Message _
            (Me, "This is a sample message in the " & _
            "custom message box." & vbCrLf & _
            "Like the windows message box, this one" & _
            " automatically resizes itself and all of its " & _
            "constituent controls to adjust to the message." & vbCrLf & _
            "Unlike the windows message box, you have to explicitly " & _
            "use vbCrLf  to break the lines apart..which means you have" & vbCrLf & _
            "more direct and precise control." & _
            "Also, you can specify any kind and size of picture, AND" & vbCrLf & _
            "you can have the messagebox self close, just as long as" & vbCrLf & _
            "your not using the YesNo buttons." & vbCrLf & _
            "You can also specify the forecolor for the message as well.", _
            "THIS MESSAGE BOX SELF CENTERS TO THE CALLING FORM", _
             vbBlue, YesNo, , Picture1, RGB(255, 0, 255))
End Sub
