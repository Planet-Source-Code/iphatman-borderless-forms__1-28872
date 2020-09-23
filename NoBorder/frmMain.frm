VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Sample"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show New Form vbModal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   705
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1230
      Width           =   2070
   End
   Begin VB.Label CloseWindow 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3090
      TabIndex        =   1
      Top             =   15
      Width           =   135
   End
   Begin VB.Label TitleBarText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2940
   End
   Begin VB.Shape TitleBar 
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   3270
   End
   Begin VB.Shape MainWindow 
      Height          =   180
      Left            =   0
      Top             =   285
      Width           =   3270
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdShow_Click()
frmSample.Show vbModal, Me
End Sub

Private Sub Form_Load()
FormDesign Me, Me.Height, Me.Width, "Main Window"
End Sub

Private Sub TitleBarText_DblClick()
MinMaxWindow Me
End Sub

Private Sub TitleBarText_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
TitleBarClick Me, button, x, y
End Sub

Private Sub TitleBarText_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
WindowMove Me, x, y
End Sub

Private Sub TitleBarText_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
WindowStopMove
End Sub

Private Sub CloseWindow_Click()
End
End Sub
