VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan's Scroll Systray"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin itime.scrollsystray scrollsystray1 
      Left            =   720
      Top             =   120
      _ExtentX        =   1376
      _ExtentY        =   423
      BackColor       =   12582912
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Visit me at wwww.priyan.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vote For Me!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Menu menux 
      Caption         =   "Menux"
      Visible         =   0   'False
      Begin VB.Menu connectx 
         Caption         =   "CONNECT"
      End
      Begin VB.Menu disconx 
         Caption         =   "DICONNECT"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu statistic 
         Caption         =   "Statistic"
      End
      Begin VB.Menu interset 
         Caption         =   "Internet settings"
      End
      Begin VB.Menu modemset 
         Caption         =   "Modem settings"
      End
      Begin VB.Menu setup 
         Caption         =   "Setup"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
scrollsystray1.scroll
End Sub

Private Sub Command2_Click()
Me.scrollsystray1.scrolltext = Text1.Text
scrollsystray1.scroll
End Sub

Private Sub Form_Load()
Me.Text1.Text = scrollsystray1.scrolltext
End Sub

Private Sub scrollsystray1_click(Button As Integer)
If Button = 1 Then
    MsgBox "Left Button", vbInformation
ElseIf Button = 2 Then
    MsgBox "Right Button", vbInformation
End If
End Sub
