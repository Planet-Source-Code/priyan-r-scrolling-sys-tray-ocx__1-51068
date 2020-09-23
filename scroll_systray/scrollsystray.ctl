VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl scrollsystray 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Systray 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox tmpIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   1440
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   840
      Width           =   240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   240
   End
   Begin MSComctlLib.ImageList IconRipper 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
   End
End
Attribute VB_Name = "scrollsystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim scroll_text$
Event click(Button As Integer)
Event doubleclick()
Private xx As Boolean
Private yy As Boolean
Private zz As Boolean
Private ss As Boolean
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private nID(2) As NOTIFYICONDATA  'change here for more width
'Default Property Values:



Private Sub Timer3_Timer()
 Dim Icon As Variant
   Dim B As Integer
    Static A As Long

    Timer3.Enabled = False

    A = A - 1

    If A < -Systray.TextWidth(scrolltext) Or nID(0).hIcon = Icon Then A = Systray.ScaleWidth
    Systray.BackColor = Systray.BackColor
    Systray.CurrentX = A

        Systray.CurrentY = 3

    Systray.Print scroll_text
    Systray.Refresh
    For B = 0 To UBound(nID)
        BitBlt tmpIcon(B).hDC, 0, 0, 16, 16, Systray.hDC, (UBound(nID) - B) * 18, 0, vbSrcCopy
        tmpIcon(B).Refresh
        IconRipper.ListImages.Add , "temp" & B, tmpIcon(B).Image
        With nID(B)
            .uFlags = NIF_ICON
            .hIcon = IconRipper.ListImages("temp" & B).ExtractIcon
        End With
        Shell_NotifyIcon NIM_MODIFY, nID(B)
    Next B
    IconRipper.ListImages.Clear
    Timer3.Enabled = True
End Sub

Private Sub UserControl_Initialize()
Systray.Left = 0
Systray.Top = 0
 Dim A As Integer
    For A = 1 To UBound(nID)
        Load tmpIcon(A)
    Next A
    Systray.Width = Systray.Height * A + UBound(nID) * 30
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Systray.Width
UserControl.Height = Systray.Height
End Sub


Public Sub scroll()
Timer3.Enabled = True
Dim A As Integer
    Dim Icon As Variant
    For A = 0 To UBound(nID)
        
        With nID(A)
            .cbSize = Len(nID(A))
            
            .hwnd = tmpIcon(A).hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Icon
            .szTip = "Internet time 3" & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, nID(A)
    Next A

End Sub

Private Sub UserControl_Terminate()
 Dim A As Integer
    For A = 0 To UBound(nID)
        Shell_NotifyIcon NIM_DELETE, nID(A)
    Next A
    For A = tmpIcon.Count - 1 To 1 Step -1
        Unload tmpIcon(A)
    Next A
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Systray,Systray,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Systray.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Systray.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Systray,Systray,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Systray.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Systray.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Systray.BackColor = PropBag.ReadProperty("BackColor", &HFF&)
    Systray.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    scroll_text = PropBag.ReadProperty("scrolltext", "priyan_rajeevan@rediffmail.com")
'    Set Systray.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Systray.BackColor, &HFF&)
    Call PropBag.WriteProperty("ForeColor", Systray.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("scrolltext", scroll_text, "priyan_rajeevan@rediffmail.com")
'    Call PropBag.WriteProperty("Font", Systray.Font, Ambient.Font)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get scrolltext() As Variant
    scrolltext = scroll_text
End Property

Public Property Let scrolltext(ByVal New_scrolltext As Variant)
    scroll_text = New_scrolltext
    PropertyChanged "scrolltext"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Systray,Systray,-1,Font
'Public Property Get Font() As Font
'    Set Font = Systray.Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set Systray.Font = New_Font
'    PropertyChanged "Font"
'End Property
'
Private Sub tmpIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    Select Case CLng(X)
        Case WM_LBUTTONDBLCLK 'left button doubleclick
            RaiseEvent doubleclick
        Case WM_RBUTTONUP 'right button click
            RaiseEvent click(vbRightButton)
        Case WM_LBUTTONUP 'left button click
            RaiseEvent click(vbLeftButton)
    End Select
End Sub
Public Sub about()
Attribute about.VB_UserMemId = -552
MsgBox "Programmed By Priyan" & vbCrLf & vbCrLf & "Visit me at www.priyan.tk", vbInformation, "Mail Active-X Control"
End Sub

