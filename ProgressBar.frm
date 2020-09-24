VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Window Information"
   ClientHeight    =   735
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   1845
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   240
      Width           =   15
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentPercent As Integer

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Form_Load()
MsgBox ("Look at you clock at System Tray")
    Dim hWnd As Long, rctemp As RECT
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    hWnd = FindWindowEx(hWnd, 0, "TrayNotifyWnd", vbNullString)
    'hWnd = FindWindowEx(hWnd, 0, "TrayClockWClass", vbNullString) 'uncomment
    'this string and look progressbar only at clock
    GetWindowRect hWnd, rctemp
    With Me
        .Top = 0
        .Left = 0
        .Height = Me.Height * (rctemp.Bottom - rctemp.Top) / Me.ScaleHeight
        .Width = Me.Width * (rctemp.Right - rctemp.Left) / Me.ScaleWidth
    End With
    Timer.Enabled = True
    SetParent Me.hWnd, hWnd
    Picture1.Height = Me.ScaleHeight
    Picture1.Width = Me.ScaleWidth
End Sub
Private Sub Form_DblClick()
MsgBox ("Please, Vote me")
Unload Me
End Sub
Private Sub Picture1_Click()
MsgBox ("Please, Vote me")
Unload Me
End Sub
Private Sub Form_Resize()
    Picture1.Top = (Me.ScaleHeight - Picture1.Height) / 2
End Sub
Public Function UpdateProgress(pb As Control, ByVal Percent)
Dim Num$
If Not pb.AutoRedraw Then
pb.AutoRedraw = -1
End If
pb.Cls
pb.ScaleWidth = 100
pb.DrawMode = 10
Num$ = Format$(Percent, "###") + "%"
pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
pb.Print Num$
pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
pb.Refresh
End Function
Private Sub Timer_Timer()
CurrentPercent = CurrentPercent + 1
If CurrentPercent < 101 Then
UpdateProgress Picture1, CurrentPercent
Else
Timer.Enabled = False
CurrentPercent = 0
End If
End Sub
