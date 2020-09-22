VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "Form1.frx":0FD4
   ScaleHeight     =   1920
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1650
      Left            =   0
      Picture         =   "Form1.frx":1023
      ScaleHeight     =   1650
      ScaleWidth      =   4110
      TabIndex        =   1
      Top             =   240
      Width           =   4110
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   600
         Picture         =   "Form1.frx":28A0
         ScaleHeight     =   240
         ScaleWidth      =   2700
         TabIndex        =   4
         Top             =   240
         Width           =   2700
         Begin VB.Image Image3 
            Height          =   255
            Left            =   120
            Top             =   0
            Width           =   495
         End
         Begin VB.Image Image4 
            Height          =   255
            Left            =   720
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image5 
            Height          =   255
            Left            =   1440
            Top             =   0
            Width           =   615
         End
         Begin VB.Image Image6 
            Height          =   255
            Left            =   2160
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   3495
      End
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":2CB6
      ScaleHeight     =   240
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FS"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AntiSociaL's Divx Player 1.0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   0
         Width           =   3180
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   3600
         ToolTipText     =   "Exit"
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
    Dim AviName
Dim FileName As String
Dim Paused As Integer
Dim AB As Integer
Dim CD As Integer

Private Sub OpenMedia()
On Error Resume Next
Dim mssg As String * 255
AviName = FileName
    comstr = "open " & FileName & " Type avivideo Alias video"
'---Set For Form or Picture Box---
'Last$ = Form2.Picture1.hwnd & " Style " & &H40000000
'ToDo$ = "open " & FileName & " Type avivideo Alias video parent " & Last$
'i = mciSendString(ToDo$, 0&, 0, 0)
'i = mciSendString("put video window at 0 0 230 115", 0&, 0, 0)
'---End Set---
    'ComStr = "play video" 'Popup window (de
    '     fault)
    'ComStr = "play video fullscreen" 'Do th
    '     is for full screen
    'ComStr = "play video from 0 to 10" 'Pla
    '     ys first 10 frames
    'ComStr = "play video from 11 to 20" 'pl
    '     ays frames in the middle
    'ComStr = "play video from 0 to wait" 'p
    '     lays from beginning
    'and lets nothing execute
    'until done
    
    'ComStr = "pause video" 'Pause your vide
    '     o during play
    'ComStr = "resume video" 'Resumes (conti
    '     nues) paused video
    'ComStr = "stop video" 'stops your video
    '     during play
    'ComStr = "close video" 'Closes the file
    'Important not to forget this!!
    
    'TIP: When AVI is finished, you must clo
    '     se it. You cannot
    'play it again until it is closed unless
    '     you specifiy a starting point
    'at the "play" command.
    RV = mciSendString(comstr, 0&, 0, 0&)
    comstr = "status video length"
    RV = mciSendString(comstr, mssg, 255, 0)
    HScroll1.Max = mssg
End Sub

Private Sub Command1_Click()
'    Dim mssg As String * 255
'    x% = mciSendString("status video mode", mssg, 255, 0)
'  Text1.Text
Dim mssg As String * 255

  i = mciSendString("set video time format ms", mssg, 255, 0)
  i = mciSendString("status video length", mssg, 255, 0)
  'MsgBox "There are " & Str(mssg) & " milliseconds"
Text1.Text = Str(mssg)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    comstr = "close video"
    RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub HScroll1_Change()
comstr = "pause video"
RV = mciSendString(comstr, 0&, 0, 0&)
comstr = "play video from " & HScroll1.Value & " to " & HScroll1.Max
RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub Image1_Click()
    comstr = "close video"
    RV = mciSendString(comstr, 0&, 0, 0&)
    End
End Sub

Private Sub Image2_Click()
    comstr = "play video fullscreen"
    RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub Image3_Click()
If Paused = 1 Then
    comstr = "resume video"
    RV = mciSendString(comstr, 0&, 0, 0&)
    Paused = 0
    Else
    comstr = "play video"
    RV = mciSendString(comstr, 0&, 0, 0&)
    End If
    'Timer1.Enabled = True
    'Form2.Show
End Sub

Private Sub Image4_Click()
    Paused = 1
    comstr = "pause video"
    RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub Image5_Click()
    comstr = "stop video"
    RV = mciSendString(comstr, 0&, 0, 0&)
    'Timer1.Enabled = False
End Sub

Private Sub Image6_Click()
 comstr = "close video"
RV = mciSendString(comstr, 0&, 0, 0&)
CommonDialog1.Filter = "All Supported Files|*.avi;*.mpeg;*.mpg;*.mpe;*.mpv;*.m1v;*.mp2;*.mpv2;*.mp2v;*.mpa;*.ivf;*.mov;*.qt;*.dat;*.mp3"
Me.CommonDialog1.ShowOpen
FileName = GetShortFileName(CommonDialog1.FileName)
Call OpenMedia
End Sub

Private Sub Label1_Click()
    comstr = "play video fullscreen"
    RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
result& = SendMessage(Me.hwnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
result& = SendMessage(Me.hwnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
comstr = "pause video"
RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
comstr = "play video from " & Slider1.Value & " to " & Slider1.Max
RV = mciSendString(comstr, 0&, 0, 0&)
End Sub

Private Sub Timer1_Timer()
    Dim mssg As String * 255
    x% = mciSendString("status video position", mssg, 255, 0)
    HScroll1.Value = mssg
End Sub
