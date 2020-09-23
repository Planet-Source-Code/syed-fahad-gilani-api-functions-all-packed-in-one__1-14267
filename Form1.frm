VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MagicBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fahad's Control Box"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Remote Control"
      ForeColor       =   &H00000000&
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame6 
         Caption         =   "Registry Control"
         Height          =   5655
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   3375
         Begin VB.CommandButton Command11 
            Caption         =   "Minimize All Windows"
            Height          =   375
            Left            =   480
            TabIndex        =   21
            Top             =   2760
            Width           =   2295
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Set"
            Height          =   255
            Left            =   2640
            TabIndex        =   20
            Top             =   1320
            Width           =   615
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Set"
            Height          =   255
            Left            =   2640
            TabIndex        =   19
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Text            =   "Recycle Bin"
            Top             =   1320
            Width           =   1095
         End
         Begin MSComCtl2.UpDown ud1 
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox menuspeed 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.Label speed 
            Caption         =   "1         -    Fastest   1000   -   Slowest"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Set Recycle Bin Name:"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Set Windows Menu Display Speed: (Requires Restart)"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Windows Programs"
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton Command8 
            Caption         =   "Open Calculator"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Open Control Panel"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Open Notepad"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mouse"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   1935
         Begin VB.CommandButton Command5 
            Caption         =   "Show Mouse"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Show Mouse Cursor"
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Hide Mouse"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "CAUTION! This will hide the mouse pointer!"
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "KeyBoard"
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
         Begin VB.CommandButton Command3 
            Caption         =   "Open Start Menu"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Caps Lock ON/OFF"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "CD-Rom"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton Command1 
            Caption         =   "Open CD Door"
            Default         =   -1  'True
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "MagicBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim result
Dim lRes As Long
Dim lDisp As Long
Dim Secure As SECURITY_ATTRIBUTES

Private Sub Command1_Click()
If Command1.Caption = "Open CD Door" Then
    result = mciSendString("set cd door open", 0, 0, hWnd)
    Command1.Caption = "Close CD Door"
Else
    result = mciSendString("set cd door closed", 0, 0, hWnd)
    Command1.Caption = "Open CD Door"
End If
End Sub

Private Sub Command10_Click()
Dim data As String
data = Text1.Text

result = RegCreateKeyEx(HKEY_CLASSES_ROOT, "CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", 0, "", 0, KEY_ALL_ACCESS, Secure, lRes, lDisp)
result = RegSetValueEx(lRes, "", 0, REG_SZ, ByVal data, Len(data))
result = RegFlushKey(lRes)
result = RegCloseKey(lRes)
End Sub

Private Sub Command11_Click()
         ' 77 is the character code for the letter 'M'
         Call keybd_event(VK_LWIN, 0, 0, 0)
         Call keybd_event(77, 0, 0, 0)
         Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)

End Sub

Private Sub Command6_Click()
result = Shell("notepad.exe", vbNormalFocus)

End Sub

Private Sub Command2_Click()
Call keybd_event(VK_CAPITAL, 0, 0, 0)
Call keybd_event(VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub Command3_Click()
If Command3.Caption = "Open Start Menu" Then
    Call keybd_event(VK_CONTROL, 0, 0, 0)
    Call keybd_event(VK_ESCAPE, 0, 0, 0)
    Call keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
    Command3.Caption = "Close Start Menu"
Else
    Call keybd_event(VK_ESCAPE, 0, 0, 0)
    Call keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0)
    Command3.Caption = "Open Start Menu"
End If
    
End Sub

Private Sub Command4_Click()
result = ShowCursor(False)
End Sub

Private Sub Command5_Click()
result = ShowCursor(True)
End Sub

Private Sub Command7_Click()
result = Shell("rundll32.exe shell32.dll,Control_RunDLL", 5)
End Sub

Private Sub Command8_Click()
result = Shell("calc", 5)
End Sub

Private Sub Command9_Click()
Dim data As String
data = ud1.Value
result = RegCreateKeyEx(HKEY_CURRENT_USER, "Control Panel\Desktop", 0, "", 0, KEY_ALL_ACCESS, Secure, lRes, lDisp)
result = RegSetValueEx(lRes, "MenuShowDelay", 0, REG_SZ, ByVal data, Len(data))
result = RegFlushKey(lRes)
result = RegCloseKey(lRes)
End Sub

Private Sub Form_Load()
result = mciSendString("close all", 0, 0, hWnd)
result = mciSendString("open cdaudio alias cd wait shareable", 0, 0, hWnd)
Command1.Caption = "Open CD Door"

With ud1
    .BuddyControl = menuspeed
End With
ud1.Min = 1
ud1.Max = 1000
ud1.Wrap = True
ud1.Value = 400
Text1.MaxLength = 12
End Sub

Private Sub ud1_Change()
menuspeed.Text = ud1.Value
End Sub
