VERSION 5.00
Begin VB.Form InterfaceWindow 
   ClientHeight    =   2310
   ClientLeft      =   105
   ClientTop       =   810
   ClientWidth     =   3630
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   242
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MessageBox 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3612
   End
   Begin VB.Menu MonitorMainMenu 
      Caption         =   "&Monitor"
      Begin VB.Menu EndMonitorMenu 
         Caption         =   "&End Monitor"
         Shortcut        =   {F1}
      End
      Begin VB.Menu StartMonitorMenu 
         Caption         =   "&Start Monitor"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'This procdure stops any active window monitors.
Private Sub EndMonitorMenu_Click()
On Error GoTo ErrorTrap
   WindowHook , Unhook:=True
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure initializes this window when this program is started.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   With App
      Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With
   
   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure adjusts this window to its new size.
Private Sub Form_Resize()
On Error Resume Next
   MessageBox.Width = Me.ScaleWidth
   MessageBox.Height = Me.ScaleHeight
End Sub


'This procedure stops any active window monitors when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   WindowHook , Unhook:=True
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to start monitoring the window specified by the user.
Private Sub StartMonitorMenu_Click()
On Error GoTo ErrorTrap
Dim DefaultH As Long
Dim WindowH As String
   
   DefaultH = WindowHook()
   If DefaultH = NO_WINDOW Then DefaultH = Me.hwnd
   WindowH = LCase$(Trim$(InputBox$("Window handle (prefix with 0x for hexadecimals):", , CStr(DefaultH))))
   If Left$(WindowH, 2) = "0x" Then WindowH = "&H" & Mid$(WindowH, 3) & "&"
   
   If Not CLng(Val(WindowH)) = NO_WINDOW Then
      WindowHook , Unhook:=True
      WindowHook NewWindowH:=CLng(Val(WindowH))
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

