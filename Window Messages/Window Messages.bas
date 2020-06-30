Attribute VB_Name = "WindowMessagesModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this program.
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const GWL_WNDPROC As Long = -4

Private Declare Function CallWindowProcA Lib "User32.dll" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function IsWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "User32.dll" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'The constants and variables used by this program.
Public Const NO_WINDOW As Long = 0         'Defines a null window handle.
Private Const MAX_STRING As Long = 65535   'Defines the maximum length allowed for a string buffer.

Private ActualWindowMessageHandler As Long   'Contains a window's actual message handler.


'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   If Not ErrorCode = ERROR_SUCCESS Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCrLf
      MsgBox Message, vbExclamation
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim ErrorCode As String
Dim Message As String

   ErrorCode = Err.Number
   Message = Err.Description
   On Error Resume Next
   MsgBox "Error: " & CStr(ErrorCode) & vbCr & Message, vbExclamation
End Sub


'This procedure handles a window's messages.
Private Function HandleWindowMessages(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo NoMessageBox

   With InterfaceWindow.MessageBox
      .Text = .Text & PadText("0x" & Hex$(hwnd), 10, " ", PadRight:=True) & "   "
      .Text = .Text & PadText("0x" & Hex$(uMsg), 10, " ", PadRight:=True) & "   "
      .Text = .Text & PadText("0x" & Hex$(wParam), 10, " ", PadRight:=True) & "   "
      .Text = .Text & PadText("0x" & Hex$(lParam), 10, " ", PadRight:=True) & vbCrLf
   End With
EndMessageBox:
   
   On Error GoTo ErrorTrap
   
EndRoutine:
   HandleWindowMessages = CheckForError(CallWindowProcA(ActualWindowMessageHandler, hwnd, uMsg, wParam, lParam))
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
   
NoMessageBox:
   Resume EndMessageBox
End Function

'This procedure returns the text with the specified padding of the specified length.
Private Function PadText(Text As String, Length As Long, Optional Padding As String = " ", Optional PadRight As Boolean = False) As String
On Error GoTo ErrorTrap
Dim PaddedText As String

   If PadRight Then
      PaddedText = String$(Length - Len(Text), Padding) & Text
   Else
      PaddedText = Text & String$(Length - Len(Text), Padding)
   End If
   
EndRoutine:
   PadText = PaddedText
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the current window hook.
Public Function WindowHook(Optional NewWindowH As Long = NO_WINDOW, Optional Unhook As Boolean = False) As Long
On Error GoTo ErrorTrap
Static CurrentWindowH As Long

   If Not (NewWindowH = NO_WINDOW Or NewWindowH = CurrentWindowH) Then
      If CBool(CheckForError(IsWindow(NewWindowH))) Then
         InterfaceWindow.MessageBox.Text = PadText("Handle:", 13) & PadText("Message:", 13) & PadText("wParam:", 13) & PadText("lParam:", 13) & vbCrLf
         CurrentWindowH = NewWindowH
         ActualWindowMessageHandler = CheckForError(SetWindowLongA(CurrentWindowH, GWL_WNDPROC, AddressOf HandleWindowMessages))
      Else
         MsgBox "The specified handle does not refer to a window (anymore.)", vbExclamation
      End If
   ElseIf Unhook And Not CurrentWindowH = NO_WINDOW Then
      If CBool(CheckForError(IsWindow(CurrentWindowH))) Then
         CheckForError SetWindowLongA(CurrentWindowH, GWL_WNDPROC, ActualWindowMessageHandler)
         CurrentWindowH = NO_WINDOW
      End If
   End If
   
EndRoutine:
   WindowHook = CurrentWindowH
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

