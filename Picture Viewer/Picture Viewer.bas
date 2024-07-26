Attribute VB_Name = "ViewerModule"
'This module contains this program's core procedures.
Option Explicit

'This structure defines the viewer settings.
Public Type ViewerSettingsStr
   Automatic As Boolean    'Indicates whether automatic browsing is on/off.
   Delay As Long           'Defines the interval in seconds between images for automatic browsing.
   Images() As String      'Defines a list of the image files to be displayed.
   Index As Long           'Defines the index of the image being viewed.
   Path As String          'Defines the path of the images to be viewed.
End Type

Public ViewerSettings As ViewerSettingsStr   'Contains the settings.

'This proedure handles any errors to occur.
Public Function HandleError(Optional DoNotAsk As Boolean = False) As Long
Dim ErrorCode As Long
Dim Message As String
Dim PreviousMousePointer As Long
Static Choice As Long

   PreviousMousePointer = Screen.MousePointer
   Screen.MousePointer = vbDefault
   
   ErrorCode = Err.Number
   Message = Err.Description
   Err.Clear
   
   On Error Resume Next
   If Not DoNotAsk Then
      If Not Right$(Message, 1) = "." Then Message = Message & "."
      Message = Message & vbCr & "Error code: " & Str$(ErrorCode)
      Choice = MsgBox(Message, vbAbortRetryIgnore Or vbExclamation, App.Title & " - Error")
   End If
   
   If Choice = vbAbort Then End
   If Choice = vbRetry Or Choice = vbIgnore Then Screen.MousePointer = PreviousMousePointer
   HandleError = Choice
End Function

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   SettingsWindow.Show
   
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub


'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & vbCr & "By: " & .CompanyName & vbCr & "**2001***"
   End With
   
   ProgramInformation = Information
   Exit Function

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Function



