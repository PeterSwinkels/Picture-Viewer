VERSION 5.00
Begin VB.Form SettingsWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Viewer"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox BrowseDelayOption 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox BrowseDelayBox 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Specify an interval of 0 to 9 seconds."
         Top             =   0
         Width           =   375
      End
      Begin VB.Label SecondsLabel 
         Caption         =   "second(s)."
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   0
         Width           =   855
      End
      Begin VB.Label DisplayEachImageForLabel 
         Caption         =   "Display each image for"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.CommandButton InformationButton 
      Caption         =   "&Information"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Displays information about this program."
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton HelpButton 
      Caption         =   "&Help"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Displays the help."
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton QuitButton 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      ToolTipText     =   "Closes this program."
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton ViewButton 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Starts viewing the pictures  in the selected directory."
      Top             =   4080
      Width           =   1095
   End
   Begin VB.OptionButton BrowseOption 
      Caption         =   "Browse &automatically."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Enables automatic browsing."
      Top             =   3360
      Width           =   1935
   End
   Begin VB.OptionButton BrowseOption 
      Caption         =   "Do &not browse automatically."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Disables automatic browsing."
      Top             =   3000
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.FileListBox FileList 
      Height          =   2040
      Hidden          =   -1  'True
      Left            =   2520
      Pattern         =   "*.bmp;*.emf;*.gif;*.ico;*.jfif;*.jpeg;*.jpg;*.wmf"
      System          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Lists any picture files of a supported type in the selected directory."
      Top             =   360
      Width           =   2295
   End
   Begin VB.DirListBox DirectoryList 
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select a directory containing pictures here."
      Top             =   360
      Width           =   2295
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Select a drive here."
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label PicturesFoundLabel 
      Caption         =   "Pictures found:"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label ViewPicturesInLabel 
      Caption         =   "View pictures in:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "SettingsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the viewer settings window.
Option Explicit

'This procedure checks and adjusts the specified delay if necessary.
Private Sub BrowseDelayBox_LostFocus()
On Error GoTo ErrorTrap
   
   If Val(BrowseDelayBox.Text) < 1 Then BrowseDelayBox.Text = "1"
   
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure hides/unhides the browse delay option depending on whether automatic browsing is enabled.
Private Sub BrowseOption_Click(Index As Integer)
On Error GoTo ErrorTrap

   BrowseDelayOption.Visible = BrowseOption(1).Value

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure updates the file list when the selected directory is changed.
Private Sub DirectoryList_Change()
On Error GoTo ErrorTrap
   
   FileList.Path = DirectoryList.Path

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure updates the directory list when the selected drive is changed.
Private Sub DriveList_Change()
On Error GoTo ErrorTrap

   DirectoryList.Path = DriveList.Drive

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   
   Unload Me
   
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure stores the viewer settings.
Private Sub GetSettings()
On Error GoTo ErrorTrap
Dim Index As Long

   With ViewerSettings
      .Automatic = BrowseOption(1).Value
      .Delay = Val(BrowseDelayBox.Text)
      .Path = FileList.Path
      ReDim .Images(0 To FileList.ListCount - 1)
      
      For Index = LBound(.Images()) To UBound(.Images())
         .Images(Index) = FileList.List(Index)
      Next Index
   
      .Index = LBound(.Images())
   End With

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub


'This procedure displays the help.
Private Sub HelpButton_Click()
On Error GoTo ErrorTrap
Dim HelpText As String

   HelpText = "While viewing pictures: " & vbCr
   HelpText = HelpText & String$(60, "-") & vbCr
   HelpText = HelpText & "-Mouse:" & vbCr
   HelpText = HelpText & "Left mouse button: Previous picture." & vbCr
   HelpText = HelpText & "Middle mouse button: First picture." & vbCr
   HelpText = HelpText & "Right mouse button: Next picture." & vbCr
   HelpText = HelpText & vbCr
   HelpText = HelpText & "-Keyboard:" & vbCr
   HelpText = HelpText & "End: Last picture." & vbCr
   HelpText = HelpText & "Escape: Back to main window." & vbCr
   HelpText = HelpText & "Home: First picture." & vbCr
   HelpText = HelpText & "Left arrow: Previous picture." & vbCr
   HelpText = HelpText & "Right arrow: Next picture." & vbCr
   HelpText = HelpText & vbCr
   HelpText = HelpText & String$(30, "=") & vbCr
   HelpText = HelpText & "Supported picture file formats:" & vbCr
   HelpText = HelpText & String$(60, "-") & vbCr
   HelpText = HelpText & ".BMP, .EMF, .GIF, .ICO, .JFIF, .JPEG, .JPG, and .WMF"

   MsgBox HelpText, vbInformation, App.Title & " - Help"

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure displays information about this program.
Private Sub InformationButton_Click()
On Error GoTo ErrorTrap
   
   MsgBox ProgramInformation(), vbInformation, App.Title & " - Information"
   
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure closes this window.
Private Sub QuitButton_Click()
On Error GoTo ErrorTrap
   
   Unload Me

Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure starts the viewing of the pictures in the specified directory.
Private Sub ViewButton_Click()
On Error GoTo ErrorTrap

   If FileList.ListCount = 0 Then
      MsgBox "There are no pictures in the directory with a supported format.", vbExclamation, App.Title & " - View"
   ElseIf FileList.ListCount > 0 Then
      GetSettings
      ViewerWindow.Show
   End If

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

