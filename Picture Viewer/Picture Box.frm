VERSION 5.00
Begin VB.Form ViewerWindow 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer AutomaticBrowser 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   0
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image ResizedPictureBox 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "ViewerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the picture viewer window.
Option Explicit


'This procedure automatically browses to the next picture if automatic browsing is enabled.
Private Sub AutomaticBrowser_Timer()
On Error GoTo ErrorTrap
   
   With ViewerSettings
      If .Index < UBound(.Images()) Then
         .Index = .Index + 1
         DisplayPicture
      End If
   End With

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure displays a picture from the selected directory.
Private Sub DisplayPicture()
On Error GoTo ErrorTrap

   Screen.MousePointer = vbHourglass
   
   With PictureBox
      ResizedPictureBox.Visible = False
      ResizedPictureBox.Picture = LoadPicture()
      .Visible = True
      .Picture = LoadPicture(ViewerSettings.Images(ViewerSettings.Index))
      .Left = (Me.ScaleWidth / 2) - (.Width / 2)
      .Top = (Me.ScaleHeight / 2) - (.Height / 2)
   End With
   
   Screen.MousePointer = vbDefault
   
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure gives the command to display a picture from the selected directory.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   
   DisplayPicture
   
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure responds to the user's key strokes.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

   With ViewerSettings
      Select Case KeyCode
         Case vbKeyEnd
            .Index = UBound(.Images())
         Case vbKeyEscape
           Unload Me
           Exit Sub
         Case vbKeyHome
           .Index = LBound(.Images())
         Case vbKeyLeft
            If .Index > 0 Then .Index = .Index - 1
         Case vbKeyRight
            If .Index < UBound(.Images()) Then .Index = .Index + 1
      End Select
   End With
   
   DisplayPicture
   
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub


'This procedure starts the viewing of the pictures in the selected directory.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   With ViewerSettings
      ChDrive Left$(.Path, InStr(.Path, ":"))
      ChDir .Path
      
      .Index = LBound(.Images())
      
      If .Automatic Then
         AutomaticBrowser.Enabled = True
         AutomaticBrowser.Interval = .Delay * 1000
      ElseIf Not .Automatic Then
         AutomaticBrowser.Enabled = False
         AutomaticBrowser.Interval = 0
      End If
   End With

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure gives the command to handle the user's mouse clicks.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap

   HandleClick Button

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure stops the viewing of the pictures.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   
   AutomaticBrowser.Enabled = False
   AutomaticBrowser.Interval = 0
   Unload Me
   
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure handles the user's mouse clicks.
Private Sub HandleClick(ByVal Button As Integer)
On Error GoTo ErrorTrap

   With ViewerSettings
      If Not .Automatic Then
         Select Case Button
            Case vbLeftButton
               If .Index > 0 Then .Index = .Index - 1
            Case vbRightButton
               If .Index < UBound(.Images()) Then .Index = .Index + 1
            Case vbMiddleButton
               .Index = LBound(.Images())
         End Select
      
         DisplayPicture
      End If
   End With
   
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub


'This procedure resizes the picture if necessary.
Private Sub PictureBox_Change()
On Error GoTo ErrorTrap
  
   With ResizedPictureBox
      If PictureBox.Width * Screen.TwipsPerPixelX > Screen.Width Or PictureBox.Height * Screen.TwipsPerPixelY > Screen.Height Then
         .Picture = PictureBox.Picture
         
         On Error Resume Next
         PictureBox.Visible = False
         .Width = Screen.Width / Screen.TwipsPerPixelX
         .Height = Screen.Height / Screen.TwipsPerPixelY
         .Left = (Me.ScaleWidth / 2) - (.Width / 2)
         .Top = (Me.ScaleHeight / 2) - (.Height / 2)
         .Visible = True
      End If
   End With

   Exit Sub
 
ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure gives the command to handle the user's mouse clicks.
Private Sub PictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap

   HandleClick Button
   
   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

'This procedure gives the command to handle the user's mouse clicks.
Private Sub ResizedPictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap

   HandleClick Button

   Exit Sub

ErrorTrap:
   If HandleError() = vbRetry Then Resume
   If HandleError(DoNotAsk:=True) = vbIgnore Then Resume Next
End Sub

