VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPicturePreview 
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "frmPicturePreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox OriginalPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4680
      ScaleHeight     =   825
      ScaleWidth      =   465
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   50
      Left            =   0
      TabIndex        =   6
      Top             =   6000
      Width           =   8415
   End
   Begin VB.VScrollBar VScroll 
      Height          =   5295
      LargeChange     =   50
      Left            =   8400
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.TextBox ShowZoom 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btn_Exit 
         Caption         =   "Exit"
         Height          =   600
         Left            =   1260
         TabIndex        =   3
         Top             =   0
         Width           =   640
      End
      Begin VB.CommandButton btn_ZoomOut 
         Caption         =   "Zoom Out"
         Height          =   600
         Left            =   635
         TabIndex        =   1
         Top             =   0
         Width           =   640
      End
      Begin VB.CommandButton btn_ZoomIn 
         Caption         =   "Zoo In"
         Height          =   600
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   640
      End
   End
   Begin VB.PictureBox PictureSlide 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   561
      TabIndex        =   4
      Top             =   720
      Width           =   8415
      Begin VB.PictureBox PicturePreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   240
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   313
         TabIndex        =   7
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmPicturePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gZoomFactor As Single

Private Sub btn_Exit_Click()
Unload Me
End Sub

Private Sub btn_ZoomIn_Click()
If gZoomFactor * OriginalPicture.Width > 2000 Then Exit Sub
If gZoomFactor * OriginalPicture.Height > 2000 Then Exit Sub
gZoomFactor = gZoomFactor * 2
Me.MousePointer = 11
ShowPicture (gZoomFactor)

Form_Resize
Me.MousePointer = 0
End Sub

Private Sub btn_ZoomOut_Click()
If gZoomFactor * OriginalPicture.Width < 100 Then Exit Sub
If gZoomFactor * OriginalPicture.Height < 100 Then Exit Sub

gZoomFactor = gZoomFactor / 2
Me.MousePointer = 11
ShowPicture (gZoomFactor)

Form_Resize
Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
Form_Resize
End Sub

Private Sub Form_Load()
PicturePreview.Top = 0
PicturePreview.Left = 0
gZoomFactor = 1
MainForm.Enabled = False
End Sub


Private Sub Form_Resize()

If Me.Height / 15 < 200 Then
    Me.Height = 200 * 15
End If
If Me.Width / 15 < 150 Then
    Me.Width = 150 * 15
End If

PictureSlide.Height = Me.Height / 15 - Toolbar1.Height - 33 - HScroll.Height
PictureSlide.Width = Me.Width / 15 - 7 - VScroll.Width
HScroll.Top = PictureSlide.Height + PictureSlide.Top
VScroll.Left = PictureSlide.Width + PictureSlide.Left
VScroll.Height = PictureSlide.Height
HScroll.Width = PictureSlide.Width

If PicturePreview.Height - PictureSlide.Height > 0 Then
    VScroll.Max = PicturePreview.Height - PictureSlide.Height
    VScroll.Visible = True
   Else
    VScroll.Visible = False
End If
If PicturePreview.Width - PictureSlide.Width > 0 Then
    HScroll.Max = PicturePreview.Width - PictureSlide.Width
    HScroll.Visible = True
   Else
    HScroll.Visible = False
End If

gZoomFactor = 1
Do Until OriginalPicture.Width * gZoomFactor <= PictureSlide.Width And OriginalPicture.Height * gZoomFactor <= PictureSlide.Height
    gZoomFactor = gZoomFactor - 0.01
Loop
ShowPicture (gZoomFactor)
CenterPicture
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainForm.Enabled = True
MainForm.SetFocus
End Sub

Private Sub HScroll_Change()
    PicturePreview.Left = -HScroll.Value
    PicturePreview.SetFocus
End Sub

Private Sub HScroll_Scroll()
    HScroll_Change
End Sub


Private Sub VScroll_Change()
    PicturePreview.Top = -VScroll.Value
    PicturePreview.SetFocus
End Sub

Private Sub VScroll_Scroll()
    VScroll_Change
End Sub
Private Sub ShowPicture(ZoomFactor)
PicturePreview.Cls
ShowZoom.Text = "Zoom : " & gZoomFactor
PicturePreview.Height = OriginalPicture.Height * ZoomFactor
PicturePreview.Width = OriginalPicture.Width * ZoomFactor
PicturePreview.PaintPicture OriginalPicture, 0, 0, OriginalPicture.Width * ZoomFactor, OriginalPicture.Height * ZoomFactor
End Sub
Private Sub CenterPicture()
If PictureSlide.Width > PicturePreview.Width Then
    PicturePreview.Left = (PictureSlide.Width - PicturePreview.Width) / 2
   Else
    PicturePreview.Left = 0
End If
If PictureSlide.Height > PicturePreview.Height Then
    PicturePreview.Top = (PictureSlide.Height - PicturePreview.Height) / 2
   Else
    PicturePreview.Top = 0
End If

End Sub
