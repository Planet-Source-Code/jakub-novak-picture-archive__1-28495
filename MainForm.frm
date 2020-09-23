VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   Caption         =   "Picture Archive"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9930
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   557
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   Begin MSComDlg.CommonDialog CD 
      Left            =   9240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btn_Stop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ImageList cdsIcons 
      Left            =   720
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4768
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":533A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Search_in 
      Caption         =   "Search in"
      Height          =   1095
      Left            =   6240
      TabIndex        =   17
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox SearchKeyWords 
         Caption         =   "Key words"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   760
         Width           =   1215
      End
      Begin VB.CheckBox SearchDescription 
         Caption         =   "Description"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   520
         Width           =   1095
      End
      Begin VB.CheckBox SearchFileName 
         Caption         =   "File name"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton btn_Search 
      Caption         =   "Search"
      Height          =   300
      Left            =   5400
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox SearchText 
      Height          =   285
      Left            =   3840
      TabIndex        =   15
      Top             =   360
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   150
      Left            =   0
      TabIndex        =   14
      Top             =   8160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox SelectedFileID 
      Height          =   285
      Left            =   8400
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicThumbnail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3840
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PicLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7800
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   8100
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FileDetails 
      Caption         =   "Edit Details"
      Enabled         =   0   'False
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   3735
      Begin RichTextLib.RichTextBox FileKeyWords 
         Height          =   1575
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2778
         _Version        =   393217
         TextRTF         =   $"MainForm.frx":5C14
      End
      Begin VB.TextBox FileDescription 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Key Words"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox ThumbFrame 
      Height          =   6855
      Left            =   3840
      ScaleHeight     =   453
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
      Begin VB.PictureBox ThumbSlide 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   0
         ScaleHeight     =   209
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   6
         Top             =   0
         Width           =   3135
         Begin VB.OptionButton ThumbNail 
            Alignment       =   1  'Right Justify
            Height          =   2175
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.VScrollBar ThumbScroll 
         Height          =   6735
         Left            =   5160
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
   End
   Begin MSComctlLib.TreeView CDs 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Right click on the left pane"
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9763
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "cdsIcons"
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "Search text"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuCDs 
      Caption         =   "CDs"
      Visible         =   0   'False
      Begin VB.Menu mnuAddCDCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu empty1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenameCD 
         Caption         =   "Rename CD"
      End
      Begin VB.Menu mnuAddCD 
         Caption         =   "Add CD"
         Begin VB.Menu smnuAddCD 
            Caption         =   "x"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRemoveCD 
         Caption         =   "Remove CD"
         Begin VB.Menu smnuRemoveCD 
            Caption         =   "x"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRescan 
         Caption         =   "Rescan Pictures"
         Begin VB.Menu smnuRescan 
            Caption         =   "x"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu smnuOpenDB 
         Caption         =   "Select Database"
      End
      Begin VB.Menu smnuCompactDB 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu empty2 
         Caption         =   "-"
      End
      Begin VB.Menu smnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ThumbRows As Long
Public ThumbCols As Long
Public ThumbRow As Long
Public ThumbCol As Long
Public ThumbHeight As Long
Public ThumbWidth As Long
Public PicLoadWidth As Long
Public PicLoadHeight As Long
Public TotalCDs As Long
Public DBPath As String
Public gSelectedThumbNail As Long
Public gDirShown As String
Public gStillReadingFiles As Boolean
Public gStopReadingFiles As Boolean
Public gShowContent As Boolean


Public xDb As Database
'Public rsCD As Recordset
'Public Dim rsFiles As Recordset
Dim TheBytes() As Byte

Private Type PointAPI
    x  As Long
    Y  As Long
End Type
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As PointAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ConvertBMPtoJPG Lib "bmpTojpg.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal JPGCompressQuality As Integer, ByVal blnKeepBMP As Boolean) As Integer
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal Default As String, ByVal ReturnedString As String, ByVal StringSize As Long, ByVal FileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal FileName As String) As Long

Private Sub btn_Search_Click()
Dim query
Dim rsFiles As Recordset
Dim Criteria


On Error GoTo err1

ClearThumbnails

Criteria = ""
If SearchFileName Then Criteria = "WHERE Name LIKE '*" & SearchText.Text & "*'"
If SearchDescription Then
    If Criteria = "" Then
        Criteria = "WHERE Description LIKE '*" & SearchText.Text & "*'"
       Else
        Criteria = " OR Description LIKE '*" & SearchText.Text & "*'"
    End If
End If
If SearchKeyWords Then
    If Criteria = "" Then
        Criteria = "WHERE KeyWords LIKE '*" & SearchText.Text & "*'"
       Else
        Criteria = " OR KeyWords LIKE '*" & SearchText.Text & "*'"
    End If
End If

If Criteria = "" Then Exit Sub
query = "SELECT * FROM Files " & Criteria
Set rsFiles = xDb.OpenRecordset(query)
If rsFiles.RecordCount = 0 Then
    rsFiles.Close
    Exit Sub
End If
rsFiles.MoveLast
rsFiles.MoveFirst

SetHourglass (True)
PB.Visible = True
PB.Max = 100
Do Until rsFiles.EOF
    ShowThumbnails rsFiles("ID")
    PB.Value = rsFiles.PercentPosition
    rsFiles.MoveNext
    DoEvents
Loop
PB.Visible = False
SetHourglass (False)

rsFiles.Close
ReSortThumbnails
StatusBar.Panels(1).Text = "Found " & ThumbNail.Count - 1 & " Pictures"

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub CDs_Collapse(ByVal Node As MSComctlLib.Node)
If gShowContent Then ShowContent (CDs.SelectedItem.Index)
End Sub

Private Sub CDs_Expand(ByVal Node As MSComctlLib.Node)
If gShowContent Then ShowContent (CDs.SelectedItem.Index)
End Sub


Private Sub CDs_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i, CDID
Dim mNode As Node
Dim rsCD As Recordset
Dim FileID
Dim NodePath
Dim ChildNode As Node
Dim ChildNodeID
Dim ParentNode As Node


On Error GoTo err1
Set mNode = CDs.HitTest(x, Y)

'Pop up menu
If Button = 2 Then
StatusBar.Panels(2).Text = ""
GetDrivesInfo
If UBound(CDDrives) = 0 Then Exit Sub


If Not mNode Is Nothing Then
    CDs.Nodes(mNode.Index).Selected = True
'    ShowContent (mNode.Index)
End If
'unload submenus
For i = 2 To smnuRemoveCD.Count: Unload smnuRemoveCD(i): Next i
For i = 2 To smnuAddCD.Count: Unload smnuAddCD(i): Next i
For i = 2 To smnuRescan.Count: Unload smnuRescan(i): Next i

'Rename CD Menu
If mNode Is Nothing Then
    mnuRenameCD.Visible = False
   Else
    If mNode.Parent Is Nothing Then
         mnuRenameCD.Visible = True
         mnuRenameCD.Tag = mNode.Tag
        Else
         mnuRenameCD.Visible = False
    End If
End If

'Remove CD Menu
    mnuRemoveCD.Enabled = False
    If Not mNode Is Nothing Then
        mnuRemoveCD.Enabled = True
        'if one CD is selected
        smnuRemoveCD(1).Caption = CDs.Nodes(GetCDNode(mNode.Index)).Text
        smnuRemoveCD(1).Tag = CDs.Nodes(GetCDNode(mNode.Index)).Tag
    Else
        ' if no CD is selected
        Set rsCD = xDb.OpenRecordset("cds")
        i = 1
        If rsCD.RecordCount <> 0 Then
            mnuRemoveCD.Enabled = True
            Do Until rsCD.EOF
                smnuRemoveCD(i).Caption = rsCD("CDLabel")
                smnuRemoveCD(i).Tag = rsCD("CDSerial")
                rsCD.MoveNext
                i = 1 + 1
                If Not rsCD.EOF Then Load smnuRemoveCD(i)
            Loop
         End If
        rsCD.Close
    End If
    
'Add CD Menu
    For i = 1 To UBound(CDDrives)
        smnuAddCD(i).Caption = CDDrives(i).CDDrive & " : " & CDDrives(i).CDLabel
    Next i
        
    
'Rescan CD Menu
mnuRescan.Enabled = True
    If Not mNode Is Nothing Then
        If Not mNode.Parent Is Nothing Then
            'it's not select CD Node
            Load smnuRescan(2)
            If mNode.Children > 0 Then
                'it's select Folder
                smnuRescan(2).Caption = "Folder" & vbTab & ": " & mNode.Text
                smnuRescan(2).Tag = mNode.Tag
                smnuRescan(1).Caption = "CD" & vbTab & ": " & CDs.Nodes(GetCDNode(mNode.Index)).Text
                smnuRescan(1).Tag = CDs.Nodes(GetCDNode(mNode.Index)).Tag
            End If
            If mNode.Children = 0 Then
                'it's select one File
                Load smnuRescan(3)
                smnuRescan(1).Caption = "CD" & vbTab & ": " & CDs.Nodes(GetCDNode(mNode.Index)).Text
                smnuRescan(1).Tag = CDs.Nodes(GetCDNode(mNode.Index)).Tag
                smnuRescan(2).Caption = "Folder" & vbTab & ": " & mNode.Parent.Text
                smnuRescan(2).Tag = mNode.Parent.Tag
                smnuRescan(3).Caption = "File" & vbTab & ": " & mNode.Text
                smnuRescan(3).Tag = mNode.Tag
            End If
        Else
                smnuRescan(1).Caption = "CD" & vbTab & ": " & mNode.Text
                smnuRescan(1).Tag = mNode.Tag
        End If
    Else
        Set rsCD = xDb.OpenRecordset("cds")
        i = 1
        If rsCD.RecordCount > 0 Then
            Do Until rsCD.EOF
                smnuRescan(i).Caption = "CD" & vbTab & ": " & rsCD("CDLabel")
                smnuRescan(i).Tag = rsCD("CDSerial")
                rsCD.MoveNext
                i = 1 + 1
                If Not rsCD.EOF Then Load smnuRescan(i)
            Loop
           Else
            mnuRescan.Enabled = False
        End If
        rsCD.Close
    End If
    
Set rsCD = Nothing
'Show Popup menu
    PopupMenu mnuCDs, , x / 15, Y / 15
End If

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub CDs_NodeClick(ByVal Node As MSComctlLib.Node)
ShowContent (Node.Index)
FindFileDetails (Node.Tag)
End Sub



Private Sub FileDescription_LostFocus()
Dim rsFiles As Recordset

On Error GoTo err1
Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & SelectedFileID)

rsFiles.Edit
rsFiles("Description") = FileDescription.Text
rsFiles.Update
rsFiles.Close
Set rsFiles = Nothing

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub FileKeyWords_LostFocus()
Dim rsFiles As Recordset

On Error GoTo err1
Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & SelectedFileID)

rsFiles.Edit
rsFiles("keywords") = FileKeyWords.Text & ""
rsFiles.Update
rsFiles.Close
Set rsFiles = Nothing

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub Form_Load()
Dim x
Dim Temp As String
Dim FormLeft, FormTop
Dim FormWidth, FormHeight

On Error GoTo err1
StatusBar.Panels.Add
StatusBar.Panels(1).Width = 200
StatusBar.Panels(2).Width = Me.Width - 200

ThumbSlide.BackColor = vbButtonFace
ThumbWidth = 120
ThumbHeight = 120
ThumbRow = 1
ThumbCol = 0


'set left
Temp = Space(255)
x = GetPrivateProfileString("Settings", "Left", 0, Temp, 255, App.Path & "\Archive.ini")
FormLeft = Mid(Temp, 1, x)
Me.Left = FormLeft
'set top
Temp = Space(255)
x = GetPrivateProfileString("Settings", "Top", 0, Temp, 255, App.Path & "\Archive.ini")
FormTop = Mid(Temp, 1, x)
Me.Top = FormTop
'set Width
Temp = Space(255)
x = GetPrivateProfileString("Settings", "Width", 600, Temp, 255, App.Path & "\Archive.ini")
FormWidth = Mid(Temp, 1, x)
Me.Width = FormWidth
'set Height
Temp = Space(255)
x = GetPrivateProfileString("Settings", "Height", 400, Temp, 255, App.Path & "\Archive.ini")
FormHeight = Mid(Temp, 1, x)
Me.Height = FormHeight


Me.Show
StatusBar.Panels(2).Text = "Right click on the left pane."

Form_Resize
GetDrivesInfo

DBPath = Space(255)
x = GetPrivateProfileString("Settings", "DefaultDB", App.Path & "\Archive.mdb", DBPath, 255, App.Path & "\Archive.ini")
DBPath = Mid(DBPath, 1, x)

If Dir(DBPath) = "" Then
    If Not CreateDB(DBPath) Then End
End If
Set xDb = DBEngine.OpenDatabase(DBPath, , False)

RefreshCDList

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub Form_Resize()

On Error GoTo err1
If Me.Width / 15 <= 600 Then Me.Width = 600 * 15
If Me.Height / 15 <= 400 Then Me.Height = 400 * 15

CDs.Top = 0
CDs.Left = 0
CDs.Width = 250
CDs.Height = Me.Height / 15 - FileDetails.Height - StatusBar.Height - 47

FileDetails.Top = CDs.Height
FileDetails.Left = 0
FileDetails.Width = 250
FileDetails.Height = 150

ThumbFrame.Top = 80
ThumbFrame.Left = CDs.Width
ThumbFrame.Width = Me.Width / 15 - CDs.Width - 8
ThumbFrame.Height = Me.Height / 15 - StatusBar.Height - 47 - ThumbFrame.Top

PicThumbnail.Top = ThumbFrame.Top + 5
PicThumbnail.Left = ThumbFrame.Left + 5

ThumbScroll.Top = 0
ThumbScroll.Left = ThumbFrame.Width - ThumbScroll.Width - 3
ThumbScroll.Height = ThumbFrame.Height - 5
ThumbScroll.SmallChange = ThumbWidth
ThumbScroll.LargeChange = ThumbFrame.Height

PB.Top = Me.Height / 15 - 59
PB.Left = 3
PB.Width = 195

ReSortThumbnails

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub


Private Sub Form_Unload(Cancel As Integer)

If Dir(App.Path & "\Archive.ini") = "" Then frmAbout.Show (1)

WritePrivateProfileString "Settings", "DefaultDB", xDb.Name, App.Path & "\Archive.ini"
WritePrivateProfileString "Settings", "Left", Me.Left, App.Path & "\Archive.ini"
WritePrivateProfileString "Settings", "Top", Me.Top, App.Path & "\Archive.ini"
WritePrivateProfileString "Settings", "Width", Me.Width, App.Path & "\Archive.ini"
WritePrivateProfileString "Settings", "Height", Me.Height, App.Path & "\Archive.ini"

Set xDb = Nothing

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuRenameCD_Click()
Dim CDSerial
Dim rsCD As Recordset
Dim NodeID

On Error GoTo err1
CDSerial = mnuRenameCD.Tag

Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDSerial=" & CDSerial)

rsCD.Edit
rsCD("CDLabel") = InputBox("Enter New CD Label", "Rename CD", rsCD("CDLabel"))
rsCD.Update
CDs.Nodes("ID " & rsCD("CDID")).Text = rsCD("CDLabel") '& " [" & rsCD("CDSerial") & "]"
rsCD.Close
Set rsCD = Nothing


'RefreshCDList

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub


Private Sub smnuAddCD_Click(Index As Integer)
Dim CDID
Dim rsCD As Recordset

On Error GoTo err1
Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDSerial=" & CDDrives(Index).CDSerial)
If rsCD.RecordCount > 0 Then
    MsgBox "CD alredy catalogized"
    Exit Sub
End If

If CDDrives(Index).CDLabel = "" Then CDDrives(Index).CDLabel = "No Label"
rsCD.AddNew
rsCD("CDLabel") = CDDrives(Index).CDLabel
rsCD("CDSerial") = CDDrives(Index).CDSerial
CDID = rsCD("CDID")
rsCD.Update
rsCD.Close
Set rsCD = Nothing

RefreshCDList
AddFiles CDDrives(Index).CDDrive, CDID

Exit Sub
err1:
MsgBox Error


End Sub

Private Sub RefreshCDList()
Dim mNode As Node
Dim rsCD As Recordset


On Error GoTo err1
CDs.Nodes.Clear

Set rsCD = xDb.OpenRecordset("CDs")

Do Until rsCD.EOF
    Set mNode = CDs.Nodes.Add(, tvwAutomatic, "ID " & rsCD("cdid"), rsCD("CDLabel") & " [" & rsCD("CDSerial") & "]", 1)
    mNode.Tag = rsCD("CDSerial")
    rsCD.MoveNext
Loop
TotalCDs = rsCD.RecordCount
rsCD.Close
Set rsCD = Nothing

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub AddFiles(sPath As String, CDID)
Dim Files() As tFileInfo
Dim i, DirIndex
Dim rsFiles As Recordset

On Error GoTo err1

Set rsFiles = xDb.OpenRecordset("files")

ReDim Directories(0)
Directories(0) = sPath
DirIndex = 0

'do for directories
Do While UBound(Directories) <> DirIndex - 1
FindFiles Directories(DirIndex), Files()
DirIndex = DirIndex + 1
For i = 0 To UBound(Files)
    If Files(i).fNAME = "" Then
        'empty directory
'        xDb.Execute "DELETE * FROM Files WHERE CDID=" & cdid & " and Fullpath='" & Mid(Directories(DirIndex - 1), 3) & "'"
        GoTo NextFile
    End If

StatusBar.Panels(2).Text = "Reading CD " & Files(i).FullPath
DoEvents
    
    If Files(i).fSize = "" Then
        ReDim Preserve Directories(LBound(Directories) To UBound(Directories) + 1)
        Directories(UBound(Directories)) = Files(i).FullPath
        
        'it's a directory
        rsFiles.AddNew
        rsFiles("CDID") = CDID
        rsFiles("Name") = Files(i).fNAME
        rsFiles("fullpath") = Mid(Files(i).FullPath, 3)
        rsFiles("type") = 1
        rsFiles.Update
         
       Else
        
        'LoadFile (Files(i).FullPath)
        'SavePicture ThumbNail(ThumbNail.Count - 1).Picture, App.Path & "\temp.bmp"
        'DoEvents

        'ConvertBMPtoJPG App.Path & "\temp.bmp", App.Path & "\temp.jpg", True, 60, True
        
        
        rsFiles.AddNew
        rsFiles("CDID") = CDID
        rsFiles("Name") = Files(i).fNAME
        rsFiles("fullpath") = Mid(Files(i).FullPath, 3)
        rsFiles("type") = 0
        'save the image into db
'            Open App.Path & "\temp.jpg" For Binary Access Read As #1
'            ReDim TheBytes(FileLen(App.Path & "\temp.jpg") - 1)
'            Get #1, , TheBytes()
'            Close #1
'        rsFiles("Thumb").AppendChunk TheBytes
        'rsFiles("Dimension") = PicLoadWidth & " x " & PicLoadHeight
        rsFiles.Update
        
    End If
NextFile:
Next i
ReDim Files(0 To 0)
Loop

Set rsFiles = xDb.OpenRecordset("DirsToDelete", dbOpenDynaset)
Do Until rsFiles.EOF
    xDb.Execute ("DELETE * FROM Files WHERE ID=" & rsFiles("ID"))
    rsFiles.MoveNext
Loop
rsFiles.Close
Set rsFiles = Nothing

StatusBar.Panels(2).Text = ""
RescanCD (GetCDSerial(CDID))

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub smnuCompactDB_Click()
Dim BackupDB
'if exist the mdb file delete the backup
SetHourglass (True)

On Error GoTo err1
BackupDB = Mid(DBPath, 1, Len(DBPath) - 3) & "bac"
If Dir(DBPath) <> "" Then
    If Dir(BackupDB) <> "" Then Kill BackupDB
    xDb.Close
    Name DBPath As BackupDB
    DBEngine.CompactDatabase BackupDB, DBPath
    Set xDb = DBEngine.OpenDatabase(DBPath, , False)
   Else
    If Dir(BackupDB) <> "" Then
        Name BackupDB As DBPath
       Else
        MsgBox "Archive Database is missing." & Chr(10) & "New DB will be created."
        CreateDB (DBPath)
    End If
End If
SetHourglass (False)

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub smnuExit_Click()
Unload Me
End Sub

Private Sub smnuOpenDB_Click()
Dim Temp

On Error GoTo err1
SelectDB:
CD.Filter = "Microsoft Access Database|*mdb"
CD.FileName = ""
CD.ShowOpen

'No file selected
If CD.FileName = "" Then Exit Sub


Set xDb = DBEngine.OpenDatabase(CD.FileName)
On Error Resume Next
Temp = xDb.TableDefs("Files").Name
Temp = xDb.TableDefs("CDs").Name
Temp = xDb.QueryDefs("DirsToDELETE").Name
Temp = xDb.QueryDefs("DirsToLeave").Name

'Wrong Database
If Err Then
    MsgBox "Incorect database!"
    GoTo SelectDB
End If
On Error GoTo 0

DBPath = CD.FileName
Set xDb = DBEngine.OpenDatabase(CD.FileName)
RefreshCDList

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub smnuRemoveCD_Click(Index As Integer)
Dim CDSerial, CDID

On Error GoTo err1

If MsgBox("If you remove the CD you lost all informations", vbOKCancel) = vbCancel Then Exit Sub
CDSerial = smnuRemoveCD(Index).Tag
CDID = GetCDID(CDSerial)
xDb.Execute "DELETE * FROM Files WHERE CDID=" & CDID
xDb.Execute "DELETE * FROM CDs WHERE CDID=" & CDID
RefreshCDList
ClearThumbnails

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub
Private Function GetCDID(CDSerial)
Dim rsCD As Recordset

On Error GoTo err1

Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDSerial=" & CDSerial)
GetCDID = rsCD("CDID")
rsCD.Close
Set rsCD = Nothing

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Sub AddChildNodes(mNode As Node, CDID)
Dim ParentNode As Node
Dim NewNode As Node
Dim ParentNodeID
Dim NodeText
Dim ImageType
Dim rsFiles As Recordset

On Error GoTo err1

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE CDID=" & CDID & " ORDER BY Type DESC, FullPath")

If rsFiles.RecordCount = 0 Then GoTo ExitSub
rsFiles.MoveLast
rsFiles.MoveFirst

PB.Value = 0
PB.Visible = True
PB.Max = 100
Set ParentNode = mNode

Do Until rsFiles.EOF
    ParentNodeID = NodeID(GetParentDir(mNode.Key & rsFiles("Fullpath")))
    If ParentNodeID > 0 Then
        Set ParentNode = CDs.Nodes(ParentNodeID)
    End If
    NodeText = rsFiles("Description") & ""
    If NodeText = "" Then NodeText = rsFiles("Name")
    Set NewNode = CDs.Nodes.Add(ParentNode, tvwChild, mNode.Key & rsFiles("fullpath"), NodeText)
    If rsFiles("type") = 0 Then
        If rsFiles("dimension") <> "" Then
            NewNode.Image = 4
           Else
            NewNode.Image = 5
        End If
       Else
        NewNode.Image = 2
        NewNode.ExpandedImage = 3
    End If
    NewNode.Tag = rsFiles("ID")
        
    PB.Value = rsFiles.PercentPosition
    rsFiles.MoveNext
Loop

ExitSub:
PB.Visible = False
rsFiles.Close
Set rsFiles = Nothing

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub
Private Function GetParentDir(sPath)
Dim i

On Error GoTo err1
If Right(sPath, 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
i = InStrRev(sPath, "\")
GetParentDir = Mid(sPath, 1, i - 1)

Exit Function
err1:
MsgBox Error
Exit Function

End Function
Private Function NodeID(NodeKey)
NodeID = -1
On Error Resume Next
NodeID = CDs.Nodes(NodeKey).Index
On Error GoTo 0
End Function
Private Function LoadFile(FileName)
Dim NewThumbnailNo

On Error GoTo err1
Set PicLoad.Picture = LoadPicture()
Set PicLoad.Picture = LoadPicture(FileName)
CreateThumbnail PicLoad, PicThumbnail
PicLoadWidth = PicLoad.Width
PicLoadHeight = PicLoad.Height

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Sub CreateThumbnail(picSource As PictureBox, PicThumb As PictureBox)

Dim lRet            As Long
Dim lLeft           As Long
Dim lTop            As Long
Dim lWidth          As Long
Dim lHeight         As Long
Dim lForeColor      As Long
Dim hBrush          As Long
Dim hDummyBrush     As Long
Dim lOrigMode       As Long
Dim fScale          As Single
Dim uBrushOrigPt    As PointAPI


On Error GoTo err1

    PicThumb.Width = ThumbWidth
    PicThumb.Height = ThumbHeight
    PicThumb.BackColor = vbButtonFace
    PicThumb.AutoRedraw = True
    PicThumb.Cls
    
    If picSource.Width <= PicThumb.Width - 2 And picSource.Height <= PicThumb.Height - 2 Then
        fScale = 1
    Else
        If picSource.Width > picSource.Height Then
            fScale = (PicThumb.Width - 14) / picSource.Width
        Else
            fScale = (PicThumb.Height - 24) / picSource.Height
    End If
    End If
    lWidth = picSource.Width * fScale
    lHeight = picSource.Height * fScale
    If picSource.Width > picSource.Height Then
        lTop = Int((PicThumb.Height - lHeight) / 2)
        lLeft = 5
       Else
        lTop = 12
        lLeft = Int((PicThumb.Width - lWidth) / 2) - 2
    End If
    'Store the original ForeColor
    lForeColor = PicThumb.ForeColor
    
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(PicThumb.hDC, STRETCH_HALFTONE)
    
    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(PicThumb.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    lRet = SetBrushOrgEx(PicThumb.hDC, lLeft, lTop, uBrushOrigPt)
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(PicThumb.hDC, hBrush)
    
    'Stretch the bitmap
    lRet = StretchBlt(PicThumb.hDC, lLeft, lTop, lWidth, lHeight, _
            picSource.hDC, 0, 0, picSource.Width, picSource.Height, SRCCOPY)
    
    'Set the stretch mode back to it's original mode
    lRet = SetStretchBltMode(PicThumb.hDC, lOrigMode)
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(PicThumb.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set the brush alignment back to the original coordinates
    lRet = SetBrushOrgEx(PicThumb.hDC, uBrushOrigPt.x, uBrushOrigPt.Y, uBrushOrigPt)
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(PicThumb.hDC, hBrush)
    'Get rid of the dummy brush
    lRet = DeleteObject(hDummyBrush)
    
    'Restore the original ForeColor
    PicThumb.ForeColor = lForeColor

    PicThumb.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Function NewThumbnail()
Dim NewThumbNo

On Error GoTo err1

NewThumbNo = ThumbNail.Count
Load ThumbNail(NewThumbNo)


If (ThumbCol + 1) * ThumbWidth <= ThumbFrame.Width - ThumbScroll.Width Then
    ThumbCol = ThumbCol + 1
   Else
    ThumbCol = 1
    ThumbRow = ThumbRow + 1
End If

If ThumbCols < ThumbCol Then ThumbCols = ThumbCol
If ThumbRows < ThumbRow Then ThumbRows = ThumbRow

ThumbSlide.Width = ThumbCols * ThumbWidth
ThumbSlide.Height = ThumbRow * ThumbHeight

ThumbNail(NewThumbNo).Left = (ThumbCol - 1) * ThumbWidth
ThumbNail(NewThumbNo).Top = (ThumbRow - 1) * ThumbHeight
ThumbNail(NewThumbNo).Height = ThumbHeight
ThumbNail(NewThumbNo).Width = ThumbWidth

NewThumbnail = NewThumbNo

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Sub ReSortThumbnails()
Dim i

On Error GoTo err1
ThumbCol = 0
ThumbRow = 1
ThumbCols = 1
ThumbRows = 1


For i = 1 To ThumbNail.Count - 1

If (ThumbCol + 1) * ThumbWidth <= ThumbFrame.Width - ThumbScroll.Width - 3 Then
    ThumbCol = ThumbCol + 1
   Else
    ThumbCol = 1
    ThumbRow = ThumbRow + 1
End If

If ThumbCols < ThumbCol Then ThumbCols = ThumbCol
If ThumbRows < ThumbRow Then ThumbRows = ThumbRow

ThumbSlide.Width = ThumbCols * ThumbWidth
ThumbSlide.Height = ThumbRow * ThumbHeight

ThumbNail(i).Left = (ThumbCol - 1) * ThumbWidth
ThumbNail(i).Top = (ThumbRow - 1) * ThumbHeight
ThumbNail(i).Height = ThumbHeight
ThumbNail(i).Width = ThumbWidth

Next i

SetThumbScroll

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub
Private Sub ClearThumbnails()
Dim i

On Error GoTo err1

ThumbRows = 1
ThumbCols = 1
ThumbRow = 1
ThumbCol = 0


For i = 1 To ThumbNail.Count - 1
    Unload ThumbNail(i)
Next i

SelectedFileID = ""
ClearFileDetails
ReSortThumbnails

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub
Private Function GetFileID(NodeID)
Dim Temp, FileID

On Error GoTo err1

NodeID = CLng(NodeID)
GetFileID = CDs.Nodes(NodeID).Tag

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Sub ShowThumbnails(FileID)
Dim i, tempFile
Dim rsFiles As Recordset
Dim ThumbNo

On Error GoTo err1

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)

If rsFiles.EOF Then Exit Sub
If rsFiles("type") <> 0 Then Exit Sub


If gSelectedThumbNail = 0 Then
    ThumbNo = NewThumbnail()
   Else
    ThumbNo = gSelectedThumbNail
End If

ThumbNail(ThumbNo).Tag = FileID 'NodeIndex
ThumbNail(ThumbNo).Caption = Mid(rsFiles("Name"), InStrRev(rsFiles("Name"), "\") + 1)
ThumbNail(ThumbNo).Visible = True

If rsFiles("Thumb").FieldSize > 0 Then
    ReDim TheBytes(rsFiles("Thumb").FieldSize)
    TheBytes() = rsFiles("Thumb").GetChunk(0, rsFiles("Thumb").FieldSize)
    
    Open App.Path & "\temp.bmp" For Binary Access Write As 1
    Put #1, , TheBytes()
    Close #1
    ThumbNail(ThumbNo).Picture = LoadPicture(App.Path & "\temp.bmp")
End If

rsFiles.Close
Set rsFiles = Nothing

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub FindFileDetails(FileID)
Dim rsFiles As Recordset


On Error GoTo err1
If FileID = "" Then Exit Sub

SelectedFileID = FileID
Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)
If rsFiles.RecordCount > 0 Then
    If rsFiles("type") = 1 Then
        ClearFileDetails
       Else
        FileDetails.Enabled = True
        FileKeyWords.Text = rsFiles("KeyWords") & ""
        FileDescription.Text = rsFiles("description") & ""
    End If
   Else
    ClearFileDetails
End If
rsFiles.Close
Set rsFiles = Nothing

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub smnuRescan_Click(Index As Integer)
Dim CDID, FileID, FolderID

If Left(smnuRescan(Index).Caption, 4) = "CD" & vbTab & ":" Then RescanCD (smnuRescan(Index).Tag)
If Left(smnuRescan(Index).Caption, 6) = "File" & vbTab & ":" Then
    ClearThumbnails
    RescanFile (smnuRescan(Index).Tag)
End If
If Left(smnuRescan(Index).Caption, 8) = "Folder" & vbTab & ":" Then RescanFolder (smnuRescan(Index).Tag)
If smnuRescan(Index).Caption = "Rescan this file" Then RescanFile (smnuRescan(Index).Tag)
    
End Sub

Private Sub ThumbNail_Click(Index As Integer)
Dim CDID
Dim mNode As Node
Dim FileID, i
Dim ClickedNodeID

On Error GoTo err1

FileID = ThumbNail(Index).Tag

FindFileDetails FileID

CDID = GetCDIDFromDB(FileID)
Set mNode = CDs.Nodes("ID " & CDID)
If mNode.Children = 0 Then
    AddChildNodes mNode, CDID
End If

'disable to show content of folder
gShowContent = False
CDs.SetFocus
ClickedNodeID = CDs.Nodes(GetFullPathFromDB(ThumbNail(Index).Tag)).Index

CDs.Nodes(ClickedNodeID).Selected = True
mNode.Selected = True
CDs.Nodes(ClickedNodeID).Selected = True
CDs.SetFocus

'enable to show content of folder
gShowContent = True

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Sub ThumbNail_DblClick(Index As Integer)
Dim rsCD As Recordset
Dim rsFiles As Recordset
Dim ScanPath As String
Dim FileID
Dim CDID, CDSerial


FileID = ThumbNail(Index).Tag
If UBound(CDDrives) = 0 Then Exit Sub
ScanPath = CDDrives(1).CDDrive

If Right(ScanPath, 1) = "\" Then ScanPath = Mid(ScanPath, 1, Len(ScanPath) - 1)

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)
CDID = rsFiles("CDID")
rsFiles.Close
Set rsFiles = Nothing


'Wait for user to insert the CD into drive
Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDID=" & CDID)
CDSerial = rsCD("CDSerial")

TryCDAgain:
If GetDriveSerial(ScanPath) <> CDSerial Then
    If MsgBox("Insert CD '" & rsCD("CDLabel") & "' into drive " & UCase(Left(ScanPath, 2)), vbOKCancel) = vbCancel Then Exit Sub
    GoTo TryCDAgain
End If
rsCD.Close
Set rsCD = Nothing

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)

Set frmPicturePreview.OriginalPicture = LoadPicture(ScanPath & rsFiles("FullPath"))

rsFiles.Close
Set rsFiles = Nothing

frmPicturePreview.Show
End Sub

Private Sub ThumbNail_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim BaseLeft, BaseTop, i

If Button = 1 Then Exit Sub

On Error GoTo err1

ThumbNail(Index).SetFocus
For i = 2 To smnuRescan.Count: Unload smnuRescan(i): Next i

Load smnuRescan(2)
Load smnuRescan(3)

smnuRescan(1).Caption = "Cancel"
smnuRescan(2).Caption = "-"
smnuRescan(3).Caption = "Rescan this file"
smnuRescan(3).Tag = ThumbNail(Index).Tag

gSelectedThumbNail = Index

BaseLeft = ThumbNail(Index).Left + ThumbFrame.Left
BaseTop = ThumbNail(Index).Top + ThumbFrame.Top - ThumbScroll.Value
PopupMenu mnuRescan, , BaseLeft + x / 15, BaseTop + Y / 15


gSelectedThumbNail = 0
Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Function GetCDNode(NodeID)
Dim CDNode As Node

Set CDNode = CDs.Nodes(NodeID)
Do Until CDNode.Parent Is Nothing
Set CDNode = CDNode.Parent
Loop
GetCDNode = CDNode.Index
End Function

Private Sub RescanCD(CDSerial)
Dim ScanPath As String
Dim CDID, i
Dim rsFiles As Recordset
Dim rsCD As Recordset
Dim LastPath
Dim CurrentNodeID
Dim CurrentNode As Node
Dim SelectedNode As Node

On Error GoTo err1

If UBound(CDDrives) = 0 Then Exit Sub
ScanPath = CDDrives(1).CDDrive


'Wait for user to insert the CD into drive
Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDserial=" & CDSerial)
CDID = rsCD("CDID")
TryCDAgain:
If GetDriveSerial(ScanPath) <> CDSerial Then
    If MsgBox("Insert CD '" & rsCD("CDLabel") & "' into drive " & UCase(Left(ScanPath, 2)), vbOKCancel) = vbCancel Then Exit Sub
    GoTo TryCDAgain
End If
rsCD.Close
Set rsCD = Nothing

If CDs.Nodes("ID " & CDID).Children = 0 Then
    AddChildNodes CDs.Nodes("ID " & CDID), CDID
End If

Set SelectedNode = CDs.Nodes("ID " & CDID)

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE CDID=" & CDID & " and type=0 ORDER By FullPath")
If Right(ScanPath, 1) = "\" Then ScanPath = Mid(ScanPath, 1, Len(ScanPath) - 1)


rsFiles.MoveLast
rsFiles.MoveFirst

PB.Max = 100
PB.Value = 0
PB.Visible = True

i = 0
gCancel = False
SetStop True
'To collapse expanded and not actual nodes
CDs.SingleSel = True
Set CurrentNode = CDs.Nodes("ID " & CDID)

SetHourglass (True)

Do Until rsFiles.EOF
    StatusBar.Panels(2).Text = "Reading CD " & ScanPath & rsFiles("fullpath")
   
    If IsNull(rsFiles("Thumb")) Then
        LoadFile (ScanPath & rsFiles("fullpath"))
        SavePicture PicThumbnail.Image, App.Path & "\temp.bmp"
        ConvertBMPtoJPG App.Path & "\temp.bmp", App.Path & "\temp.jpg", True, 60, True
    
        rsFiles.Edit
    
        Open App.Path & "\temp.jpg" For Binary Access Read As #1
        ReDim TheBytes(FileLen(App.Path & "\temp.jpg") - 1)
        Get #1, , TheBytes()
        Close #1
        rsFiles("Thumb").AppendChunk TheBytes
        rsFiles("Dimension") = PicLoadWidth & " x " & PicLoadHeight
        rsFiles.Update
    End If
    If gCancel Then GoTo ExitSub
    DoEvents
    
    CurrentNodeID = CDs.Nodes("ID " & rsFiles("CDID") & rsFiles("fullpath")).Index
    'Change icon
    CDs.Nodes(CurrentNodeID).Image = 4
    
    If CurrentNode <> CDs.Nodes(CurrentNodeID).Parent Then
        ClearThumbnails
        CDs.Nodes(CurrentNodeID).Parent.Selected = True
        Set CurrentNode = CDs.Nodes(CurrentNodeID).Parent
        CurrentNode.Expanded = False
    End If
    ShowThumbnails rsFiles("ID")
    SetThumbScroll
    ThumbScroll.Value = ThumbScroll.Max
    DoEvents
    
    PB.Value = rsFiles.PercentPosition
    i = i + 1
    
    rsFiles.MoveNext
Loop

ExitSub:

SelectedNode.Selected = True
SelectedNode.Expanded = False
gDirShown = ""
ShowContent (SelectedNode.Index)

CDs.SingleSel = False
gDirShown = ""
PB.Visible = False
SetStop (False)
SetHourglass (False)
rsFiles.Close
Set rsFiles = Nothing

StatusBar.Panels(2).Text = ""

Exit Sub
err1:
MsgBox Error
Resume
Exit Sub

End Sub
Private Sub RescanFolder(FolderID)
Dim ScanPath As String
Dim CDSerial, CDID
Dim FolderPath
Dim i
Dim rsFiles As Recordset
Dim rsCD As Recordset
Dim CurrentNodeID
Dim CurrentNode As Node
Dim SelectedNode As Node

On Error GoTo err1

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FolderID)

If rsFiles.EOF Then
    'if no folder found, it's a root directory of the CD
    Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDSerial=" & FolderID)
    CDID = rsCD("CDID")
    FolderPath = ""
    rsCD.Close
   Else
   'it's a folder
    CDID = rsFiles("CDID")
    FolderPath = rsFiles("FullPath")
End If
rsFiles.Close

Set SelectedNode = CDs.Nodes("ID " & CDID & FolderPath)

If UBound(CDDrives) = 0 Then Exit Sub
ScanPath = CDDrives(1).CDDrive

'Wait for user to insert the CD into drive
Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDID=" & CDID)
CDSerial = rsCD("CDSerial")

TryCDAgain:
If GetDriveSerial(ScanPath) <> CDSerial Then
    If MsgBox("Insert CD '" & rsCD("CDLabel") & "' into drive " & UCase(Left(ScanPath, 2)), vbOKCancel) = vbCancel Then Exit Sub
    GoTo TryCDAgain
End If
rsCD.Close
Set rsCD = Nothing

If FolderPath = "" Then
    Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE CDID=" & CDID & " and type=0 and FullPath like '\' & [Name] ORDER By fullpath")
   Else
    Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE CDID=" & CDID & " and type=0 and FullPath like '" & FolderPath & "\*' & [Name] ORDER By fullpath")
End If
If Right(ScanPath, 1) = "\" Then ScanPath = Mid(ScanPath, 1, Len(ScanPath) - 1)


rsFiles.MoveLast
rsFiles.MoveFirst

PB.Value = 0
PB.Max = 100
PB.Visible = True

i = 0
gCancel = False
SetStop (True)
SetHourglass (True)
Set CurrentNode = CDs.Nodes("ID " & CDID & FolderPath)
'To collapse expanded and not actual nodes
CDs.SingleSel = True
ClearThumbnails

Do Until rsFiles.EOF
    StatusBar.Panels(2).Text = "Reading CD " & ScanPath & rsFiles("fullpath")
    
    If IsNull(rsFiles("Thumb")) Then
        LoadFile (ScanPath & rsFiles("fullpath"))
        SavePicture PicThumbnail.Image, App.Path & "\temp.bmp"
        ConvertBMPtoJPG App.Path & "\temp.bmp", App.Path & "\temp.jpg", True, 60, True
        
        rsFiles.Edit
        
        Open App.Path & "\temp.jpg" For Binary Access Read As #1
        ReDim TheBytes(FileLen(App.Path & "\temp.jpg") - 1)
        Get #1, , TheBytes()
        Close #1
        rsFiles("Thumb").AppendChunk TheBytes
        rsFiles("Dimension") = PicLoadWidth & " x " & PicLoadHeight
        rsFiles.Update
    End If
    
    If gCancel Then GoTo ExitSub
    DoEvents
        
    CurrentNodeID = CDs.Nodes("ID " & rsFiles("CDID") & rsFiles("fullpath")).Index
    'Change icon
    CDs.Nodes(CurrentNodeID).Image = 4
    
    If CurrentNode.FullPath <> CDs.Nodes(CurrentNodeID).Parent.FullPath Then
        ClearThumbnails
        CDs.Nodes(CurrentNodeID).Parent.Selected = True
        Set CurrentNode = CDs.Nodes(CurrentNodeID).Parent
        CurrentNode.Expanded = False
    End If
    gSelectedThumbNail = 0
    ShowThumbnails rsFiles("ID")
    SetThumbScroll
    ThumbScroll.Value = ThumbScroll.Max
    DoEvents
    
        
    PB.Value = rsFiles.PercentPosition
    i = i + 1
    rsFiles.MoveNext
Loop

ExitSub:

SelectedNode.Selected = True
SelectedNode.Expanded = False
gDirShown = ""
ShowContent (SelectedNode.Index)

CDs.SingleSel = False
PB.Visible = False
SetStop (False)
SetHourglass (False)
rsFiles.Close
Set rsFiles = Nothing

StatusBar.Panels(2).Text = ""

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub
Private Sub RescanFile(FileID)
Dim ScanPath As String
Dim CDSerial, CDID
Dim FilePath, i
Dim rsFiles As Recordset
Dim rsCD As Recordset

On Error GoTo err1

Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)
CDID = rsFiles("CDID")
FilePath = rsFiles("FullPath")
rsFiles.Close

If UBound(CDDrives) = 0 Then Exit Sub
ScanPath = CDDrives(1).CDDrive

'Wait for user to insert the CD into drive
Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDID=" & CDID)
CDSerial = rsCD("CDSerial")

TryCDAgain:
If GetDriveSerial(ScanPath) <> CDSerial Then
    If MsgBox("Insert CD '" & rsCD("CDLabel") & "' into drive " & UCase(Left(ScanPath, 2)), vbOKCancel) = vbCancel Then Exit Sub
    GoTo TryCDAgain
End If
rsCD.Close
Set rsCD = Nothing

'if is the CD OK rescan picture
'ClearThumbnails
Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)
If Right(ScanPath, 1) = "\" Then ScanPath = Mid(ScanPath, 1, Len(ScanPath) - 1)


SetHourglass (True)
StatusBar.Panels(2).Text = "Reading CD " & ScanPath & rsFiles("fullpath")
DoEvents

LoadFile (ScanPath & rsFiles("fullpath"))
SavePicture PicThumbnail.Image, App.Path & "\temp.bmp"
ConvertBMPtoJPG App.Path & "\temp.bmp", App.Path & "\temp.jpg", True, 60, True

rsFiles.Edit

Open App.Path & "\temp.jpg" For Binary Access Read As #1
ReDim TheBytes(FileLen(App.Path & "\temp.jpg") - 1)
Get #1, , TheBytes()
Close #1

rsFiles("Thumb").AppendChunk TheBytes
rsFiles("Dimension") = PicLoadWidth & " x " & PicLoadHeight
rsFiles.Update

CDs.Nodes("ID " & rsFiles("CDID") & rsFiles("FullPath")).Image = 4

rsFiles.Close
Set rsFiles = Nothing

ShowThumbnails (FileID)
FindFileDetails FileID

SetHourglass (False)
StatusBar.Panels(2).Text = ""

Exit Sub
err1:
MsgBox Error
Exit Sub

End Sub

Private Function GetCDSerial(CDID)
Dim rsCD As Recordset

On Error GoTo err1
Set rsCD = xDb.OpenRecordset("SELECT * FROM CDs WHERE CDID=" & CDID)
GetCDSerial = rsCD("CDSerial")
rsCD.Close
Set rsCD = Nothing

Exit Function
err1:
MsgBox Error
Exit Function

End Function
Private Sub SetHourglass(Value As Boolean)

If Value Then
    Me.MousePointer = 11
    Me.CDs.MousePointer = 11
   Else
    Me.MousePointer = 0
    Me.CDs.MousePointer = 0
End If
DoEvents
End Sub
Private Sub ClearFileDetails()
FileKeyWords.Text = ""
FileDescription.Text = ""
FileDetails.Enabled = False
End Sub
Private Function GetFullPathFromDB(FileID)
Dim rsFiles As Recordset

On Error GoTo err1
Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)
If rsFiles.RecordCount = 0 Then
    GetFullPathFromDB = ""
    Exit Function
End If

GetFullPathFromDB = "ID " & rsFiles("CDID") & rsFiles("FullPath")
rsFiles.Close
Set rsFiles = Nothing

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Function GetCDIDFromDB(FileID)
Dim rsFiles As Recordset

On Error GoTo err1
Set rsFiles = xDb.OpenRecordset("SELECT * FROM Files WHERE ID=" & FileID)
If rsFiles.RecordCount = 0 Then
    GetCDIDFromDB = ""
    Exit Function
End If

GetCDIDFromDB = rsFiles("CDID")
rsFiles.Close
Set rsFiles = Nothing

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Sub SetThumbScroll()
If (ThumbRows - Fix(ThumbFrame.Height / ThumbHeight)) * ThumbHeight > 0 Then
    ThumbScroll.Max = (ThumbRows - Fix(ThumbFrame.Height / ThumbHeight)) * ThumbHeight
    ThumbScroll.Visible = True
   Else
    ThumbScroll.Max = 0
    ThumbScroll.Visible = False
End If
End Sub
Private Function SetStop(Status As Boolean)
If Status Then
    frmStopProcess.Show
    MainForm.Enabled = False
'    btn_stop.Enabled = True
'    btn_Search.Enabled = False
'    Search_in.Enabled = False
'    CDs.Enabled = False
   Else
    Unload frmStopProcess
    MainForm.Enabled = True
    MainForm.SetFocus
    'btn_stop.Enabled = False
    'btn_Search.Enabled = True
    'Search_in.Enabled = True
    'CDs.Enabled = True
End If
End Function

Private Function ShowContent(NodeIndex)
Dim i, ChildNodeID, FileID, CDID
Dim ChildNode As Node
Dim mNode As Node
Dim xSetStop As Boolean

On Error GoTo err1

'If gStillReadingFiles Then
'    xSetStop = True
'   Else
'    xSetStop = False
'End If
    
gSelectedThumbNail = 0
Set mNode = CDs.Nodes(NodeIndex)

'Add child nodes
If mNode.Parent Is Nothing And mNode.Children = 0 Then
    CDID = GetCDID(mNode.Tag)
    AddChildNodes mNode, CDID
End If


If gDirShown = mNode.FullPath Then
    Exit Function
   Else
    gDirShown = mNode.FullPath
End If

ClearThumbnails
gStillReadingFiles = True


If mNode.Children = 0 Then
    'single picture
    FileID = GetFileID(mNode.Index)
    ShowThumbnails (FileID)
    PB.Visible = False
    PB.Value = 0
   Else
    SetHourglass (True)
    PB.Visible = True
    PB.Max = mNode.Children
    
    Set ChildNode = mNode.Child
    i = 0
    Do Until ChildNode Is Nothing
        ' if another instanec of this function stop then exit this instance
        If Not gStillReadingFiles Then GoTo ExitFunction
        ChildNodeID = ChildNode.Index
        FileID = GetFileID(ChildNodeID)
        ShowThumbnails (FileID)
        PB.Value = i
        Set ChildNode = ChildNode.Next
        i = i + 1
        DoEvents
    Loop
    SetHourglass (False)
    PB.Visible = False
    ReSortThumbnails
End If

ExitFunction:
StatusBar.Panels(1).Text = "Found " & ThumbNail.Count - 1 & " Pictures"
ClearFileDetails

'gStopReadingFiles = xSetStop

gStillReadingFiles = False

Exit Function
err1:
MsgBox Error
Exit Function

End Function

Private Sub ThumbScroll_Change()
    ThumbSlide.Top = -ThumbScroll.Value
    'ThumbSlide.SetFocus
End Sub

Private Sub ThumbScroll_Scroll()
    ThumbSlide.Top = -ThumbScroll.Value
    'ThumbSlide.SetFocus
End Sub

