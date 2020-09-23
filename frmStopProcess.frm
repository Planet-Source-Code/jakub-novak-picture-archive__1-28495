VERSION 5.00
Begin VB.Form frmStopProcess 
   Caption         =   "Cancel Proccess"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   Icon            =   "frmStopProcess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   2775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btn_stop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmStopProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_stop_Click()
gCancel = True
End Sub
