VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membuka /Menutup CD-Drive"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "TUTUP"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUKA"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  retvalue = mciSendString("set Cdaudio door open", returnstring, 127, 0)
End Sub

Private Sub Command2_Click()
  retvalue = mciSendString("set Cdaudio door closed", returnstring, 127, 0)
End Sub



