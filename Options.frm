VERSION 5.00
Begin VB.Form Options 
   Caption         =   "RPG "
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load a character"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create a new character"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Options.Hide
Unload Options
CreateChar.Show
End Sub

Private Sub Command3_Click()
End
End Sub
