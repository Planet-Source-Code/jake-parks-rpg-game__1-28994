VERSION 5.00
Begin VB.Form CreateChar 
   Caption         =   "Form2"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   ScaleHeight     =   4695
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Choose a Race"
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
      Begin VB.OptionButton GiantOption 
         Caption         =   "Giant"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton TrollOption 
         Caption         =   "Troll"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton DwarfOption 
         Caption         =   "Dwarf"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton ElfOption 
         Caption         =   "Elf"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton HumanOption 
         Caption         =   "Human"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose a name for your character"
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   240
      Width           =   2655
      Begin VB.TextBox NameText 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Your Sex"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1575
      Begin VB.OptionButton FemaleOption 
         Caption         =   "Female"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton MaleOption 
         Caption         =   "Male"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Accept"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Roll Again"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   $"CreateChar.frx":0000
      Height          =   1695
      Left            =   5160
      TabIndex        =   19
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "INT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DEX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "STR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "CreateChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Rolling new stats
Hero.str = Val(Rnd * 10 + 10)
Hero.dex = Val(Rnd * 10 + 10)
Hero.int = Val(Rnd * 10 + 10)
'Displaying new stats
Label1.Caption = Hero.str
Label2.Caption = Hero.dex
Label3.Caption = Hero.int
End Sub

Private Sub Command2_Click()
'Checking Character sex choice
If MaleOption = True Then
    Hero.sex = "Male"
ElseIf FemaleOption = True Then
    Hero.sex = "Female" 'Setting Sex
Else
    MsgBox "Please choose a sex before you continue", , "Missing Entry"
End If

'Checking Character name
If NameText.Text = "" Then
    MsgBox "Please enter a name before you continue.", , "Missing Entry"
End If
    Hero.name = NameText.Text 'Setting Name

'Checking Character race
If HumanOption = True Then
    Hero.race = "Human"
ElseIf ElfOption = True Then
    Hero.race = "Elf"
ElseIf TrollOption = True Then
    Hero.race = "Troll"
ElseIf GiantOption = True Then
    Hero.race = "Giant"
ElseIf DwarfOption = True Then
    Hero.race = "Dwarf"
Else
    MsgBox "Please choose a race before you continue", , "Missing Entry"
End If

CreateChar.Hide 'Removing Creation Screen
Combat.Show 'Showing the Main

End Sub

Private Sub Form_Load()
'Setting initial Character stats
Hero.str = Val(Rnd * 10 + 10)
Hero.dex = Val(Rnd * 10 + 10)
Hero.int = Val(Rnd * 10 + 10)
'Displaying initial Character stats
Label1.Caption = Hero.str
Label2.Caption = Hero.dex
Label3.Caption = Hero.int
'Setting new hero's level to 1
Hero.lvl = 1
'Setting Experience Points to 0 for new char
Hero.exp = 0
'Setting New Chars' Experience Points till level 2 to 50
Hero.nextexp = 50

End Sub

