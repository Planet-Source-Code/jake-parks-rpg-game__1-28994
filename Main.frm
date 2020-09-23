VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Combat 
   BackColor       =   &H8000000A&
   Caption         =   "Rpg Combat "
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   3120
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   49
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame Frame8 
      Caption         =   "Testing Buttons"
      Height          =   1455
      Left            =   6360
      TabIndex        =   46
      Top             =   4800
      Width           =   2535
      Begin VB.CommandButton Command2 
         Caption         =   "Find something to fight"
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "HEAL ME !!!"
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Experiece Points"
      Height          =   1455
      Left            =   120
      TabIndex        =   41
      Top             =   360
      Width           =   2415
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label LvlLab 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Level:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.Label ExpLab 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Movement"
      Height          =   2055
      Left            =   3000
      TabIndex        =   31
      Top             =   3480
      Width           =   2775
      Begin VB.CommandButton NorthButton 
         Caption         =   "N"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CampButton 
         Caption         =   "Camp"
         Enabled         =   0   'False
         Height          =   495
         Left            =   960
         TabIndex        =   39
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton SouthButton 
         Caption         =   "S"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton EastButton 
         Caption         =   "E"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   37
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton WestButton 
         Caption         =   "W"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton NEButton 
         Caption         =   "NE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton NWButton 
         Caption         =   "NW"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton SWButton 
         Caption         =   "SW"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton SEButton 
         Caption         =   "SE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1920
         TabIndex        =   32
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Character Stats"
      Height          =   3495
      Left            =   6360
      TabIndex        =   16
      Top             =   240
      Width           =   2535
      Begin VB.Label condlbl 
         Height          =   255
         Left            =   960
         TabIndex        =   45
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label LAbel3 
         Caption         =   "Condition:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label GoldLab 
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Gold:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Sex:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Race:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "STR:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "DEX:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "INT:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label NameLab 
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label SexLab 
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label RaceLab 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label StrLab 
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label DexLab 
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label IntLab 
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Combat Options"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   2415
      Begin VB.CommandButton FleeButton 
         Caption         =   "Flee"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton CastButton 
         Caption         =   "Cast Spell"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton atkButton 
         Caption         =   "Attack"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Enemy Hit Points"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblHpE 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton BackpackButton 
      Caption         =   "Backpack"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your Hit Points"
      Height          =   855
      Left            =   6360
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblHp 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer EnemyAttack 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   6000
   End
   Begin VB.TextBox status 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton QuitButton 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Your  Weapon"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2415
      Begin VB.OptionButton TestOption 
         Caption         =   "Test Wep - Damage =50"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   1020
         Width           =   2055
      End
      Begin VB.OptionButton SwordOption 
         Caption         =   "Sword - Damage = 1d8"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton FistOption 
         Caption         =   "Fists - Damage = 1d4"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Graphics will go here ^"
      Height          =   255
      Left            =   3120
      TabIndex        =   50
      Top             =   2760
      Width           =   2535
   End
End
Attribute VB_Name = "Combat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Author: Jake Parks
'Project Name: RPG Game
'Copyright(C) 2001 by Jake Parks
'This is a simple program. I have only been programming
'about a two weeks or so, and I wanted to test out my skills. I'm
'sure there are better ways to go about this, any comments
'will be appreciated :)

Private Sub atkButton_Click()
'Setting Players attack strength according to the weapon they chose.
If FistOption = True Then
    Hero.atkstr = Val(Rnd * 4) 'Fist
ElseIf SwordOption = True Then
    Hero.atkstr = Val(Rnd * 8) 'Sword
ElseIf TestOption = True Then
    Hero.atkstr = 50 'Test
End If
'End of setting players attack strength

'Attacking the Enemy(CurMonNo%)
Enemy(CurMonNo%).hp = Enemy(CurMonNo%).hp - Hero.atkstr 'the actual hit
lblHpE.Caption = Enemy(CurMonNo%).hp & "/" & Enemy(CurMonNo%).maxhp
If Hero.atkstr > 0 Then status.Text = "You hit for " & Hero.atkstr & " points!"
If Hero.atkstr = 0 Then status.Text = "You Missed!" 'If hero misses
If Enemy(CurMonNo%).hp < 0.25 * Enemy(CurMonNo%).maxhp Then lblHpE.ForeColor = vbRed 'Diplay turns red if hp's fall below 25 % of total
If Enemy(CurMonNo%).hp <= 0 Then Enemy(CurMonNo%).hp = 0 'Enemy(CurMonNo%) hp's are set to 0 if they fall below 0
ProgressBar3.Value = Enemy(CurMonNo%).hp
If Enemy(CurMonNo%).hp <= 0 Then
    lblHpE.Caption = Enemy(CurMonNo%).hp & "/" & Enemy(CurMonNo%).maxhp 'updates Enemy(CurMonNo%) hp caption
    MsgBox "You Won the Fight!", , "Win" 'Message if fight is won
    Call gold
    Call exper
    atkButton.Enabled = False 'Disable the attack button till next fight
    FleeButton.Enabled = False 'Disable the Flee button till next fight
    Frame4.Enabled = False 'Greying out Enemy(CurMonNo%) hp frame till next fight
    lblHpE.Enabled = False 'Greying out Enemy(CurMonNo%) hp frame till next fight
ElseIf Enemy(CurMonNo%).hp > 0 Then
    atkButton.Enabled = False 'Disable the attack button till next turn
    FleeButton.Enabled = False 'Disable the Flee button till next turn
    EnemyAttack.Enabled = True 'If enemies hp are > 0 attack again
End If
'End of Attacking the Enemy(CurMonNo%)
End Sub

Private Sub Command1_Click()
'Complete Heal! -This might be the formula for a complete
'healing potion in the future-
Hero.hp = Hero.maxhp 'Set hp back to max
lblHp.ForeColor = vbBlack
lblHp.Caption = Hero.hp & "/" & Hero.maxhp 'Disply in caption
ProgressBar1.Value = Hero.hp 'Reset progress bar
End Sub

Private Sub Command2_Click()
'Reset The fight - This is a temporary button for testing -
Call CreateEnemy(0)
Enemy(CurMonNo%).atkstr = Enemy(CurMonNo%).strmax
ProgressBar3.Enabled = True 're-enable progress bar
ProgressBar3.Value = Enemy(CurMonNo%).hp 'Reset progress bar
lblHpE.Enabled = True 'Enable Enemy hp caption
lblHpE.ForeColor = vbBlack 'reset caption color to black
Frame4.Enabled = True 're-enable frame around Enemy(CurMonNo%) hp
lblHpE.Caption = Enemy(CurMonNo%).hp & "/" & Enemy(CurMonNo%).maxhp 'Reset Enemy hp caption
atkButton.Enabled = True 're-enable the attack button
FleeButton.Enabled = True 're-enable the flee button
'An Enemy will attack first this time :)
MsgBox "You are being attacked by a " & Enemy(CurMonNo%).name & "!!!", , "A Friendly Reminder" 'Tell the player he is being attacked
EnemyAttack.Enabled = True 'Start the fight!
End Sub



Private Sub EnemyAttack_Timer()


'Enemy(CurMonNo%) attack routine
Hero.hp = Hero.hp - Enemy(CurMonNo%).atkstr 'The actual hit
lblHp.Caption = Hero.hp & "/" & Hero.maxhp '
    If Enemy(CurMonNo%).atkstr > 0 Then status.Text = "You get hit for " & Enemy(CurMonNo%).atkstr & " points!"
    If Enemy(CurMonNo%).atkstr = 0 Then status.Text = "The " & Enemy(CurMonNo%).name & "Missed!"
    If Hero.hp <= 0.25 * Hero.maxhp Then lblHp.ForeColor = vbRed 'Caption turns red if hp fall below 25% of total
    If Hero.hp <= 0 Then Hero.hp = 0
    ProgressBar1.Value = Hero.hp
    lblHp.Caption = Hero.hp & "/" & Hero.maxhp
    If Hero.hp <= 0 Then
        MsgBox "You have died...", , "Lost fight"
        atkButton.Enabled = False
        FleeButton.Enabled = False
        EnemyAttack.Enabled = False
        Frame4.Enabled = False
        lblHpE.Enabled = False
Else
EnemyAttack.Enabled = False 'Ending Enemy(CurMonNo%)'s turn
atkButton.Enabled = True 'Starting Hero's turn
FleeButton.Enabled = True 'Allowing Hero to fee during his turn
End If

End Sub

Private Sub FleeButton_Click()
'Give the Char a 50/50 Chance to get away.
Dim random As Integer 'creating a variable called "random"
random = Val(Rnd * 1) 'give a 50/50 chance, random will = 0 or 1
If random = 1 Then 'Char does not get away
    MsgBox "You could not get away!", , "Flee Attempt"
    FleeButton.Enabled = False 'Char looses turn
    atkButton.Enabled = False 'Char looses turn
    EnemyAttack.Enabled = True 'Enemy(CurMonNo%) attacks when trying to flee
ElseIf random = 0 Then 'Char gets away
    MsgBox "You got away... You fight like a....", , "Flee Attempt" 'The message box text
    lblHpE.ForeColor = vbBlack '"  " 'resetting Enemy(CurMonNo%) caption to black
    lblHpE.Caption = Enemy(CurMonNo%).hp & "/" & Enemy(CurMonNo%).maxhp 'Diplays monster display status
    status.Text = "" 'Clearing the status box
    atkButton.Enabled = False 'Disable attacks
    FleeButton.Enabled = False 'Disable flee
    lblHpE.Enabled = False 'disable
    Frame4.Enabled = False 'disabe
    ProgressBar3.Enabled = False 'disable
    
End If
End Sub

Private Sub Form_Load()
'Setting Initial Values
Hero.maxhp = 20 'Setting New Hero's maxhp to 10
Hero.hp = Hero.maxhp 'Setting hp to the max
lblHp.Caption = Hero.hp & "/" & Hero.maxhp 'Diplays hp status
ExpLab.Caption = Hero.exp & "/" & Hero.nextexp 'Displys hero's exp

'Diplaying Character Stats
NameLab.Caption = Hero.name 'name
SexLab.Caption = Hero.sex 'sex
RaceLab.Caption = Hero.race 'race
StrLab.Caption = Hero.str 'strength
DexLab.Caption = Hero.dex 'dexterity
IntLab.Caption = Hero.int 'inteligence
GoldLab.Caption = Hero.gold 'gold
LvlLab.Caption = Hero.lvl 'level

'Setting progress bar to hit point values
ProgressBar1.Max = Hero.maxhp
ProgressBar1.Value = Hero.hp

'Setting progress bar to exp values
ProgressBar2.Max = Hero.nextexp
ProgressBar2.Value = Hero.exp

'Disabling Enemy Hp bar and captions until a fight starts
Frame4.Enabled = False
lblHpE.Enabled = False
ProgressBar3.Enabled = False

atkButton.Enabled = False
FleeButton.Enabled = fase

End Sub
Private Sub QuitButton_Click()
'Shutdown the program!
End
End Sub




