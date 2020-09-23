Attribute VB_Name = "modAttributes"
'Creating Character Type
Public Type Character               'Start of type
     atkstr As Integer 'attack strength (ex. 1d8 if sword)
     maxhp As Integer 'Max hitpoints
     hp As Integer 'hitpoints
     gold As Integer 'gold
     exp As Integer 'experience
     nextexp As Integer 'experience till next level
     lvl As Integer 'char lvl
     str As Integer 'strength
     strmax As Integer 'max str
     dex As Integer 'dexterity
     dexmax As Integer 'max dex
     int As Integer 'intelegence
     intmax As Integer 'max int
     sex As String 'sex
     name As String 'Char name
     race As String 'Char race
     cond As String 'Chars condition Good/Fair/Critical/Poisoned/etc..
     
End Type 'End of type
Public Hero As Character 'Assigning "Hero" to the type
Public Enemy(100) As Character 'Assigning "Enemy" to the type with possibility of 100 enemies
'End of Character/Enemy attribute Type

'Setting up world
Public Type WorldArray   ' define the contents of the World array
    TotEnemys As Integer    ' how many enemies are defined
    TotObjects As Integer      ' how many objects are defined - Not yet used
End Type
Public World As WorldArray
'End of world settings

Sub Main()
Randomize 'Start the random engine
Form1.Show 'Show the character creation form
End Sub


Sub CreateEnemy(ByVal Mon As Integer)
    'Time to create a enemy!
    ' Enemy #'s:
    '   1 = Small Snake
    '   2 = Weak Skeleton
    '   3 = Zombie
    '   4 = Spiderling
    '   5 = Bear Cub
    ' Call CreateEnemy(0) Chooses a random enemy
    ' Call CreateEnemy(#) Chooses that enemy
    '   Universe.TotEnemys to hold the number of existing Enemys
    
    World.TotEnemys = 5 '<- SET THIS TO HOW MANY ENEMIES YOU HAVE CREATED!
    
    If Mon = 0 Then   ' random object desired, pick the object
        CurMonType% = Int(Rnd * 5) + 1
    ElseIf Mon > 0 Then CurMonType% = Mon
    End If
    If CurMonType% < 1 Or CurMonType% > 5 Then ' check for invalid Enemy no
        MsgBox "CreateEnemy discovered invalid CurMonType% of" + str(CurMonType%)
        End
    End If
    

   
Select Case CurMonType%   ' select a enemy based on the value of CurMonType%
    Case 1   ' Small snake
        Enemy(CurMonNo%).name = "Small Snake"
        Enemy(CurMonNo%).strmax = Rnd * 3 + 1
        Enemy(CurMonNo%).intmax = 1
        Enemy(CurMonNo%).dexmax = Rnd * 3 + 5
        Enemy(CurMonNo%).maxhp = 1
        Enemy(CurMonNo%).exp = Rnd * 10 + 1 * Hero.lvl
        Enemy(CurMonNo%).atkstr = Rnd * Enemy(CurMonNo%).strmax
    Case 2      'Weak skeleton
        Enemy(CurMonNo%).name = "Weak Skeleton"
        Enemy(CurMonNo%).strmax = Rnd * 3 + 3
        Enemy(CurMonNo%).intmax = 1
        Enemy(CurMonNo%).dexmax = Rnd * 3 + 1
        Enemy(CurMonNo%).maxhp = 3 + Rnd * 3
        Enemy(CurMonNo%).exp = Rnd * 10 + 1 * Hero.lvl
        Enemy(CurMonNo%).atkstr = Rnd * Enemy(CurMonNo%).strmax
    Case 3      ' Zombie
        Enemy(CurMonNo%).name = "Zombie"
        Enemy(CurMonNo%).strmax = Rnd * 4 + 3
        Enemy(CurMonNo%).intmax = 1
        Enemy(CurMonNo%).dexmax = Rnd * 2 + 1
        Enemy(CurMonNo%).maxhp = 5 + Rnd * 5
        Enemy(CurMonNo%).exp = Rnd * 10 + 1 * Hero.lvl
        Enemy(CurMonNo%).atkstr = Rnd * Enemy(CurMonNo%).strmax
    Case 4      ' Spiderling
        Enemy(CurMonNo%).name = "Spiderling"
        Enemy(CurMonNo%).str = Rnd * 4 + 4
        Enemy(CurMonNo%).int = 1
        Enemy(CurMonNo%).dex = Rnd * 4 + 5
        Enemy(CurMonNo%).maxhp = 10 + Rnd * 5
        Enemy(CurMonNo%).exp = Rnd * 10 + 1 * Hero.lvl
        Enemy(CurMonNo%).atkstr = Rnd * Enemy(CurMonNo%).strmax
    Case 5      'Bear Cub
        Enemy(CurMonNo%).name = "Bear Cub"
        Enemy(CurMonNo%).strmax = Rnd * 5 + 5
        Enemy(CurMonNo%).intmax = 1
        Enemy(CurMonNo%).dexmax = Rnd * 4 + 5
        Enemy(CurMonNo%).maxhp = 10 + Rnd * 10
        Enemy(CurMonNo%).exp = Rnd * 10 + 1 * Hero.lvl
        Enemy(CurMonNo%).atkstr = Rnd * Enemy(CurMonNo%).strmax
    Case Else ' something is wrong because an invalid CurMon% was used
        MsgBox "HEY! THATS NOT A MONSTER NUBER YET!"
        End
    End Select

    ' Setting enemy levels to "Max" levels
    Enemy(CurMonNo%).str = Enemy(CurMonNo%).strmax
    Enemy(CurMonNo%).int = Enemy(CurMonNo%).intmax
    Enemy(CurMonNo%).dex = Enemy(CurMonNo%).dexmax
    Enemy(CurMonNo%).hp = Enemy(CurMonNo%).maxhp
    Combat.ProgressBar3.Max = Enemy(CurMonNo%).maxhp
    Combat.ProgressBar3.Value = Enemy(CurMonNo%).hp
    Combat.Frame4.Caption = Enemy(CurMon%).name & "'s Hitpoints"
End Sub



 
