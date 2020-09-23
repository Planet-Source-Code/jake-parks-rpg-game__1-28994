Attribute VB_Name = "modGoldandExp"

Sub exper()

'Experience Calculations
 'Allows the moster to give a random ammount of exp
    MsgBox "You got " & Enemy(CurMonNo%).exp & " experience points!" 'announce exp
    Hero.exp = Hero.exp + Enemy(CurMonNo%).exp 'add experience points
    
'A simple way to go up in levels, will probably be modified later
'once I figure out how all the numbers should work.
If Hero.exp >= Hero.nextexp Then 'Determining if char has enough exp to go up a level
    Hero.lvl = Hero.lvl + 1 'Adding a level
    Hero.nextexp = Hero.nextexp * Hero.lvl 'Setting a new exp goal for next level
    MsgBox "You went up a level! You are now Level " & Hero.lvl & "!", , "Congratulations!" 'Telling user they went up a level
    Hero.maxhp = Hero.maxhp + Hero.lvl * 10 'Giving hp bonus
    MsgBox "You gained " & Hero.lvl * 10 & " Hitpoints!" 'Telling user they gaind Hp's
    Combat.ProgressBar2.Max = Hero.nextexp 'Setting new max for bar
    Combat.ProgressBar2.Value = Hero.exp 'Update exp progress bar
    Combat.ExpLab.Caption = Hero.exp & "/" & Hero.nextexp 'Displys hero's exp
    Combat.LvlLab.Caption = Hero.lvl 'Updating Level caption
    Hero.hp = Hero.maxhp 'Refilling hitpoints
    Combat.lblHp.Caption = Hero.hp & "/" & Hero.maxhp 'Updating hp caption
    Combat.ProgressBar1.Max = Hero.maxhp 'Updating hp progressbar max
    Combat.ProgressBar1.Value = Hero.hp 'Updating hp progressbar value
    
Else
    Combat.ProgressBar2.Max = Hero.nextexp 'Re-setting progressbar max
    Combat.ProgressBar2.Value = Hero.exp 'Update exp progress bar
    Combat.ExpLab.Caption = Hero.exp & "/" & Hero.nextexp 'Displys hero's exp
End If
End Sub

Sub gold()
'Gold Calculations
Enemy(CurMonNo%).gold = Val(Rnd * 5 + Hero.lvl) 'Random amount of gold pieces
    MsgBox "You got " & Enemy(CurMonNo%).gold & " Gold Pieces!", , "Treasure" 'Tells how much trease the hero got
    Hero.gold = Hero.gold + Enemy(CurMonNo%).gold 'Adds new treasure to the existing
    Combat.GoldLab.Caption = Hero.gold 'Displays new amount in the gold caption

End Sub

Sub Condition()
'This does not work yet for some reason.
'It is supposed to update the character's condition
'based on the percentage of hitpoints they have left.
'Maybe in a later version.
If (Hero.maxhp * 0.75) < Hero.hp Then
    Hero.cond = "Good"
    Combat.condlbl.ForeColor = vbGreen
    Combat.condlbl.Caption = Hero.cond
ElseIf (Hero.maxhp * 0.75) > Hero.hp Then
    Hero.cond = "Fair"
    Combat.condlbl.ForeColor = vbBlue
    Combat.condlbl.Caption = Hero.cond
ElseIf (Hero.maxhp * 0.25) > Hero.hp Then
    Hero.cond = "Critical"
    condlbl.ForeColor = vbYellow
    condlbl.Caption = Hero.cond
ElseIf Hero.hp = 0 Then
    Hero.cond = "Dead"
    condlbl.ForeColor = vbRed
    condlbl.Caption = Hero.cond
End If
End Sub
