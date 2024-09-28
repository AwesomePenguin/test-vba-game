' Character Class Module
Option Explicit

' Private variables to store property values
Private pHP As Integer
Private pMP As Integer
Private pAttack As Integer
Private pDefense As Integer
Private pGold As Integer
Private pExp As Integer
Private pLevel As Integer

Private pName As String

' Property for HP
Public Property Get HP() As Integer
    HP = pHP
End Property

Public Property Let HP(Value As Integer)
    pHP = Value
End Property

' Property for MP
Public Property Get MP() As Integer
    MP = pMP
End Property

Public Property Let MP(Value As Integer)
    pMP = Value
End Property

' Property for Attack
Public Property Get Attack() As Integer
    Attack = pAttack
End Property

Public Property Let Attack(Value As Integer)
    pAttack = Value
End Property

' Property for Defense
Public Property Get Defense() As Integer
    Defense = pDefense
End Property

Public Property Let Defense(Value As Integer)
    pDefense = Value
End Property

' Property for Gold
Public Property Get Gold() As Integer
    Gold = pGold
End Property

Public Property Let Gold(Value As Integer)
    pGold = Value
End Property

' Property for Experience
Public Property Get Exp() As Integer
    Exp = pExp
End Property

Public Property Let Exp(Value As Integer)
    pExp = Value
End Property

' Property for Level
Public Property Get Level() As Integer
    Level = pLevel
End Property

Public Property Let Level(Value As Integer)
    pLevel = Value
End Property

' Property for Name
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(Value As String)
    pName = Value
End Property

' Method to take damage
Public Sub TakeDamage(ByVal Damage As Integer)
    Dim actualDamage As Integer
    actualDamage = Damage - pDefense
    If actualDamage < 0 Then actualDamage = 0
    pHP = pHP - actualDamage
    If pHP < 0 Then pHP = 0
End Sub

' Method to use magic
Public Sub UseMagic(ByVal MagicCost As Integer)
    pMP = pMP - MagicCost
    If pMP < 0 Then pMP = 0
End Sub

' Method to gain experience
Public Sub GainExp(ByVal ExpGain As Integer)
    pExp = pExp + ExpGain
    If pExp >= 100 Then
        pExp = pExp - 100
        pLevel = pLevel + 1
        pHP = pHP + 10
        pMP = pMP + 5
        pAttack = pAttack + 2
        pDefense = pDefense + 1
    End If
End Sub