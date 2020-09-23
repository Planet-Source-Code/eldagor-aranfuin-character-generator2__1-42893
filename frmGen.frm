VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fantasy Character Generator"
   ClientHeight    =   4515
   ClientLeft      =   4095
   ClientTop       =   3315
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   4515
   ScaleWidth      =   8175
   Begin VB.Frame Frame8 
      Caption         =   "Weapons"
      Height          =   1335
      Left            =   4440
      TabIndex        =   32
      Top             =   360
      Width           =   1935
      Begin VB.Image imgSsword2 
         Height          =   555
         Left            =   1320
         Picture         =   "frmGen.frx":0000
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgSsword 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":0FDE
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label12 
         Caption         =   "Sec."
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   960
         Width           =   375
      End
      Begin VB.Image imgmornstar2 
         Height          =   555
         Left            =   1320
         Picture         =   "frmGen.frx":1FBC
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgLsword2 
         Height          =   555
         Left            =   1320
         Picture         =   "frmGen.frx":2F9A
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgmace2 
         Height          =   555
         Left            =   1320
         Picture         =   "frmGen.frx":3F78
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgbow2 
         Height          =   555
         Left            =   1320
         Picture         =   "frmGen.frx":4F56
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgdagger 
         Height          =   555
         Left            =   1320
         Picture         =   "frmGen.frx":5F34
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label11 
         Caption         =   "Prim."
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   615
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Image imgbow 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":6F12
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgstaff 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":7EF0
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgmornstar 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":8ECE
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgmace 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":9EAC
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgaxe 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":AE8A
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgLsword 
         Height          =   555
         Left            =   120
         Picture         =   "frmGen.frx":BE68
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Legs"
      Height          =   1215
      Left            =   6480
      TabIndex        =   31
      Top             =   2880
      Width           =   1095
      Begin VB.Image imgLlegs 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":CE46
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgPlegs 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":DE24
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgClegs 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":EE02
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imglegs 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":FDE0
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Body"
      Height          =   1215
      Left            =   6960
      TabIndex        =   30
      Top             =   1680
      Width           =   1095
      Begin VB.Image imgtunic 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":10DBE
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgParmor 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":11D9C
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgCarmor 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":12D7A
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgLarmor 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":13D58
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Hands"
      Height          =   1215
      Left            =   5880
      TabIndex        =   29
      Top             =   1680
      Width           =   1095
      Begin VB.Image imgPglove 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":14D36
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgCgloves 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":15D14
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgLglove 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":16CF2
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Head"
      Height          =   1215
      Left            =   6480
      TabIndex        =   28
      Top             =   480
      Width           =   1095
      Begin VB.Image imghelm 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":17CD0
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgchaincap 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":18CAE
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Image imgLcap 
         Height          =   555
         Left            =   240
         Picture         =   "frmGen.frx":19C8C
         Top             =   360
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Program"
      Height          =   255
      Left            =   4440
      TabIndex        =   27
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Generate"
      Height          =   615
      Left            =   2880
      TabIndex        =   25
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   2760
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   2640
      Width           =   3015
      Begin VB.TextBox txtArmor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtWeap2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtWeap1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Character Armor:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Secondary                         Weapon:"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Primary Weapon:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
      Begin VB.OptionButton optOld 
         Caption         =   "Older"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton optYoung 
         Caption         =   "Young"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtEyes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtStyle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtLength 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtHair 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   2280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   2280
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label7 
         Caption         =   "Eye Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Hair Style:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Hair Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hair Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      Begin VB.TextBox txtClass 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtRace 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   1920
         X2              =   1920
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Class:"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Race:"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Character Name:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Clear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click()
If txtName.Text = "" Then
    MsgBox "Everything is already cleared.", vbExclamation, "Already clear"
Else
    Clear_Click
End If
End Sub


Private Sub Clear_Click()
If txtName.Text = "" Then
    MsgBox "Everything is already cleared.", vbExclamation, "Already clear"
Else
    Dim Response As Integer
    intResponse = MsgBox("Are you sure you want to clear all the information?", vbYesNo + vbQuestion, "Clear")
    If intResponse = vbYes Then
        txtName.Text = ""
        txtRace.Text = ""
        txtClass.Text = ""
        txtHair.Text = ""
        txtStyle.Text = ""
        txtLength.Text = ""
        txtEyes.Text = ""
        txtWeap1.Text = ""
        txtWeap2.Text = ""
        txtArmor.Text = ""
        optYoung.Value = False
        optOld.Value = False
        imghelm.Visible = False
        imgPglove.Visible = False
        imgPlegs.Visible = False
        imgParmor.Visible = False
        '----
        imgLcap.Visible = False
        imgLglove.Visible = False
        imgLlegs.Visible = False
        imgLarmor.Visible = False
        '----
        imgchaincap.Visible = False
        imgCgloves.Visible = False
        imgClegs.Visible = False
        imgCarmor.Visible = False
        imgtunic.Visible = False
        imglegs.Visible = False
        '----
        imgstaff.Visible = False
        imgmornstar.Visible = False
        imgbow.Visible = False
        imgLsword.Visible = False
        imgSsword.Visible = False
        imgmace.Visible = False
        imgaxe.Visible = False
        '----
        imgLsword2.Visible = False
        imgSsword2.Visible = False
        imgmace2.Visible = False
        imgmornstar2.Visible = False
        imgbow2.Visible = False
        imgdagger.Visible = False
        'Ouch, hand cramps :(
    End If
End If
End Sub

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub cmdGo_Click()
'Generates Random Race
Randomize Timer
Let r = Int(Rnd * 9) + 1
If r = 1 Then
    txtRace.Text = "Human"
ElseIf r = 2 Then
    txtRace.Text = "Elf"
ElseIf r = 3 Then
    txtRace.Text = "Half-Elf"
ElseIf r = 4 Then
    txtRace.Text = "Dwarf"
ElseIf r = 5 Then
    txtRace.Text = "Gnome"
ElseIf r = 6 Then
    txtRace.Text = "Halfling"
ElseIf r = 7 Then
    txtRace.Text = "Orc"
ElseIf r = 8 Then
    txtRace.Text = "Half-Orc"
ElseIf r = 9 Then
    txtRace.Text = "Dark Elf"
    txtEyes.Text = "Red"
End If

'Generates Random Class
Randomize Timer
Let q = Int(Rnd * 7) + 1
If q = 1 Then
    txtClass.Text = "Warrior/Fighter"
ElseIf q = 2 Then
    txtClass.Text = "Theif"
ElseIf q = 3 Then
    txtClass.Text = "Mage"
ElseIf q = 4 Then
    txtClass.Text = "Ranger"
ElseIf q = 5 Then
    txtClass.Text = "Assassin"
ElseIf q = 6 Then
    txtClass.Text = "Cleric"
ElseIf q = 7 Then
    txtClass.Text = "Common Citizen"
End If

'Generates Eye Color
If r = 9 Then
    txtEyes.Text = "Red"
Else
    Randomize Timer
    Let u = Int(Rnd * 4) + 1
    If u = 1 Then
        txtEyes.Text = "Blue"
    ElseIf u = 2 Then
        txtEyes.Text = "Brown"
    ElseIf u = 3 Then
        txtEyes.Text = "Hazel"
    ElseIf u = 4 Then
        txtEyes.Text = "Green"
    End If
End If

'Hair stuffs
If r = 6 Then
    txtStyle.Text = "Curly"
    txtLength.Text = "Short"
Else
    Randomize Timer
    Let h = Int(Rnd * 5) + 1
    If h = 1 Then
        txtLength.Text = "Bald"
    ElseIf h = 2 Then
        txtLength.Text = "Short"
    ElseIf h = 3 Then
        txtLength.Text = "Long"
    ElseIf h = 4 Then
        txtLength.Text = "Short"
    ElseIf h = 5 Then
        txtLength.Text = "Long"
    End If
End If
Randomize Timer
Let p = Int(Rnd * 5) + 1
If h = 1 Then
    txtHair.Text = "---"
ElseIf r = 9 Then
    txtHair.Text = "White"
Else
    If p = 1 Then
        txtHair.Text = "Brown"
    ElseIf p = 2 Then
        txtHair.Text = "Blonde"
    ElseIf p = 3 Then
        txtHair.Text = "Black"
    ElseIf p = 4 Then
        txtHair.Text = "Auburn"
    ElseIf p = 5 Then
        txtHair.Text = "White"
    End If
End If
Randomize Timer
Let s = Int(Rnd * 2) + 1
If h = 1 Then
    txtStyle.Text = "---"
ElseIf r = 6 Then
    txtStyle.Text = "Curly"
Else
    If s = 1 Then
        txtStyle.Text = "Straight"
    ElseIf s = 2 Then
        txtStyle.Text = "Wavey"
    ElseIf s = 3 Then
        txtStyle.Text = "Curly"
    End If
End If



'Primary Weapons
If q = 3 Then
    txtWeap1.Text = "Staff"
ElseIf q = 6 Then
    txtWeap1.Text = "Clerical Powers"
ElseIf r = 4 Then
    If q <> 3 And q <> 6 Then
        txtWeap1.Text = "Axe/Battle Axe"
    End If
Else
    Randomize Timer
    Let w1 = Int(Rnd * 6) + 1
    If w1 = 1 Then
        txtWeap1.Text = "Long Sword"
    ElseIf w1 = 2 Then
        txtWeap1.Text = "Short Sword"
    ElseIf w1 = 3 Then
        txtWeap1.Text = "Mace"
    ElseIf w1 = 4 Then
        txtWeap1.Text = "Axe/Battle Axe"
    ElseIf w1 = 5 Then
        txtWeap1.Text = "Bow"
    ElseIf w1 = 6 Then
        txtWeap1.Text = "Morning Star"
    End If
End If

'Secondary Weapons
If q = 3 Or q = 6 Then
    Randomize Timer
    Let w2 = Int(Rnd * 5) + 1
    If w2 = 1 Then
        txtWeap2.Text = "Long Sword"
    ElseIf w2 = 2 Then
        txtWeap2.Text = "Short Sword"
    ElseIf w2 = 3 Then
        txtWeap2.Text = "Mace"
    ElseIf w2 = 4 Then
        txtWeap2.Text = "Morning Star"
    ElseIf w2 = 5 Then
        txtWeap2.Text = "Bow"
    End If
Else
    Randomize Timer
    Let w2 = Int(Rnd * 3) + 1
    If w2 = 1 Then
        txtWeap2.Text = "Short Sword"
    ElseIf w2 = 2 Then
        txtWeap2.Text = "Dagger"
    ElseIf w2 = 3 Then
        txtWeap2.Text = "Bow"
    End If
End If

'Armor
If q = 1 Then
    txtArmor.Text = "Full Armor"
ElseIf q = 2 Then
    Randomize Timer
    Let t = Int(Rnd * 2) + 1
    If t = 1 Then
        txtArmor.Text = "Leather"
    ElseIf t = 2 Then
        txtArmor.Text = "Chainmail"
    End If
ElseIf q = 3 Then
    txtArmor.Text = "None"
ElseIf q = 4 Then
    Randomize Timer
    Let t = Int(Rnd * 2) + 1
    If t = 1 Then
        txtArmor.Text = "Leather"
    ElseIf t = 2 Then
        txtArmor.Text = "Chainmail"
    End If
ElseIf q = 5 Then
    Randomize Timer
    Let t = Int(Rnd * 2) + 1
    If t = 1 Then
        txtArmor.Text = "Leather"
    ElseIf t = 2 Then
        txtArmor.Text = "Chainmail"
    End If
ElseIf q = 6 Then
    txtArmor.Text = "Leather"
ElseIf q = 7 Then
    txtArmor.Text = "Leather"
End If

'Ages
Randomize Timer
Let age = Int(Rnd * 3) + 1
If age = 1 Then
    optOld.Value = True
ElseIf age = 2 Or age = 3 Then
    optYoung.Value = True
End If

'Weight
'------------
'If r = 1 Then
    'Randomize Timer
    'Let m = Int(Rnd * 235) + 150
    'txtWeight.Text = m
'ElseIf r = 2 Then
    'Randomize Timer
    'Let n = Int(Rnd * 150) + 100
    'txtWeight.Text = n
'ElseIf r = 3 Then
    'Randomize Timer
    'Let j = Int(Rnd * 200) + 125
    'txtWeight.Text = j
'ElseIf r = 4 Then
    'Randomize Timer
    'Let k = Int(Rnd * 200) + 150
    'txtWeight.Text = k
'ElseIf r = 5 Then
    'Randomize Timer
    'Let l = Int(Rnd * 175) + 125
    'txtWeight.Text = l
'ElseIf r = 6 Then
    'Randomize Timer
    'Let i = Int(Rnd * 190) + 115
    'txtWeight.Text = i
'ElseIf r = 7 Then
    'Randomize Timer
    'Let v = Int(Rnd * 250) + 150
    'txtWeight.Text = v
'ElseIf r = 8 Then
    'Randomize Timer
    'Let u = Int(Rnd * 225) + 150
    'txtWeight.Text = u
'End If


'Names!
Dim x As Integer
Dim Nm(100)    'Human names
Dim Enms(100)  'Elven names
Dim Onms(100)  'Orc Names
Dim Hnms(60)   'Halfling Names
Dim Dnms(60)   'Dwarf Names
Dim Gnms(50)   'Gnome Names
Dim x1 As Integer
Dim x2 As Integer
Dim x3 As Integer
Dim x4 As Integer
Dim x5 As Integer
Dim x6 As Integer

'100 Human Names
Nm(1) = "Fohamond"
Nm(2) = "Acad"
Nm(3) = "Gwaunad"
Nm(4) = "Trirawyth"
Nm(5) = "Asic"
Nm(6) = "Ybieth"
Nm(7) = "Criwyn"
Nm(8) = "Adaegord"
Nm(9) = "Asorewyn"
Nm(10) = "Afeladric"
Nm(11) = "Chaoldan"
Nm(12) = "Agreaband"
Nm(13) = "Ocoemond"
Nm(14) = "Aciekith"
Nm(15) = "Mirireder"
Nm(16) = "Frilinn"
Nm(17) = "Roha"
Nm(18) = "Traon"
Nm(19) = "Erendawin"
Nm(20) = "Pilawyn"
Nm(21) = "Gleakor"
Nm(22) = "Arorenyth"
Nm(23) = "Saonydd"
Nm(24) = "Acilabard"
Nm(25) = "Sevaoron"
Nm(26) = "Agrardon"
Nm(27) = "Ererrali"
Nm(28) = "GedrGran"
Nm(29) = "Ibied"
Nm(30) = "Elei"
Nm(31) = "Galenyth"
Nm(32) = "Thep"
Nm(33) = "Seveap"
Nm(34) = "Eowerajar"
Nm(35) = "Sigof"
Nm(36) = "Olaymar"
Nm(37) = "Werig"
Nm(38) = "Jirasean"
Nm(39) = "Grirebard"
Nm(40) = "Adralitlan"
Nm(41) = "Wicare"
Nm(42) = "Verird"
Nm(43) = "Etoba"
Nm(44) = "Sevaymas"
Nm(45) = "Zamos"
Nm(46) = "Rhalibaen"
Nm(47) = "Raejar"
Nm(48) = "Eowoihar"
Nm(49) = "Raywyr"
Nm(50) = "Grenymar"
Nm(51) = "Thelajan"
Nm(52) = "Unaekon"
Nm(53) = "Thoewan"
Nm(54) = "Alerrav"
Nm(55) = "Ereigan"
Nm(56) = "Nauk"
Nm(57) = "Pigored"
Nm(58) = "Adaobard"
Nm(59) = "Elyn"
Nm(60) = "Taeth"
Nm(61) = "Alore"
Nm(62) = "Erardoldan"
Nm(63) = "Zievudd"
Nm(64) = "Triravudd"
Nm(65) = "Gledrikath"
Nm(66) = "Agrienwan"
Nm(67) = "Neaba"
Nm(68) = "Rhaleband"
Nm(69) = "Trigowyn"
Nm(70) = "Ethe"
Nm(71) = "Teawin"
Nm(72) = "Chaolin"
Nm(73) = "Praremos"
Nm(74) = "Nydelanyth"
Nm(75) = "Koegord"
Nm(76) = "Adroer"
Nm(77) = "Gloe"
Nm(78) = "Riranydd"
Nm(79) = "Trerrav"
Nm(80) = "Cromader"
Nm(81) = "Cedrich"
Nm(82) = "Eliab"
Nm(83) = "Piehan"
Nm(84) = "Drabaen"
Nm(85) = "Abowyr"
Nm(86) = "Larea"
Nm(87) = "Cay"
Nm(88) = "Prera"
Nm(89) = "Ocybaen"
Nm(90) = "Eowareld"
Nm(91) = "Acywyr"
Nm(92) = "Etheard"
Nm(93) = "Umoreg"
Nm(94) = "Oloider"
Nm(95) = "Cubard"
Nm(96) = "Grohadan"
Nm(97) = "Vutram"
Nm(98) = "Lotheawin"
Nm(99) = "Telalin"
Nm(100) = "Aseriric"

'100 Elven Names
Enms(1) = "Isithrandir"
Enms(2) = "Fil-Gar"
Enms(3) = "Belilmariand"
Enms(4) = "Elrér"
Enms(5) = "Galénduil"
Enms(6) = "Urilman"
Enms(7) = "Niombor"
Enms(8) = "Filmalas"
Enms(9) = "Nithralith"
Enms(10) = "Méng"
Enms(11) = "Legarandir"
Enms(12) = "Nadriemir"
Enms(13) = "Elrur"
Enms(14) = "Siondir"
Enms(15) = "Niowyn"
Enms(16) = "Delarandir"
Enms(17) = "Delithrawyn"
Enms(18) = "Delynd"
Enms(19) = "Garand"
Enms(20) = "Glithrambor"
Enms(21) = "Céndel"
Enms(22) = "Thruwyn"
Enms(23) = "Hadrierion"
Enms(24) = "Lómiomir"
Enms(25) = "Cál"
Enms(26) = "Glendil"
Enms(27) = "Elorfin"
Enms(28) = "Elvándel"
Enms(29) = "Belawyn"
Enms(30) = "Rithrandel"
Enms(31) = "Pan"
Enms(32) = "Galorfirion"
Enms(33) = "Cadrieldor"
Enms(34) = "Rer"
Enms(35) = "Sil-Gandil"
Enms(36) = "Pilas"
Enms(37) = "Elilmar"
Enms(38) = "Noril"
Enms(39) = "Elrémir"
Enms(40) = "Calaldur"
Enms(41) = "Unondel"
Enms(42) = "Celáldur"
Enms(43) = "Cithralas"
Enms(44) = "Urioldur"
Enms(45) = "Eliondel"
Enms(46) = "Thrán"
Enms(47) = "Réndir"
Enms(48) = "Amyril"
Enms(49) = "Unarawyn"
Enms(50) = "Tadriendil"
Enms(51) = "Eowul"
Enms(52) = "Vil-Gand"
Enms(53) = "Eowurion"
Enms(54) = "Celithrandil"
Enms(55) = "Faralas"
Enms(56) = "Tinebririon"
Enms(57) = "Hiondir"
Enms(58) = "Torfir"
Enms(59) = "Mól"
Enms(60) = "Cándel"
Enms(61) = "Ril-Galdur"
Enms(62) = "Miriand"
Enms(63) = "Legorfind"
Enms(64) = "Lómómir"
Enms(65) = "Celadriemir"
Enms(66) = "Calilmal"
Enms(67) = "Amadrieng"
Enms(68) = "Thraldor"
Enms(69) = "Mil-Ganduil"
Enms(70) = "Pilmaldur"
Enms(71) = "Sebrind"
Enms(72) = "Legál"
Enms(73) = "Nilmaldur"
Enms(74) = "Padriel"
Enms(75) = "Amálith"
Enms(76) = "Delion"
Enms(77) = "Delamir"
Enms(78) = "Cun"
Enms(79) = "Glalas"
Enms(80) = "Glindil"
Enms(81) = "Legarang"
Enms(82) = "Elil-Gan"
Enms(83) = "Ision"
Enms(84) = "Hebrildur"
Enms(85) = "Pilmamir"
Enms(86) = "Miondil"
Enms(87) = "Málas"
Enms(88) = "Thrung"
Enms(89) = "Vyng"
Enms(90) = "Lómebrildor"
Enms(91) = "Glur"
Enms(92) = "Mundir"
Enms(93) = "Tior"
Enms(94) = "Amorfindel"
Enms(95) = "Isulas"
Enms(96) = "Séndil"
Enms(97) = "Elvarambor"
Enms(98) = "Glilmandir"
Enms(99) = "Caliondel"
Enms(100) = "Legilmariand"


'60 Dwarf Names
Dnms(1) = "Hugnar"
Dnms(2) = "Laran"
Dnms(3) = "Hefur"
Dnms(4) = "Kibur"
Dnms(5) = "Sali"
Dnms(6) = "Negnar"
Dnms(7) = "Lognar"
Dnms(8) = "Masur"
Dnms(9) = "Sognus"
Dnms(10) = "Foisur"
Dnms(11) = "Goinar"
Dnms(12) = "Venus"
Dnms(13) = "Bali"
Dnms(14) = "Sonus"
Dnms(15) = "Bunus"
Dnms(16) = "Voli"
Dnms(17) = "Sesil"
Dnms(18) = "Glulin"
Dnms(19) = "Legan"
Dnms(20) = "Ranus"
Dnms(21) = "Vorin"
Dnms(22) = "Somli"
Dnms(23) = "Goisil"
Dnms(24) = "Vafur"
Dnms(25) = "Honus"
Dnms(26) = "Tusur"
Dnms(27) = "Hoibur"
Dnms(28) = "Kosil"
Dnms(29) = "Tusin"
Dnms(30) = "Negnar"
Dnms(31) = "Glibur"
Dnms(32) = "Rilin"
Dnms(33) = "Humli"
Dnms(34) = "Glulin"
Dnms(35) = "Masil"
Dnms(36) = "Noifur"
Dnms(37) = "Koinus"
Dnms(38) = "Velin"
Dnms(39) = "Fagnus"
Dnms(40) = "Hofur"
Dnms(41) = "Fesin"
Dnms(42) = "Rugnus"
Dnms(43) = "Ririn"
Dnms(44) = "Hegnus"
Dnms(45) = "Segnar"
Dnms(46) = "Gorin"
Dnms(47) = "Kegnar"
Dnms(48) = "Duli"
Dnms(49) = "Gegan"
Dnms(50) = "Glafur"
Dnms(51) = "Soilir"
Dnms(52) = "Misin"
Dnms(53) = "Dognar"
Dnms(54) = "Soli"
Dnms(55) = "Vunar"
Dnms(56) = "Folir"
Dnms(57) = "Masur"
Dnms(58) = "Namli"
Dnms(59) = "Hoigan"
Dnms(60) = "Mignus"



'60 Halfling Names
Hnms(1) = "Peruppi"
Hnms(2) = "Radoc"
Hnms(3) = "Rebo"
Hnms(4) = "Perebo"
Hnms(5) = "Droigrin"
Hnms(6) = "Friarry"
Hnms(7) = "Suppi"
Hnms(8) = "Perabo"
Hnms(9) = "Bugo"
Hnms(10) = "Bim"
Hnms(11) = "Fragrin"
Hnms(12) = "Frerry"
Hnms(13) = "Sidoc"
Hnms(14) = "Biadoc"
Hnms(15) = "Rugo"
Hnms(16) = "Siarry"
Hnms(17) = "Rigo"
Hnms(18) = "Periagrin"
Hnms(19) = "Drarry"
Hnms(20) = "Rigo"
Hnms(21) = "Sado"
Hnms(22) = "Budo"
Hnms(23) = "Friabo"
Hnms(24) = "Sirry"
Hnms(25) = "Merudoc"
Hnms(26) = "Froigrin"
Hnms(27) = "Friado"
Hnms(28) = "Perubo"
Hnms(29) = "Meroirry"
Hnms(30) = "Biado"
Hnms(31) = "Biarry"
Hnms(32) = "Babo"
Hnms(33) = "Fruppi"
Hnms(34) = "Boidoc"
Hnms(35) = "Sego"
Hnms(36) = "Frorry"
Hnms(37) = "Perebo"
Hnms(38) = "Frubo"
Hnms(39) = "Bodoc"
Hnms(40) = "Siadoc"
Hnms(41) = "Rudoc"
Hnms(42) = "Riam"
Hnms(43) = "Peram"
Hnms(44) = "Soigrin"
Hnms(45) = "Peradoc"
Hnms(46) = "Sagrin"
Hnms(47) = "Drerry"
Hnms(48) = "Soppi"
Hnms(49) = "Sirry"
Hnms(50) = "Frego"
Hnms(51) = "Driago"
Hnms(52) = "Bugrin"
Hnms(53) = "Peroim"
Hnms(54) = "Bego"
Hnms(55) = "Bim"
Hnms(56) = "Serry"
Hnms(57) = "Robo"
Hnms(58) = "Siarry"
Hnms(59) = "Biappi"
Hnms(60) = "Robo"


'100 Orc Names
Onms(1) = "Hiol"
Onms(2) = "Virg"
Onms(3) = "Rurk"
Onms(4) = "Gung"
Onms(5) = "Vrorag"
Onms(6) = "Piurk"
Onms(7) = "Erung"
Onms(8) = "Vrorg"
Onms(9) = "Gorag"
Onms(10) = "Vaurk"
Onms(11) = "Bing"
Onms(12) = "Pralg"
Onms(13) = "Vrugdush"
Onms(14) = "Vank"
Onms(15) = "Punk"
Onms(16) = "Vrigdish"
Onms(17) = "Vashnak"
Onms(18) = "Vang"
Onms(19) = "Grogar"
Onms(20) = "Grilo"
Onms(21) = "Podash"
Onms(22) = "Bidish"
Onms(23) = "Grunk"
Onms(24) = "Prulg"
Onms(25) = "Gurg"
Onms(26) = "Gulg"
Onms(27) = "Grudush"
Onms(28) = "Bashnak"
Onms(29) = "Bogor"
Onms(30) = "Bugdish"
Onms(31) = "Griol"
Onms(32) = "Vrugor"
Onms(33) = "Godish"
Onms(34) = "Prudush"
Onms(35) = "Huk"
Onms(36) = "Rurbag"
Onms(37) = "Virbag"
Onms(38) = "Vrilo"
Onms(39) = "Pank"
Onms(40) = "Vadash"
Onms(41) = "Higdish"
Onms(42) = "Vrulg"
Onms(43) = "Hanak"
Onms(44) = "Rark"
Onms(45) = "Purbag"
Onms(46) = "Hodash"
Onms(47) = "Prurk"
Onms(48) = "Rirk"
Onms(49) = "Prirbag"
Onms(50) = "Prinak"
Onms(51) = "Vrurag"
Onms(52) = "Barbag"
Onms(53) = "Gidash"
Onms(54) = "Huurk"
Onms(55) = "Higar"
Onms(56) = "Goshnak"
Onms(57) = "Ragor"
Onms(58) = "Pranak"
Onms(59) = "Erarg"
Onms(60) = "Girag"
Onms(61) = "Ponk"
Onms(62) = "Grorag"
Onms(63) = "Vriurk"
Onms(64) = "Bik"
Onms(65) = "Varg"
Onms(66) = "Vagor"
Onms(67) = "Erugdish"
Onms(68) = "Hidish"
Onms(69) = "Gank"
Onms(70) = "Erushnak"
Onms(71) = "Rorg"
Onms(72) = "Ponk"
Onms(73) = "Bidush"
Onms(74) = "Gridash"
Onms(75) = "Pragar"
Onms(76) = "Vaol"
Onms(77) = "Hadush"
Onms(78) = "Vridush"
Onms(79) = "Prank"
Onms(80) = "Eruk"
Onms(81) = "Erart"
Onms(82) = "Erirt"
Onms(83) = "Prik"
Onms(84) = "Vrogar"
Onms(85) = "Vorg"
Onms(86) = "Gok"
Onms(87) = "Pishnak"
Onms(88) = "Bogdish"
Onms(89) = "Vradish"
Onms(90) = "Pigar"
Onms(91) = "Hudush"
Onms(92) = "Erork"
Onms(93) = "Bagdish"
Onms(94) = "Vrorg"
Onms(95) = "Rorag"
Onms(96) = "Eradish"
Onms(97) = "Grurt"
Onms(98) = "Grirt"
Onms(99) = "Vurag"
Onms(100) = "Vanak"


'50 Gnome Names
Gnms(1) = "Nwrka "
Gnms(2) = "Alora "
Gnms(3) = "Gweollyn "
Gnms(4) = "Carabryn "
Gnms(5) = "Bleddry "
Gnms(6) = "Addrarcyn "
Gnms(7) = "Glaeddry "
Gnms(8) = "Leonnyn "
Gnms(9) = "Mac "
Gnms(10) = "Gaer "
Gnms(11) = "Cecyn "
Gnms(12) = "Rhenry "
Gnms(13) = "Gaen "
Gnms(14) = "Aethullyn "
Gnms(15) = "Aetheobryn "
Gnms(16) = "Vinry "
Gnms(17) = "Vonvan "
Gnms(18) = "Caeddyn "
Gnms(19) = "Blyrcyn "
Gnms(20) = "Miran "
Gnms(21) = "Dunvan "
Gnms(22) = "Aetharcyn "
Gnms(23) = "Terraent "
Gnms(24) = "Seomyr "
Gnms(25) = "Addror "
Gnms(26) = "Rhunry "
Gnms(27) = "Addrunnyn "
Gnms(28) = "Carodd "
Gnms(29) = "Yrobryn "
Gnms(30) = "Glyrraent "
Gnms(31) = "Dubryn "
Gnms(32) = "Mamyr "
Gnms(33) = "Blaenry "
Gnms(34) = "Yrynry "
Gnms(35) = "Yrirraent "
Gnms(36) = "Aetheorcyn "
Gnms(37) = "Owoddyn "
Gnms(38) = "Seonry "
Gnms(39) = "Vedry "
Gnms(40) = "Recyn "
Gnms(41) = "Glon "
Gnms(42) = "Rhidoc "
Gnms(43) = "Jearka "
Gnms(44) = "Lyrcyn "
Gnms(45) = "Addraemyr "
Gnms(46) = "Dyran "
Gnms(47) = "Meraena "
Gnms(48) = "Ganvan "
Gnms(49) = "Rhyrraent "
Gnms(50) = "Lynoic "


If r = 1 Then
    For x = 1 To 50
        Let x1 = Int(Rnd * 50) + 1
        txtName.Text = Nm(x1)
    Next x
ElseIf r = 2 Or r = 3 Or r = 9 Then
    For x = 1 To 50
        Let x2 = Int(Rnd * 50) + 1
        txtName.Text = Enms(x2)
    Next x
ElseIf r = 4 Then
    For x = 1 To 30
        Let x3 = Int(Rnd * 30) + 1
        txtName.Text = Dnms(x3)
    Next x
ElseIf r = 5 Then
    For x = 1 To 50
        Let x4 = Int(Rnd * 50) + 1
        txtName.Text = Gnms(x4)
    Next x
ElseIf r = 6 Then
    For x = 1 To 30
        Let x5 = Int(Rnd * 30) + 1
        txtName.Text = Hnms(x5)
    Next x
ElseIf r = 7 Or r = 8 Then
    For x = 1 To 50
        Let x6 = Int(Rnd * 50) + 1
        txtName.Text = Onms(x6)
    Next x
Else
    txtName.Text = ""
End If

'PICTURES!!!
'------------
'Amrmor pics
If txtArmor.Text = "Full Armor" Then
    imghelm.Visible = True
    imgPglove.Visible = True
    imgPlegs.Visible = True
    imgParmor.Visible = True
    '----
    imgLcap.Visible = False
    imgLglove.Visible = False
    imgLlegs.Visible = False
    imgLarmor.Visible = False
    '----
    imgchaincap.Visible = False
    imgCgloves.Visible = False
    imgClegs.Visible = False
    imgCarmor.Visible = False
    '----
    imgtunic.Visible = False
    imglegs.Visible = False
ElseIf txtArmor.Text = "Leather" Then
    imgLcap.Visible = True
    imgLglove.Visible = True
    imgLlegs.Visible = True
    imgLarmor.Visible = True
    '----
    imghelm.Visible = False
    imgPglove.Visible = False
    imgPlegs.Visible = False
    imgParmor.Visible = False
    '----
    imgchaincap.Visible = False
    imgCgloves.Visible = False
    imgClegs.Visible = False
    imgCarmor.Visible = False
    '----
    imgtunic.Visible = False
    imglegs.Visible = False

ElseIf txtArmor.Text = "Chainmail" Then
    imgchaincap.Visible = True
    imgCgloves.Visible = True
    imgClegs.Visible = True
    imgCarmor.Visible = True
    '----
    imghelm.Visible = False
    imgPglove.Visible = False
    imgPlegs.Visible = False
    imgParmor.Visible = False
    '----
    imgLcap.Visible = False
    imgLglove.Visible = False
    imgLlegs.Visible = False
    imgLarmor.Visible = False
    '----
    imgtunic.Visible = False
    imglegs.Visible = False
ElseIf txtArmor.Text = "None" Then
    imgtunic.Visible = True
    imglegs.Visible = True
    '----
    imghelm.Visible = False
    imgPglove.Visible = False
    imgPlegs.Visible = False
    imgParmor.Visible = False
    '----
    imgLcap.Visible = False
    imgLglove.Visible = False
    imgLlegs.Visible = False
    imgLarmor.Visible = False
    '----
    imgchaincap.Visible = False
    imgCgloves.Visible = False
    imgClegs.Visible = False
    imgCarmor.Visible = False
End If

'Primary Weapon Pics
If txtWeap1.Text = "Long Sword" Then
    imgLsword.Visible = True
    imgSsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = False
    imgbow.Visible = False
    imgmornstar.Visible = False
    imgstaff.Visible = False
ElseIf txtWeap1.Text = "Short Sword" Then
    imgSsword.Visible = True
    imgLsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = False
    imgbow.Visible = False
    imgmornstar.Visible = False
    imgstaff.Visible = False
ElseIf txtWeap1.Text = "Mace" Then
    imgmace.Visible = True
    imgLsword.Visible = False
    imgSsword.Visible = False
    imgaxe.Visible = False
    imgbow.Visible = False
    imgmornstar.Visible = False
    imgstaff.Visible = False
ElseIf txtWeap1.Text = "Axe/Battle Axe" Then
    imgLsword.Visible = False
    imgSsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = True
    imgbow.Visible = False
    imgmornstar.Visible = False
    imgstaff.Visible = False
ElseIf txtWeap1.Text = "Bow" Then
    imgbow.Visible = True
    imgLsword.Visible = False
    imgSsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = False
    imgmornstar.Visible = False
    imgstaff.Visible = False
ElseIf txtWeap1.Text = "Morning Star" Then
    imgmornstar.Visible = True
    imgbow.Visible = False
    imgLsword.Visible = False
    imgSsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = False
    imgstaff.Visible = False
ElseIf txtWeap1.Text = "Staff" Then
    imgstaff.Visible = True
    imgmornstar.Visible = False
    imgbow.Visible = False
    imgLsword.Visible = False
    imgSsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = False
ElseIf txtWeap1.Text = "Clerical Powers" Then
    imgstaff.Visible = True
    imgmornstar.Visible = False
    imgbow.Visible = False
    imgLsword.Visible = False
    imgSsword.Visible = False
    imgmace.Visible = False
    imgaxe.Visible = False
End If

'Secondary Waepons Pics
If txtWeap2.Text = "Long Sword" Then
    imgLsword2.Visible = True
    imgSsword2.Visible = False
    imgmace2.Visible = False
    imgmornstar2.Visible = False
    imgbow2.Visible = False
    imgdagger.Visible = False
ElseIf txtWeap2.Text = "Short Sword" Then
    imgLsword2.Visible = False
    imgSsword2.Visible = True
    imgmace2.Visible = False
    imgmornstar2.Visible = False
    imgbow2.Visible = False
    imgdagger.Visible = False
ElseIf txtWeap2.Text = "Mace" Then
    imgLsword2.Visible = False
    imgSsword2.Visible = False
    imgmace2.Visible = True
    imgmornstar2.Visible = False
    imgbow2.Visible = False
    imgdagger.Visible = False
ElseIf txtWeap2.Text = "Morning Star" Then
    imgLsword2.Visible = False
    imgSsword2.Visible = False
    imgmace2.Visible = False
    imgmornstar2.Visible = True
    imgbow2.Visible = False
    imgdagger.Visible = False
ElseIf txtWeap2.Text = "Bow" Then
    imgLsword2.Visible = False
    imgSsword2.Visible = False
    imgmace2.Visible = False
    imgmornstar2.Visible = False
    imgbow2.Visible = True
    imgdagger.Visible = False
ElseIf txtWeap2.Text = "Dagger" Then
    imgLsword2.Visible = False
    imgSsword2.Visible = False
    imgmace2.Visible = False
    imgmornstar2.Visible = False
    imgbow2.Visible = False
    imgdagger.Visible = True
End If
End Sub


