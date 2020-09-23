VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Circus Slots !~[715]~!"
   ClientHeight    =   2850
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4200
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   2850
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bet"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton Option3 
         Caption         =   "$75"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "$50"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "$25"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4560
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4680
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Spin All"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "You Just Won:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label text3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label text2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +++++++++++++++++=========================================++++++++++++++++++++
' +++++++++++++++++=========================================++++++++++++++++++++
' +++++++++++++++++        Greetings To You All!            ++++++++++++++++++++
' +++++++++++++++++     Sorry for the crapy coding :(       ++++++++++++++++++++
' +++++++++++++++++       Code By: Justin Lilley            ++++++++++++++++++++
' +++++++++++++++++         And By: Nick DeLuca             ++++++++++++++++++++
' +++++++++++++++++             Please Vote!                ++++++++++++++++++++
' +++++++++++++++++       Check Out Mindzpro.com!           ++++++++++++++++++++
' +++++++++++++++++     Certain Copyrights may apply        ++++++++++++++++++++
' +++++++++++++++++=========================================++++++++++++++++++++
' +++++++++++++++++=========================================++++++++++++++++++++
Dim i ' Define your dims
Dim j
Dim l
Dim ii
Dim jj
Dim ll
Dim money
Dim push
Dim kitty


Private Sub Command1_Click()
kitty = True
ii = True
jj = True
ll = True
Timer1.Enabled = True 'Turns Timers ON
Timer2.Enabled = True
Timer3.Enabled = True
Label2.Caption = ""
If Option1.Value = True Then ' Checkes if Option1 is checked
money = "$" & money - 25
Label1.Caption = money 'Defines Label as Money
End If
If Option2.Value = True Then
If ((money - 50) < 0) Then
push = MsgBox("You Don't Have Enough Money", vbExclamation, "Error") ' Displays you dont have enough money to continue
Timer1.Enabled = False ' Turns timers OFF
Timer2.Enabled = False
Timer3.Enabled = False
text1.Caption = "" ' clears label
text2.Caption = ""
text3.Caption = ""
Command4.Enabled = False ' Disables stop buttons
Command3.Enabled = False
Command2.Enabled = False
End If
If ((money - 50) > 0) Then
money = "$" & money - 50 ' checks for money
Label1.Caption = money
End If
End If
If Option3.Value = True Then
If ((money - 75) < 0) Then
push = MsgBox("You Don't Have Enough Money", vbExclamation, "Error") ' see above
Timer1.Enabled = False
Timer2.Enabled = False ' Turns timers OFF
Timer3.Enabled = False
text1.Caption = ""
text2.Caption = "" ' clears label
text3.Caption = ""
Command4.Enabled = False
Command3.Enabled = False 'Disables stop buttons
Command2.Enabled = False
End If
If ((money - 50) > 0) Then
money = "$" & money - 75
Label1.Caption = money
End If
End If

If Command1.Enabled = True Then
Command4.Enabled = True
Command3.Enabled = True
Command2.Enabled = True 'Enables stop buttons
End If

End Sub

Private Sub Command2_Click()

Timer1.Enabled = False  '.... IM confused
Beep
Command1.Enabled = False
Command2.Enabled = False
If Command2.Enabled = False And Command3.Enabled = False And Command4.Enabled = False Then
Command1.Enabled = True
If text1.Caption = text2.Caption Or text1.Caption = text3.Caption Or text3.Caption = text2.Caption Then
kitty = False
If Option1.Value = True Then
money = money + 50
Label2.Caption = "$" & 50
End If
If Option2.Value = True Then
money = money + 100
Label2.Caption = "$" & 100
End If
If Option3.Value = True Then
money = money + 150
Label2.Caption = "$" & 150
End If
Label1.Caption = "$" & money
End If
If text1.Caption = text2.Caption And text1.Caption = text3.Caption Then
kitty = False
If Option1.Value = True Then
money = money + 75
Label2.Caption = "$" & 75
End If
If Option2.Value = True Then
money = money + 150
Label2.Caption = "$" & 150
End If
If Option3.Value = True Then
money = money + 225
Label2.Caption = "$" & 225
End If
Label1.Caption = "$" & money
End If
If kitty = True Then
Label2.Caption = "$" & 0
End If
End If

End Sub

Private Sub Command3_Click()

Timer2.Enabled = False ' same as above the above :O
Beep
Command1.Enabled = False
Command3.Enabled = False
If Command2.Enabled = False And Command3.Enabled = False And Command4.Enabled = False Then
Command1.Enabled = True
If text1.Caption = text2.Caption Or text1.Caption = text3.Caption Or text3.Caption = text2.Caption Then
kitty = False
If Option1.Value = True Then
money = money + 50
Label2.Caption = "$" & 50
End If
If Option2.Value = True Then
money = money + 100
Label2.Caption = "$" & 100
End If
If Option3.Value = True Then
money = money + 150
Label2.Caption = "$" & 150
End If
Label1.Caption = "$" & money
End If
If text1.Caption = text2.Caption And text1.Caption = text3.Caption Then
kitty = False
If Option1.Value = True Then
money = money + 75
Label2.Caption = "$" & 75
End If
If Option2.Value = True Then
money = money + 150
Label2.Caption = "$" & 150
End If
If Option3.Value = True Then
money = money + 225
Label2.Caption = "$" & 225
End If
Label1.Caption = "$" & money
End If
If kitty = True Then
Label2.Caption = "$" & 0
End If
End If

End Sub

Private Sub Command4_Click()

Timer3.Enabled = False ' Same as the above the above the above =/
Beep
Command1.Enabled = False
Command4.Enabled = False
If Command2.Enabled = False And Command3.Enabled = False And Command4.Enabled = False Then
Command1.Enabled = True
If text1.Caption = text2.Caption Or text1.Caption = text3.Caption Or text3.Caption = text2.Caption Then
kitty = False
If Option1.Value = True Then
money = money + 50
Label2.Caption = "$" & 50
End If
If Option2.Value = True Then
money = money + 100
Label2.Caption = "$" & 100
End If
If Option3.Value = True Then
money = money + 150
Label2.Caption = "$" & 150
End If
Label1.Caption = "$" & money
End If
If text1.Caption = text2.Caption And text1.Caption = text3.Caption Then
kitty = False
If Option1.Value = True Then
money = money + 75
Label2.Caption = "$" & 75
End If
If Option2.Value = True Then
money = money + 150
Label2.Caption = "$" & 150
End If
If Option3.Value = True Then
money = money + 225
Label2.Caption = "$" & 225
End If
Label1.Caption = "$" & money
End If
If kitty = True Then
Label2.Caption = "$" & 0
End If
End If

End Sub

Private Sub Command5_Click()
Unload Me
frmAbout.Show ' Gotta see the about :)
End Sub

Private Sub Command6_Click()
frmAbout.Show ' LOOKY!
End Sub

Private Sub Form_Load()
Timer1.Enabled = False ' enables timers
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = True
Command4.Enabled = False ' disables stop buttons
Command3.Enabled = False
Command2.Enabled = False
Option1.Value = True
money = 500
Label1.Caption = "$" & money
kitty = True

End Sub





Private Sub mnuexit_Click()
Unload Me
frmAbout.Show ' Did you see the about yet? =]
End Sub

Private Sub Timer1_Timer()
If ii = True Then 'Creats the Number
i = Int(Rnd * 5)
End If
ii = False
i = i + 1
text1.Caption = i
If i = 6 Then
i = 0
End If

End Sub

Private Sub Timer2_Timer()
If jj = True Then ' ^^^
j = Int(Rnd * 5)
End If
jj = False
j = j + 1
text2.Caption = j
If j = 6 Then
j = 0
End If
End Sub

Private Sub Timer3_Timer()
If ll = True Then
l = Int(Rnd * 5) ' ^^^
End If
ll = False
l = l + 1
text3.Caption = l
If l = 6 Then
l = 0
End If
End Sub
Private Sub reset()
money = 500
Label1.Caption = "$" & money
Timer1.Enabled = False
Timer2.Enabled = False ' disables timers
Timer3.Enabled = False
text1.Caption = ""
text2.Caption = ""
text3.Caption = ""
End Sub

Private Sub Timer4_Timer()
If money <= 0 Then
push = MsgBox("You have spent all your money", vbExclamation, "Loser")
reset
End If ' you've heard it before :)
End Sub
