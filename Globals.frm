VERSION 5.00
Begin VB.Form Globals 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Global Variables"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   345
      Left            =   2880
      TabIndex        =   14
      Top             =   660
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2880
      TabIndex        =   13
      Top             =   1020
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   345
      Left            =   2880
      TabIndex        =   12
      Top             =   210
      Width           =   1065
   End
   Begin VB.TextBox Tens 
      Height          =   285
      Left            =   1530
      TabIndex        =   11
      Top             =   1530
      Width           =   975
   End
   Begin VB.TextBox Grav 
      Height          =   285
      Left            =   1530
      TabIndex        =   10
      Top             =   210
      Width           =   975
   End
   Begin VB.TextBox Bounce 
      Height          =   285
      Left            =   1530
      TabIndex        =   9
      Top             =   2460
      Width           =   975
   End
   Begin VB.TextBox Fric 
      Height          =   285
      Left            =   1530
      TabIndex        =   8
      Top             =   1980
      Width           =   975
   End
   Begin VB.TextBox Wind 
      Height          =   285
      Left            =   1530
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Atmos 
      Height          =   285
      Left            =   1530
      TabIndex        =   6
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wall Bounce"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   2550
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wall Friction"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   2040
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Tension"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wind"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   1140
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atmosphere"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   690
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gravity"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Globals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Atmos_Change()
If (Val(Atmos)) >= 1 Then Atmos = "1"
If (Val(Atmos)) < 0 Then Atmos = "0"
End Sub

Private Sub Command1_Click()

Gravity = Grav / 2
Atmosphere = Atmos
LeftWind = Wind
Tension = Tens
WallFriction = Fric
WallBounce = Bounce

Form1.Slider1.Value = LeftWind
Form1.Slider1_Click
Form1.Slider2.Value = Atmosphere * 100
Form1.Slider2_click
Form1.Slider3.Value = Gravity * 100
Form1.Slider3_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

Gravity = Grav / 2
Atmosphere = Atmos
LeftWind = Wind
Tension = Tens
WallFriction = Fric
WallBounce = Bounce

Form1.Slider1.Value = LeftWind
Form1.Slider1_Click
Form1.Slider2.Value = Atmosphere * 100
Form1.Slider2_click
Form1.Slider3.Value = Gravity * 100
Form1.Slider3_Click


Unload Me

End Sub

Private Sub Form_Load()

Grav = Gravity * 2
Atmos = Atmosphere
Wind = LeftWind
Tens = Tension
Fric = WallFriction
Bounce = WallBounce

End Sub

Private Sub Grav_Change()
If (Val(Grav)) > 2 Then Grav = "2"
If (Val(Grav)) < -2 Then Grav = "-2"
End Sub

Private Sub Wind_Change()
If (Val(Wind)) > 20 Then Wind = "20"
If (Val(Wind)) < -20 Then Wind = "-20"

End Sub
