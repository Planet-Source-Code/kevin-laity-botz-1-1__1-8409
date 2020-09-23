VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Botz"
   ClientHeight    =   7095
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Frame_ControlPanel 
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   30
      ScaleHeight     =   1905
      ScaleWidth      =   9525
      TabIndex        =   7
      Top             =   5160
      Width           =   9525
      Begin VB.CommandButton btnResetLinks 
         Caption         =   "Reset Links"
         Height          =   360
         Left            =   3390
         TabIndex        =   34
         ToolTipText     =   "Default Links to their current length"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton btnSave 
         Height          =   360
         Left            =   750
         Picture         =   "Form1.frx":1D82
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Save Bot"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton btnNew 
         Height          =   360
         Left            =   30
         Picture         =   "Form1.frx":202C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "New Bot"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton btnOpen 
         Height          =   360
         Left            =   390
         Picture         =   "Form1.frx":2242
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Load Bot"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton BTNpicCon 
         DisabledPicture =   "Form1.frx":24F4
         DownPicture     =   "Form1.frx":29E6
         Height          =   360
         Left            =   1230
         Picture         =   "Form1.frx":2ED8
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Construct Mode (F2)"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton BTNpicSim 
         DisabledPicture =   "Form1.frx":33CA
         DownPicture     =   "Form1.frx":38BC
         Height          =   360
         Left            =   1590
         Picture         =   "Form1.frx":3DAE
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Simulate Mode (F3)"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton btnAddwheel 
         DisabledPicture =   "Form1.frx":42A0
         Height          =   360
         Left            =   2070
         Picture         =   "Form1.frx":4656
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Add Wheel to Vertex (Ctrl+W)"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton btnDelete 
         DisabledPicture =   "Form1.frx":4A0C
         Height          =   360
         Left            =   2430
         Picture         =   "Form1.frx":4BFE
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Delete Current Vertex or Link"
         Top             =   30
         Width           =   360
      End
      Begin VB.CommandButton BTNGlobals 
         Height          =   360
         Left            =   2910
         Picture         =   "Form1.frx":4DF0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Change Global Variables"
         Top             =   30
         Width           =   360
      End
      Begin VB.Frame Frame2 
         Caption         =   "Muscle Simulation"
         Height          =   1485
         Left            =   6120
         TabIndex        =   20
         Top             =   30
         Width           =   3345
         Begin VB.CommandButton btnRevers 
            Caption         =   "Reverse"
            Height          =   285
            Left            =   2430
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   990
            Width           =   735
         End
         Begin VB.PictureBox cycleview 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000008&
            ForeColor       =   &H80000008&
            Height          =   600
            Left            =   180
            ScaleHeight     =   38
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   198
            TabIndex        =   22
            Top             =   330
            Width           =   3000
         End
         Begin VB.CommandButton btnPause 
            Caption         =   "Pause"
            Height          =   285
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   990
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   240
            Left            =   930
            TabIndex        =   24
            Top             =   1020
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   423
            _Version        =   393216
            Orientation     =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label INDspd 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Spd: 3"
            Height          =   195
            Left            =   150
            TabIndex        =   25
            Top             =   1050
            Width           =   675
         End
      End
      Begin VB.Frame VBOX 
         Caption         =   "Vertex 0"
         Height          =   1485
         Left            =   4740
         TabIndex        =   13
         Top             =   30
         Width           =   1335
         Begin VB.TextBox INDrad 
            Height          =   255
            Left            =   390
            TabIndex        =   16
            Text            =   "0"
            Top             =   1110
            Width           =   525
         End
         Begin VB.TextBox txtY 
            Height          =   225
            Left            =   270
            TabIndex        =   15
            Text            =   "0"
            Top             =   600
            Width           =   405
         End
         Begin VB.TextBox txtX 
            Height          =   255
            Left            =   270
            TabIndex        =   14
            Text            =   "0"
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Radius"
            Height          =   195
            Left            =   390
            TabIndex        =   19
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Y"
            Height          =   225
            Left            =   90
            TabIndex        =   18
            Top             =   630
            Width           =   165
         End
         Begin VB.Label Label1 
            Caption         =   "X"
            Height          =   225
            Left            =   90
            TabIndex        =   17
            Top             =   360
            Width           =   165
         End
      End
      Begin VB.Frame LBOX 
         Caption         =   "Link 0"
         Height          =   1485
         Left            =   4740
         TabIndex        =   10
         Top             =   30
         Width           =   1335
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Length"
            Height          =   195
            Left            =   420
            TabIndex        =   12
            Top             =   450
            Width           =   495
         End
         Begin VB.Label INDlen 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   195
            Left            =   135
            TabIndex        =   11
            Top             =   660
            Width           =   1065
         End
      End
      Begin VB.CheckBox CHKTop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ceiling Enabled"
         Height          =   195
         Left            =   1170
         TabIndex        =   9
         Top             =   1350
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2940
         TabIndex        =   8
         Text            =   "Presets..."
         Top             =   630
         Width           =   1575
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   630
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   -20
         Max             =   20
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   945
         Left            =   270
         TabIndex        =   37
         Top             =   660
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1667
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label ModeIND 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   42
         Top             =   1620
         Width           =   9525
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "No Selection"
         Height          =   525
         Left            =   5010
         TabIndex        =   41
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Wind"
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   450
         Width           =   1605
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Air Friction"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   900
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.PictureBox VHoverVertex 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1680
      Picture         =   "Form1.frx":52E2
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   6
      Top             =   690
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox VSelVertex 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1440
      Picture         =   "Form1.frx":5420
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   5
      Top             =   690
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox CycleBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   60
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox Buffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   90
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.OptionButton focusdummy 
      Caption         =   "Option1"
      Height          =   405
      Left            =   -840
      TabIndex        =   2
      Top             =   4080
      Width           =   795
   End
   Begin VB.PictureBox VDot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   1200
      Picture         =   "Form1.frx":555E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Main 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   0
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   900
      Top             =   690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.botz|Bot Files|*.*|All Files"
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuglobal 
         Caption         =   "Set Global Variables"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuAutoRev 
         Caption         =   "Auto Reverse"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVert 
         Caption         =   "&Vertices"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLink 
         Caption         =   "&Links"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuhandl 
         Caption         =   "&Link Handles"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu fs 
         Caption         =   "Full Screen"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnusel 
      Caption         =   "Vertex"
      Begin VB.Menu mnuVertexDelete 
         Caption         =   "Delete Vertex"
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVertexAddwheel 
         Caption         =   "Add Wheel..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuVertexDeleteWheel 
         Caption         =   "Delete Wheel"
      End
   End
   Begin VB.Menu mnulinkthing 
      Caption         =   "Link"
      Begin VB.Menu mnuLinkDelete 
         Caption         =   "Delete Link                Del"
      End
      Begin VB.Menu mnudash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinkLengthReset 
         Caption         =   "Reset Link Length"
      End
      Begin VB.Menu mnuDeleteMuscle 
         Caption         =   "Delete Muscle"
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "&Mode"
      Begin VB.Menu mnuConstruct 
         Caption         =   "Construct"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSimulate 
         Caption         =   "Simulate"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
'
'                                  Botz 1.00
'
' I've done everything I can to make this code readable and understandable.
' Just about every procedure is thoroughly commented, and divided into
' nice, manageable segments to help with comprehension.
'
' Note: When the program is running, everything on the screen is flipped.
'       That way, the 'floor' is at y-coordinate 0.  VB usually has the
'       top of an object as y=0.  However I found that made things harder
'       as I kept having to remember to reverse anything that was vertical.
'
'       Anyways the point is that if you fiddle with the physics cycle,
'       remember that y=0 is the floor, not the ceiling.
'
'
' You can do whatever you want to this code for non-commercial purposes.
' If you want to use some of this code for another project please let me
' know, and give me a little recognition.
'
' If you make improvements to this code please let me know so I can
' include it in future releases.
'
' If you need to contact me, my email addy is KevinLaity@Cadvision.com
' No bug reports unless the bug is important and can be replicated
'--------------------------------------------------------------------------




Dim vertex(200) As vertex_type
Const MaxVertices As Integer = 200
Const pi As Double = 3.14159265358979
'The maximum number of vertexes is 300, after 300 no more may be added
'By changing the dimensions of vertex and the value of MaxVertexs, you
'can change the maximum

Dim Link(200) As link_type
Const MaxLinks As Integer = 200
'The maximum number of links is also 300.

Dim MouseX, MouseY As Single
'Keep track of the mouse position on the main board.

Dim VertexCount As Integer
Dim LinkCount As Integer
'Keep track of how many links and vertices there are



Dim CycleTime As Integer
Const CycleSize As Integer = 200
Dim ClockPause As Boolean
'CycleTime - the current place in the muscle cycle
'CycleSize - the amount that the value of cycletime must reach before
'            wrapping around.
'ClockPause - whether the muscle cycle is paused
'ClockSpeed - how fast the muscle cycle goes per turn
'             --Defined in Declares.bas


Dim FS_Mode As Boolean 'Full screen mode
Dim mode As Byte
Dim SubMode As Byte
Dim SubModeData As Integer
'Mode 0 = Construction Mode
'Mode 1 = Simulation Mode
'SubMode 0 = Nothing.  If in mode 0, ready to create a vertex.
'SubMode 1 = Creating a link
'SubMode 2 = Dragging a Vertex
'SubMode 4 = Link is selected
'SubModeData = Whatever needed data is relevant to the current submode
'              eg. in SubMode 1, the location of the vertex to start
'                  the link from.


Dim SelVertex
Dim SelLink
Dim HoverVertex
Dim HoverLink As Integer
Dim DragDot As Integer
'Sel* - number of the vertex or link that is currently selected.
'       (if any) 0 = none
'Hover* - The number of the vertex or link that the mouse is currently
'         over (if any) 0 = none
'DragDot - which vertex is being dragged

Dim BoardX As Integer 'width
Dim BoardY As Integer 'height
Dim RightWall As Single
Dim Ceiling As Single
'These are the dimensions of the "playing field"

Dim AutoReverseCycle As Integer

Dim DrawColor, BGColor As Long
Dim CurrentPhase As Byte














Function AddLink(id1, id2) As Boolean

'This Subroutine adds a link between two vertices.
'The link will push or pull them to keep them at the same distance
'They were at when the link was created.
'This action is what makes it possible to make objects to stand up.

If id1 = id2 Then AddLink = False: Exit Function 'you can't link a vertex to itself



'Check to see if this link exists already  --------------------------
For i = 1 To MaxLinks
  With Link(i)
  If .used = True Then
    If .target1_id = id1 And .target2_id = id2 Then AddLink = False: Exit Function
    If .target1_id = id2 And .target2_id = id1 Then AddLink = False: Exit Function
  End If
  End With
If id1 = id2 Then AddLink = False: Exit Function 'you can't link a vertex to itself

Next i
'--------------------------------------------------------------------







'Make the link-------------------------------------------------------
Dim CurLink As Integer
Dim XLen, YLen, Leng As Single

For i = 1 To 300                   'Find a link number that is
  If Link(i).used = False Then     'not being used.  And use
     CurLink = i                   'Curlink'
     Exit For
  End If
Next i

With Link(CurLink)
    .target1_id = id1 'Each Target_id is the number of one of the
    .target2_id = id2 'vertices the link is attached to.
    
    XLen = (vertex(.target2_id).X - vertex(.target1_id).X)
    YLen = (vertex(.target2_id).y - vertex(.target1_id).y)
    Leng = Sqr(Abs(XLen ^ 2 + YLen ^ 2))
            'Calculate length.
    
    .linklength = Leng
    .used = True
    .midx = vertex(.target2_id).X + ((vertex(.target1_id).X - vertex(.target2_id).X) / 2)
    .midy = vertex(.target2_id).y + ((vertex(.target1_id).y - vertex(.target2_id).y) / 2)
    .linktension = Tension
    .pushstrength = 0
    .pushtiming = 180
    .pushspan = 40
    .phase = CurrentPhase
End With

LinkCount = LinkCount + 1  'Count the new link.
AddLink = True             'Hey it worked.
'--------------------------------------------------------------------

End Function

Function AddLink2(id1, id2, Leng, Tens, pspan, pushr, stren, lastlen, timing, phase As Variant) As Boolean



'This Function is just like AddLink but it has more parameters.
'The extra parameters are needed when loading from a file.
'See Addlink for comments and explanations


If id1 = id2 Then AddLink2 = False: Exit Function
Dim CurLink As Integer
For i = 1 To 300
  If Link(i).used = False Then
     CurLink = i
     Exit For
  End If
Next i
LinkCount = LinkCount + 1

With Link(CurLink)
    .target1_id = id1
    .target2_id = id2
    .linklength = Leng
    .used = True
    .linktension = Tens
    .midx = vertex(.target2_id).X + ((vertex(.target1_id).X - vertex(.target2_id).X) / 2)
    .midy = vertex(.target2_id).y + ((vertex(.target1_id).y - vertex(.target2_id).y) / 2)
    .pushspan = pspan
    .Push = pushr
    .pushstrength = stren
    .pushtiming = timing
    .lastlen = lastlen
    .phase = phase
    '.Phase = 1
End With
AddLink2 = True


End Function




Function AddVertex(X, y, MomentX, MomentY, Radius, MomentC, Optional phase As Variant) As Integer


'MsgBox Phase
If phase = "" Then phase = 1



'This Subroutine adds a vertex.

'Find a vertex number thats not taken ------------------------------
    For i = 1 To MaxVertices
     If vertex(i).used = False Then
        currentvertex = i
        Exit For
      End If
    Next i
'-------------------------------------------------------------------


'Make the Vertex ---------------------------------------------------
With vertex(currentvertex)

    .used = True
    .X = X
    .y = y
    .momentum_x = MomentX
    .momentum_y = MomentY
    .momentum_c = MomentC
    .heading = 0
    
    SelVertex = currentvertex      'When a vertex is created it
    VertexCount = VertexCount + 1  'is automatically selected
    .Selected = True
    
    If Radius > 0 Then             'For purposes of loading from
      .wheel = True                'a file.  By default a vertex
      .Radius = Radius             'has no wheel attached.
    Else
      .wheel = False
      .Radius = 0
    End If
    .phase = phase
End With
'-------------------------------------------------------------------

AddVertex = currentvertex   'return the number of the vertex

End Function


Sub AddWheel(vertx, Radius As Integer)

With vertex(vertx)
  .wheel = True
  .Radius = Radius
  .heading = 0 'the default heading is arbitrarily 0
End With

'This should be pretty straightforward

End Sub

Sub Cycle_Botz(delay As Single)


Static Start As Long
Dim timedelay As Boolean
Dim Elapsed As Long
Dim actualdelay As Single

Start = GetTickCount
'Time Delay Code --------------------------------------------------
restarter:
If mode = 0 Then actualdelay = 15
If mode = 1 Then actualdelay = delay
                                       'This code will delay
  Do                                    'the amount of time in
    DoEvents                            'the delay Variable.
    Elapsed = GetTickCount              'since it compares itself
    DoEvents                            'with the system time
    
    If (Elapsed - Start) >= actualdelay Then  'this will be more
        timedelay = True                'accurate than a timer
        DoEvents                        'control.
        Start = 0
    Else: timedelay = False
    End If
    DoEvents
  Loop While timedelay = False
'------------------------------------------------------------------
Start = GetTickCount

'Perform Normal Operations ----------------------------------------
     DoEvents           'lots of these are neccessary or else
                        'the program will slow down other programs
     
     Cycle_Misc
     DoEvents
     If mode = 1 Then X = Cycle_Physics '-Only do the physics cycle
                                       'if we are in Simulate mode.
                                       'Physics do not need to be
                                       'simulated in construct mode.
     DoEvents
     Cycle_Display  'Display everything.
     DoEvents
'------------------------------------------------------------------

GoTo restarter 'loop!

End Sub

Sub ClearMultiSelect()

For i = 1 To MaxVertices
   vertex(i).Selected = False
Next i


SelVertex = 0

End Sub

Function DeleteVertex(id) As Boolean

'First, delete any links attached to this vertex

For i = 1 To MaxLinks
   With Link(i)
     If .target1_id = id Or .target2_id = id Then
     .used = False  'by setting .used to false,
                    'the link is no longer displayed, and it is
                    'exempted from the physics cycle
     End If
   End With
Next i
vertex(id).used = False 'Delete the vertex itself
DeleteVertex = True 'It worked.

End Function

Sub Cycle_Display()


'Check to see if the user has hit the maximize button
'instead of using F11 to go fullscreen.
If Me.WindowState = 2 And fs.Checked = False Then Call fs_Click
If Me.WindowState = 0 And fs.Checked = True Then Call fs_Click

    

'watch out, this subroutine will get confusing
'but not as confusing as the physics cycle ;)


Dim VX, VY, HT As Integer 'some temporary variables
Dim color1 As Long


Buffer.Picture = Nothing          'clear the buffer
HT = Buffer.Height - 3            '-for tweaking purposes




'Draw Links, Link Handles, etc. --------------------------------------
For i = 1 To MaxLinks
  With Link(i)
  
        .midx = vertex(.target2_id).X + ((vertex(.target1_id).X - vertex(.target2_id).X) / 2)
        .midy = vertex(.target2_id).y + ((vertex(.target1_id).y - vertex(.target2_id).y) / 2)
         'calculate link handles, those are the little +'s you'll see
         'in the middle of each link if you turn on View->Link Handles
         'it lets you know where to put your mouse to select a link
         'Also, I reused these calculations for determining whether the
         'mouse is over the link handle or not.
  
  If .used = True Then 'ignore unused links
     
     If SelLink = i Then color1 = RGB(0, 0, 255)
     If mnuLink.Checked = True Then Buffer.Line (vertex(.target1_id).X, HT - vertex(.target1_id).y)-(vertex(.target2_id).X, HT - vertex(.target2_id).y), Phase_Color(.phase)
     If mnuhandl.Checked = True And HoverLink <> i Then
         Buffer.Line (.midx, HT - .midy - 5)-(.midx, HT - .midy + 5), color1
         Buffer.Line (.midx + 5, HT - .midy)-(.midx - 5, HT - .midy), color1
     End If
     'this will only make visible a link if View->Links is checked
     'and only make visible a link handle if View->Link Handles is checked
     
     
     VX = (.midx) - (5 / 2)
     VY = (HT - .midy) - (5 / 2)
     If HoverLink = i Then vari = BitBlt(Buffer.hDC, VX - 2, VY - 2, 9, 9, VHoverVertex.hDC, 0, 0, SRCCOPY)
     'If the mouse is over this link then display a little blue
     'circle over it (stored in a picturebox named VHoverVertex).
     
         
     color1 = RGB(128, 128, 128)
     If SelLink = i Then color1 = RGB(0, 0, 255)
     If HoverLink = i Then color1 = RGB(255, 0, 0)
     CycleBuffer.Line (.pushtiming + 2, 19 - .pushstrength)-(.pushtiming - .pushspan, 19), color1
     CycleBuffer.Line (.pushtiming + 2, 19 - .pushstrength)-(.pushtiming + .pushspan, 19), color1
     'This draws the lines on the muscle cycle. (the rectangle under
     'the main area).  Grey for inactive, Blue for selected link,
     'Red for hovering link
     
     If .pushtiming + .pushspan > CycleSize Then CycleBuffer.Line (.pushtiming + 2 - CycleSize, 19 - .pushstrength)-(.pushtiming + .pushspan - CycleSize, 19), color1
     If .pushtiming - .pushspan < 0 Then CycleBuffer.Line (.pushtiming + 2 + CycleSize, 19 - .pushstrength)-(.pushtiming - .pushspan + CycleSize, 19), color1
     'draw lines that wrap around
     
  End If
  End With
Next i
'---------------------------------------------------------------------

DoEvents


'Verticies and Wheels ------------------------------------------------
If mnuVert.Checked = True Then  'only display this stuff if
For i = 1 To MaxVertices        'view->vertices is checked
  With vertex(i)
  If .used = True Then
      

    
      VX = (.X) - (2)
      VY = (HT - .y) - (2)
      'vari = BitBlt(Buffer.hdc, VX, VY, 5, 5, VDot.hdc, 0, 0, SRCCOPY)
      Buffer.FillStyle = 0
      Buffer.FillColor = Phase_Color(.phase)
      Buffer.Circle (VX + 2, VY + 2), 2, Phase_Color(.phase)
      'display the dot in the appropriate location
      
      If HoverVertex = i Then vari = BitBlt(Buffer.hDC, VX - 2, VY - 2, 9, 9, VHoverVertex.hDC, 0, 0, SRCCOPY)
      If SelVertex = i Then vari = BitBlt(Buffer.hDC, VX - 2, VY - 2, 9, 9, VSelVertex.hDC, 0, 0, SRCCOPY)
      

      If .Selected = True Then vari = BitBlt(Buffer.hDC, VX - 2, VY - 2, 9, 9, VSelVertex.hDC, 0, 0, SRCCOPY)
      'display the hovering dot image or the selected dot image
      'if the current vertex is selected or hovered over
      
      
      
      If .wheel = True Then
           Buffer.FillStyle = 1 'so the wheels won't be solid
           Buffer.Circle (VX + 2, VY + 2), .Radius, DrawColor
           Display_MakeCircleSpokes i, .heading, 3 '3 is the number of spokes
                                           'the wheel has
                                           'you can set this to any number
      End If
      'create the wheel image if there is a wheel

    
   End If
  End With
Next i
End If
'---------------------------------------------------------------------


DoEvents

'Display Everything --------------------------------------------------
If mode = 0 And SubMode = 1 Then Buffer.Line (vertex(SelVertex).X, HT - vertex(SelVertex).y)-(MouseX, MouseY)
    'if its ready to make a new link, draw a line from the
    'starting vertex to the mouse



vari = BitBlt(Main.hDC, 0, 0, BoardX, BoardY, Buffer.hDC, 0, 0, SRCCOPY)
    'blt from the buffer to the main screen

CycleBuffer.Line (0, 19)-(CycleSize, 19)
CycleBuffer.Line (CycleTime, 0)-(CycleTime, 40)
CycleBuffer.Line (CycleSize / 2, 0)-(CycleSize / 2, 40), RGB(128, 128, 128)
CycleBuffer.Line (CycleSize / 4, 0)-(CycleSize / 4, 40), RGB(128, 128, 128)
CycleBuffer.Line ((CycleSize / 4) + (CycleSize / 2), 0)-((CycleSize / 4) + (CycleSize / 2), 40), RGB(128, 128, 128)
    '-the lines on the muscle cycle

vari = BitBlt(cycleview.hDC, 0, 0, CycleSize, 40, CycleBuffer.hDC, 0, 0, SRCCOPY)
    'blt from the cycle buffer to the main cycle picturebox


Main.Refresh
cycleview.Refresh

    'Gotta refresh these or you won't see squat
'---------------------------------------------------------------------



End Sub

Function HowManySelected()

Dim counter As Integer

For i = 1 To MaxVertices
   If vertex(i).Selected = True And vertex(i).used = True Then
      counter = counter + 1
   End If
Next i
HowManySelected = counter


End Function

Sub Display_MakeCircleSpokes(dot, heading, spokenumber)

Dim xer, yer, HT As Double
Dim subheading As Double
Dim color1 As Long


HT = Buffer.Height - 3


color1 = RGB(0, 200, 100)

With vertex(dot)
For i = 0 To (spokenumber - 1)
    xer = 0: yer = 0
    subheading = heading + ((360 / spokenumber) * i)
    xer = Sin(subheading * (pi / 180)) * .Radius
    yer = Cos(subheading * (pi / 180)) * .Radius
    Buffer.Line (.X, HT - .y)-Step(xer, yer), color1
Next i
End With

End Sub

Sub Cycle_Misc()

'In this cycle we do everything that has nothing to do with display,
'but can't be put in the physics cycle because they have to happen
'when the program is in Construct mode

CycleBuffer.Picture = Nothing 'clear the muscle cycle buffer


'Show context items ------------------------------------------------------
If SubMode <> 4 Then SelLink = 0
    If SubMode = 4 Then
        With Link(SelLink)
        If .used = True Then
        LBOX.Visible = True
        VBOX.Visible = False
        If LBOX <> "Link " & SelLink Then LBOX = "Link " & SelLink
        If INDlen <> Link(SelLink).linklength Then INDlen = Link(SelLink).linklength: INDlen.Refresh
        vari = BitBlt(CycleBuffer.hDC, .pushtiming, 19 - .pushstrength - 2, 5, 5, VDot.hDC, 0, 0, SRCCOPY)
        End If
        End With
    Else
        If SelVertex = 0 Then
           VBOX.Visible = False
           LBOX.Visible = False
         Else
           If VBOX <> "Vertex " & SelVertex Then VBOX = "Vertex " & SelVertex
           VBOX.Visible = True
           LBOX.Visible = False
           mnuVert.Visible = True
           If mode = 0 Then txtX.Enabled = True: txtY.Enabled = True
           If mode = 1 Then txtX.Enabled = False: txtY.Enabled = False
           txtX = Int(vertex(SelVertex).X)
           txtY = Int(vertex(SelVertex).y)
           INDrad = vertex(SelVertex).Radius
        End If
    End If
'------------------------------------------------------------------------
    
    
If SelLink > 0 Then
   mnulinkthing.Visible = True
Else
   mnulinkthing.Visible = False
End If
'If the selected object is a link, then make the link menu visible

If SelVertex > 0 Then
   mnusel.Visible = True
Else
   mnusel.Visible = False
End If
'If the selected object is a vertex, then make the vertex menu visible


End Sub

Sub File_Parse(msg As Variant)

'This subroutine is used by the File_Read subroutine.

Dim A As String
Dim Buffer, VX, VY, VR, VC, VH, VU, VP
Dim id1, id2, Leng, Tens, pushr, pstren, pspan, ll, timing, phase

A = Left$(msg, 1)
If UCase(A) = "G" Then Gravity = Mid$(msg, 2)
If UCase(A) = "A" Then Atmosphere = Mid$(msg, 2)
If UCase(A) = "F" Then WallFriction = Mid$(msg, 2)
If UCase(A) = "B" Then WallBounce = Mid$(msg, 2)
If UCase(A) = "W" Then LeftWind = Mid$(msg, 2)
If UCase(A) = "T" Then Tension = Mid$(msg, 2)
If UCase(A) = "C" Then ClockSpeed = Mid$(msg, 2)
If UCase(A) = "M" Then mode = Mid$(msg, 2)

If mode = 0 Then BTNpicCon_Click
If mode = 1 Then BTNpicSim_Click

INDGrav = Gravity * 2

If UCase(A) = "V" Then
   VX = 0: VY = 0: VH = 0: VU = 0
   For i = 2 To Len(msg)
     A = Mid$(msg, i, 1)
     If A = "|" Then
        If Left$(Buffer, 1) = "X" Then VX = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "Y" Then VY = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "H" Then VH = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "U" Then VU = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "C" Then VC = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "R" Then VR = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "P" Then VP = Mid(Buffer, 2)
        Buffer = ""
     Else
        Buffer = Buffer & A
     End If
   Next i
   vari = AddVertex(VX, VY, VH, VU, VR, VC, VP)
End If

If UCase(A) = "L" Then
   phase = "blank"
   id1 = 0: id2 = 0: Leng = 0: Tens = 0
   For i = 2 To Len(msg)
     A = Mid$(msg, i, 1)
     If A = "|" Then
        If Left$(Buffer, 1) = "A" Then id1 = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "B" Then id2 = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "L" Then Leng = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "T" Then Tens = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "S" Then pspan = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "P" Then pushr = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "N" Then pstren = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "E" Then ll = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "M" Then timing = Mid(Buffer, 2)
        If Left$(Buffer, 1) = "P" Then phase = Mid$(Buffer, 2)
        Buffer = ""
     Else
        Buffer = Buffer & A
     End If
   Next i
   vari = AddLink2(id1, id2, Leng, Tens, pspan, pushr, pstren, ll, timing, phase)
End If

End Sub

Function Cycle_Physics() As Boolean

'Lots of Variables--------------------------------------------------
Dim xer, yer, newx, newy As Single
Dim Leng As Single
Dim Leng2Go_x As Single
Dim Leng2Go_y As Single
Dim LengthTotal As Single
Dim TempTime As Single
Dim Fric
'-------------------------------------------------------------------


'Advance the Muscle Cycle Clock if its not paused-------------------
If ClockPause = False Then CycleTime = CycleTime + ClockSpeed
If CycleTime > CycleSize Then CycleTime = CycleTime - CycleSize
If CycleTime < 0 Then CycleTime = CycleTime + CycleSize
'-------------------------------------------------------------------


'Calculate Link Movement -------------------------------------------
For i = 1 To MaxLinks
  With Link(i)
  If .used = True Then 'only bother with used links
  
  
     .Push = 0
     If CycleTime > .pushtiming - .pushspan And CycleTime < .pushtiming + .pushspan Then
        .Push = (.pushstrength * (1 - (Abs(.pushtiming - CycleTime) / (.pushspan))))
        .Push = (.Push / 30) * .linklength
     End If
     If .pushtiming + .pushspan > CycleSize And CycleTime < .pushtiming + .pushspan - CycleSize Then
        TempTime = .pushtiming - CycleSize
        .Push = (.pushstrength * (1 - ((Abs(TempTime - CycleTime) / (.pushspan)))))
        .Push = (.Push / 30) * .linklength
     End If
     If .pushtiming - .pushspan < 0 And CycleTime > .pushtiming - .pushspan + CycleSize Then
        TempTime = .pushtiming + CycleSize
        .Push = (.pushstrength * (1 - ((Abs(TempTime - CycleTime) / (.pushspan)))))
        .Push = (.Push / 30) * .linklength
     End If
     'This stuff calculates how far the link should increase or decrease
     'its length, in order to simulate 'pushing' or 'pulling'.
     'A positive value for .push indicates that the link will push
     'A negative value indicates that it will pull
     '  Note: pulling is more like a real muscle ;) some other 2d robot
     '        programs whose names I will not mention only support pushing
     
     
     LengthTotal = .linklength
     If ClockPause = False Then LengthTotal = LengthTotal + .Push
     'The link's apparent length will temporarily be changed according to
     '.push
     
     
     T1 = .target1_id: T2 = .target2_id
     xer = (vertex(T2).X + vertex(T2).momentum_x) - (vertex(T1).X + vertex(T1).momentum_x)
     yer = (vertex(T2).y + vertex(T2).momentum_y) - (vertex(T1).y + vertex(T1).momentum_y)
     Leng = Sqr(Abs(xer ^ 2 + yer ^ 2))
     'This will calculate the links 'true' length.  That is the distance
     'between its 2 vertices.   The length stored in .linklength is
     'the length that the link 'should' be.  The link will push or pull
     'to bring the vertices back to that distance.
          
     Leng2Go_x = ((Leng - LengthTotal) / Leng) * xer
     Leng2Go_y = ((Leng - LengthTotal) / Leng) * yer
     Leng2Go_x = Leng2Go_x
     Leng2Go_y = Leng2Go_y
     'Calculate how far in each direction the vertices must go
     'in order to get the link back to its regular length
       
     vertex(T1).momentum_x = vertex(T1).momentum_x + (Leng2Go_x / 2) * .linktension
     vertex(T1).momentum_y = vertex(T1).momentum_y + (Leng2Go_y / 2) * .linktension
     vertex(T2).momentum_x = vertex(T2).momentum_x + (Leng2Go_x / 2) * -1 * .linktension
     vertex(T2).momentum_y = vertex(T2).momentum_y + (Leng2Go_y / 2) * -1 * .linktension
     'These lines actually add the neccessary momentum to the Link's
     'two vertices to make them snap into place.
     'It factors in the amount of tension the link has
     'If the link has a .linktension value of 1, it will snap back
     'into place almost instantly.
     'If it has a value of 0, it will not try to re-establish its
     'length.
       
     
  End If
  End With
Next i
'-------------------------------------------------------------------


'Calculate Vertex Momentum and Implement Movement ------------------
For i = 1 To MaxVertices
   With vertex(i)
   If .used = True Then


     If .y > 0.1 Then .momentum_y = .momentum_y - (Gravity * 1.5)
     If .justreleased = True Then .momentum_x = 0: .momentum_y = 0: .justreleased = False
         'Gravity: decrease the y momentum by the value of Gravity
         '         each turn
    
     .momentum_x = .momentum_x + (LeftWind / 10)
         'Wind: Increase the x momentum by the value of Leftwind.
         '      To make wind that blows from the right, make
         '      Lefwind negative.

     .momentum_x = .momentum_x * (1 - Atmosphere)
     .momentum_y = .momentum_y * (1 - Atmosphere)
        'Slow down the vertices based on how much air resistance
        'there is.
                
     If DragDot = i And SubMode = 2 Then .momentum_x = 0: .momentum_y = 0
        'Put the vertex that's being dragged to the mouse location
               
     .LastX = .X
     .Lasty = .y
     .X = .X + .momentum_x
     .y = .y + .momentum_y
        
     newx = .X + .momentum_x
     newy = .y + .momentum_y
        
        'Now actually make the vertices move, base on how much
        'momentum they have.  Everything up till this section
        'has changed the momentum in some way, even the code
        'to change the links.  Now we simply add the value of the
        'momentum to the position of the vertex.
                     
        'If the vertex goes thru the wall, floor or ceiling,
        'it will be corrected below:

     Fric = WallFriction
     If .Radius < 0 Then .Radius = Abs(.Radius)
     If .Radius > 0 Then .wheel = True
     If .wheel = True Then Fric = 0
     If .y - .Radius < 0.1 Then 'floor
            .y = 0 + .Radius
            .momentum_x = .momentum_x * (1 - Fric)
            .momentum_y = (.momentum_y * WallBounce) * -1
            If .wheel = True Then .momentum_c = -1 * .momentum_x
     End If
     If .X - .Radius < 0.1 Then 'left wall
            .X = 0 + .Radius
            .momentum_x = (.momentum_x * WallBounce) * -1
            .momentum_y = .momentum_y * (1 - Fric)
            If mnuAutoRev.Checked Then
                If AutoReverseCycle = 0 Then AutoReverseCycle = 2: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
                If AutoReverseCycle = 1 Then AutoReverseCycle = 2: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
            End If
            If .wheel = True Then .momentum_c = .momentum_y
     End If
     If .X + .Radius > (RightWall - 0.1) Then 'right wall
            .X = RightWall - .Radius
            .momentum_x = (.momentum_x * WallBounce) * -1
            .momentum_y = .momentum_y * (1 - Fric)
            If mnuAutoRev.Checked Then
                If AutoReverseCycle = 0 Then AutoReverseCycle = 1: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
                If AutoReverseCycle = 2 Then AutoReverseCycle = 1: ClockSpeed = ClockSpeed * -1: INDspd = "Spd: " & ClockSpeed
            End If
            If .wheel = True Then .momentum_c = -1 * .momentum_y
     End If
     If .y + .Radius > (Ceiling - 0.1) And CHKTop.Value = 1 Then 'ceiling
            .y = Ceiling - .Radius
            .momentum_y = (.momentum_y * WallBounce) * -1
            .momentum_x = .momentum_x * (1 - Fric)
            If .wheel = True Then .momentum_c = .momentum_x
     End If
     
                
        'Wow that was big.
        'Basically, if the vertex goes thru a floor or wall, it is
        'stopped, put directly on the surface, and its momentum is
        'reflected and reduced by the WallBounce variable.
        'If WallBounce = 1 then all of the vertex's momentum will
        'be reflected.
        'If WallBounce = 0 then the vertex will be stripped of its
        'momentum in that direction. (the object will not bounce at all)
        
        'Also this section calculates the momentum added to the
        'vertex's wheel (if there is a wheel) as it rubs against
        'the floor or ceiling.  Wheels ignore wall friction.
        
        'Note: as a side effect of link correction,  some objects will
        '      bounce even if WallBounce = 0.  Because when it hits
        '      the ground, the momentum of the upper vertex will make
        '      the link(s) compress,  and when they spring back the
        '      upper vertex will retain some of the momentum of the
        '      link correcting, and carry the object upward.
        
      
        .heading = .heading + .momentum_c
        If .heading > 360 Then .heading = .heading - 360
        If .heading < 0 Then .heading = .heading + 360
        'momentum_c is clockwise momentum.  If the wheel has momentum
        'it will turn
        
    End If
   End With
Next i
'-------------------------------------------------------------------



Cycle_Physics = True  'hey, it worked!


End Function


Sub File_Read(infile As String)

'This reads a file into memory

Dim msg As String
msg = " "
Dim Buffer As Variant


Open infile For Binary As 1
  Do While Not EOF(1)
   Get #1, , msg
   If msg = ";" Then
      File_Parse Trim(Buffer)
      Buffer = ""
   Else
      Buffer = Buffer & msg
   End If
  Loop
Close 1


INDspd = "Spd: " & ClockSpeed
ClearMultiSelect

Slider1.Value = LeftWind
Slider1_Click
Slider2.Value = Atmosphere
Slider2_click
Slider3.Value = Gravity * 100
Slider3_Click

End Sub

Sub File_Save(outfile As String)

'This saves all the vertices and links and their properties to a file.
'It also saves all Global variables such as gravity and atmosphere

'I wrote this subroutine and the File_Read Subroutine so that new
'Variables can easily be added to vertexes and links without losing
'backward compatability with older files

'However, I didn't write it to be easily understood by others so
'you should probably leave this stuff alone.

Open outfile For Output As 1
  Print #1, "G" & Gravity & ";";
  Print #1, "A" & Atmosphere & ";";
  Print #1, "F" & WallFriction & ";";
  Print #1, "B" & WallBounce & ";";
  Print #1, "W" & LeftWind & ";";
  Print #1, "T" & Tension & ";";
  Print #1, "C" & ClockSpeed & ";";
  Print #1, "M" & mode & ";";
  
  For i = 1 To MaxVertices
     With vertex(i)
     If .used = True Then
        Print #1, "V";
        Print #1, "X" & .X & "|";
        Print #1, "Y" & .y & "|";
        Print #1, "D" & i & "|";
        Print #1, "H" & .momentum_x & "|";
        Print #1, "U" & .momentum_y & "|";
        Print #1, "C" & .momentum_c & "|";
        Print #1, "R" & .Radius & "|";
        Print #1, "P" & .phase & "|";
        Print #1, ";"; 'terminate
     End If
     End With
  Next i
  
   For i = 1 To MaxLinks
     With Link(i)
     If .used = True Then
        Print #1, "L";
        Print #1, "A" & .target1_id & "|";
        Print #1, "B" & .target2_id & "|";
        Print #1, "L" & .linklength & "|";
        Print #1, "T" & .linktension & "|";
        Print #1, "S" & .pushspan & "|";
        Print #1, "P" & .Push & "|";
        Print #1, "N" & .pushstrength & "|";
        Print #1, "E" & .lastlen & "|";
        Print #1, "M" & .pushtiming & "|";
        Print #1, "P" & .phase & "|";
        Print #1, ";"; 'terminate
     End If
     End With
  Next i
Close 1


End Sub

Sub ToggleSelction(applies)

vertex(applies).Selected = Not vertex(applies).Selected


End Sub


Private Sub antigrav_Click()


End Sub

Private Sub btnAddwheel_Click()

For i = 1 To MaxVertices
  With vertex(i)
    If .Selected = True Then
      AddWheel i, 25
    End If
  End With
Next i

focusdummy.SetFocus

End Sub



Private Sub btnDelete_Click()


'delete whatever is selected, be it a link or vertex
'if a vertex is deleted, all links it is attached to must go as well
'because a link needs 2 vertices to function

focusdummy.SetFocus



    If SelVertex > 0 Then
      For i = 1 To MaxVertices
            If vertex(i).Selected = True Then
            vari = DeleteVertex(i)
            End If
      Next i
    End If

    If SelLink > 0 Then
    Link(SelLink).used = False
   'deleting a link is much simpler than deleting a vertex
   'because a vertex does not need to be attached to a link
   'in order to function
    End If

End Sub


Sub btnNew_Click()

'Delete all links and vertices.
'this does not reset the global variables


Dim msg, Style, Title, Response

msg = "Destroy Current Scene?"
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "New Scene"
     
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then
    CycleTime = 0
    Call BTNpicCon_Click
    DoEvents
    For i = 1 To MaxVertices
      vertex(i).used = False
    Next i
    For i = 1 To MaxLinks
      Link(i).used = False
    Next i
    VertexCount = 0
    LinkCount = 0
    mode = 0
    SubMode = 0
    focusdummy.SetFocus

Else
   focusdummy.SetFocus
   Exit Sub
End If

End Sub

Private Sub BTNGlobals_Click()

focusdummy.SetFocus
Globals.Show

End Sub

Private Sub btnOpen_Click()

Call mnuOpen_Click
focusdummy.SetFocus

End Sub

Private Sub btnPause_Click()


ClockPause = Not ClockPause
If ClockPause = True Then
   btnPause.BackColor = &HE0E0E0
Else
   btnPause.BackColor = &H8000000F
End If



focusdummy.SetFocus


End Sub

Private Sub BTNpicCon_Click()

CurrentPhase = 1

BTNpicCon.Picture = BTNpicCon.DownPicture
BTNpicSim.Picture = BTNpicSim.DisabledPicture

btnAddwheel.Enabled = True
btnDelete.Enabled = True

ModeIND = " Construct Mode    Click to create vertexes and links.  Click and drag to move vertexes.  Right click to cancel a link."




focusdummy.SetFocus 'FocusDummy is the name of an option button
                    'thats off the left side of the form.
                    'Usually when you click a button, it makes
                    'what I think is a very ugly selection rectangle
                    'around the button's caption.
                    'By setting the focus elsewhere, the ugly
                    'rectangle goes away :)
                    'FocusDummy itself goes nowhere and does nothing.


mode = 0  'switch to construct mode



End Sub

Private Sub BTNpicSim_Click()

BTNpicCon.Picture = BTNpicCon.DisabledPicture
BTNpicSim.Picture = BTNpicSim.DownPicture

btnAddwheel.Enabled = False
btnDelete.Enabled = False

ModeIND = " Simulate Mode:    Click and drag to move vertexes."

INDspd = "Spd: " & ClockSpeed



focusdummy.SetFocus
mode = 1


End Sub

Private Sub btnResetLinks_Click()


ModeIND = " Links Reset."

'should be used in construction mode.
'this will reset the length of all links without having to destroy
'and re-create the link

'I realize that the usefullness of this function may not be clear
'at first.  If you make a link in construction mode, and then change
'the position of one of the vertices,  you'll find that when you
'click simulate, the link will pop back out to its original size.
'This will reset that size so it doesn't do that.
'I put this here because it was driving me completely insane.

'Try to make sure your construct is at rest (make gravity=0 and pause
'the muscle cycle) or else bad things may happen such as your links
'may suddenly 'slouch'.
focusdummy.SetFocus

For i = 1 To MaxLinks
   With Link(i)
     T1 = .target1_id
     T2 = .target2_id
     xer = (vertex(T2).X) - (vertex(T1).X)
     yer = (vertex(T2).y) - (vertex(T1).y)
     Leng = Sqr(Abs(xer ^ 2 + yer ^ 2))
     .linklength = Leng
   End With
Next i

End Sub

Private Sub btnRevers_Click()

'reverse the clock, if you make a robot that walks,
'this may make it walk backwards

ClockSpeed = ClockSpeed * -1
INDspd = "Spd: " & ClockSpeed

End Sub

Private Sub btnSave_Click()

Call mnuSaveAs_Click
focusdummy.SetFocus

End Sub

Private Sub Combo1_Click()

If Combo1.Text = "Presets..." Or Combo1.Text = "" Then Exit Sub


    CycleTime = 0
    focusdummy.SetFocus
    Call BTNpicCon_Click
    DoEvents
    For i = 1 To MaxVertices
      vertex(i).used = False
    Next i
    For i = 1 To MaxLinks
      Link(i).used = False
    Next i
    VertexCount = 0
    LinkCount = 0
    mode = 0
    SubMode = 0
    
    
Call File_Read(App.Path & "\presets\" & Combo1.Text)

End Sub



Private Sub cycleview_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)


'this may be hard to understand but I'm not going to explain it
'in length because I doubt anyone will want to tinker with it.

'Basically when a link is selected, you can click on the cycle box
'with the left button to make a dot appear.  The higher above the
'line the dot, the farther the link will push when the cycle hits it.
'The lower the dot below the line, the more the link will contract.

'By right clicking or dragging to the side, you can alter how
'gradually the link will reach its fully enlarged (or contracted)
'state.

If SubMode = 4 Then
  With Link(SelLink)
If Button = 1 Then
   .pushtiming = X - 2
   Do While .pushtiming > CycleSize Or .pushtiming < 0
   If .pushtiming > CycleSize Then .pushtiming = .pushtiming - CycleSize
   If .pushtiming < 0 Then .pushtiming = .pushtiming + CycleSize
   Loop
   .pushstrength = (0 - y) + (CycleBuffer.ScaleHeight / 2)
   
End If

If Button = 2 Then
   .pushspan = Abs(.pushtiming - (X - 2))
   If .pushspan > (CycleSize / 2) Then .pushspan = (CycleSize / 2)
End If
End With


End If




End Sub


Private Sub cycleview_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

'again, easy way to make dragging possible
Call cycleview_MouseDown(Button, Shift, X, y)


End Sub


Private Sub Form_Load()



FS_Mode = False

ModeIND = " Construct Mode    Click to create vertexes and links.  Click and drag to move vertexes.  Right click to cancel a link."
Combo1.AddItem ("MuscleDemo.botz")
Combo1.AddItem ("Walker.botz")
Combo1.AddItem ("AntiGrav.botz")
Combo1.AddItem ("Unicycle.botz")
Combo1.AddItem ("Dancer.botz")
Combo1.AddItem ("Spike-ball.botz")
Combo1.AddItem ("Jumper.botz")

'Set Variables------------------------------------------------------
DrawColor = RGB(0, 0, 0)
BGColor = RGB(255, 255, 255)
CurrentPhase = 1

Gravity = 0.4
WallBounce = 0.4
Atmosphere = 0.01
LeftWind = 0
Tension = 0.9
WallFriction = 0.7
ClockSpeed = 3

    'Calibrate the dimensions of the playing field
    'This way you can change the playing field just
    'by resizing Main.
    
BoardX = Main.Width
BoardY = Main.Height
RightWall = BoardX - 3   'these are adjusted to keep the vertices
Ceiling = BoardY - 6     'from seeming to sink under the floor
Buffer.Width = Main.Width
Buffer.Height = Main.Height
    'Everything is drawn on the buffer and then blitted to the main
    'picturebox.  That way theres no flicker because Main never has
    'to be cleared.  This idea is inspired by DirectX.
'-------------------------------------------------------------------




'Set up Visuals-----------------------------------------------------
Me.Show   'make sure the form is shown
DoEvents  'let windows do what it needs to
mode = 0  'the default mode is construction mode.

Buffer.BackColor = BGColor


'-------------------------------------------------------------------

File_Read (App.Path & "\resume.botz")
Slider1.Value = LeftWind
Slider1_Click
Slider2.Value = 100 * Atmosphere
Slider2_click
Slider3.Value = Gravity * 100
Slider3_Click


'Now that everthing is good and set up, we can begin the subroutine
'that governs the operation of the program.   The value in the
'brackets is the speed at which the program cycles. lower is faster.
'15 milleseconds is the fastest my computer can go, play with this
'value if you like


Cycle_Botz (25)

End Sub

Private Sub Form_Unload(Cancel As Integer)

File_Save (App.Path & "\resume.botz")
End

End Sub



Private Sub fs_Click()
 If fs.Checked = False Then
    Me.WindowState = 2
    Main.Height = Form1.ScaleHeight
    ModeIND.Visible = False
    BoardX = Main.Width
    BoardY = Main.Height
    RightWall = BoardX - 3
    Ceiling = BoardY - 6
    Buffer.Width = Main.Width
    Buffer.Height = Main.Height
    fs.Checked = True
    CHKTop.Visible = False
    Me.BorderStyle = 0
    FS_Mode = True
    
    Frame_ControlPanel.Visible = False
Else
    FS_Mode = False
    Me.WindowState = 0
    Main.Height = 341
    ModeIND.Visible = True
    BoardX = Main.Width
    BoardY = Main.Height
    RightWall = BoardX - 3
    Ceiling = BoardY - 6
    Buffer.Width = Main.Width
    Buffer.Height = Main.Height
    fs.Checked = False
    CHKTop.Visible = True
    
    Frame_ControlPanel.Visible = True
End If
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub INDrad_Change()


'I don't think this needs explaining

If INDrad = "" Then Exit Sub
vertex(SelVertex).Radius = INDrad

End Sub

Private Sub Main_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then Call btnDelete_Click

End Sub

Private Sub Main_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)


Dim applies, applieslink, inty As Integer
xer = X
yer = Main.Height - y - 3
applies = 0


If Button = 2 Then
    If HowManySelected > 1 Then
      SubMode = 0
      ClearMultiSelect
      Exit Sub
    Else
      SubMode = 0
    End If
End If

If HoverVertex > 0 Then applies = HoverVertex
If HoverLink > 0 Then applieslink = HoverLink: applies = 0

i = applies
If mode = 1 And Button = 1 Then
    ClearMultiSelect
    SelVertex = i
    vertex(i).Selected = True
    SubModeData = i
    SubMode = 2
    DragDot = i
    SelLink = 0
End If


If mode = 0 Then
If Button = 1 And Shift <> 1 Then
  

  If SubMode = 1 Then
     If applies = 0 And applieslink = 0 Then
        ClearMultiSelect
        inty = AddVertex(xer, yer, 0, 0, 0, 0, CurrentPhase)
        vari = AddLink(inty, SubModeData)
         SubModeData = inty
        Exit Sub
     Else
        If applieslink = 0 Then
        inty = applies
        vari = AddLink(inty, SubModeData)
        SubModeData = inty
        SelVertex = applies
        vertex(applies).Selected = True
        Exit Sub
        End If
     End If
  End If
  
  If SubMode = 0 And applies = 0 And applieslink = 0 Then
     ClearMultiSelect
     inty = AddVertex(xer, yer, 0, 0, 0, 0, CurrentPhase)
     SubModeData = inty
     SubMode = 1
     Exit Sub
  End If
  
    If applies > 0 Then
        If vertex(applies).Selected = False Then
            ClearMultiSelect
            SelVertex = i
            vertex(i).Selected = True
            SubModeData = i
            SubMode = 2
            DragDot = i
            SelLink = 0
        Else
            SelVertex = i
            vertex(i).Selected = True
            SubModeData = i
            SubMode = 2
            DragDot = i
            SelLink = 0
        End If
    End If
End If


End If





i = applies
If Button = 1 And Shift = 1 And mode = 0 Then
   'multi select
   If applies > 0 Then
          ToggleSelction applies
   End If
   Exit Sub
Else
 If Button = 1 And mode = 0 Then
'if nothing else, then
    If applies > 0 Then
        SelVertex = i
        vertex(i).Selected = True
        SubModeData = i
        SubMode = 2
        DragDot = i
        SelLink = 0
    End If
    If applieslink > 0 Then
       SelLink = applieslink
       SubMode = 4
       ClearMultiSelect
    End If
  End If
End If





         
End Sub


Private Sub Main_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

MouseX = X
MouseY = y

xer = X
yer = Main.Height - y - 3

  If SubMode = 2 Then
     With vertex(SubModeData)
     .X = xer
     .y = yer
     End With
  End If
     
For i = 1 To MaxVertices
   With vertex(i)
   If .used = True Then
      If xer > (.X - 6) And xer < (.X + 6) Then
         If yer > (.y - 6) And yer < (.y + 6) Then
           'this vertex is meant to be hovered over

           HoverVertex = i
           HoverLink = 0
           Exit Sub
           End If
      End If
   End If
   End With
Next i

For i = 1 To MaxLinks
   With Link(i)
   If .used = True Then
      If xer > (.midx - 6) And xer < (.midx + 6) Then
         If yer > (.midy - 6) And yer < (.midy + 6) Then
           'this vertex is meant to be hovered over

           HoverVertex = 0
           HoverLink = i
           Exit Sub
           End If
      End If
   End If
   End With
Next i


HoverVertex = 0
HoverLink = 0

End Sub


Private Sub Main_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)



Dim HowMany As Integer
HowMany = HowManySelected
If Button = 1 Then
    If mode = 0 And SubMode = 2 And HowMany = 1 Then SubMode = 1
    If mode = 0 And SubMode = 2 And HowMany > 1 Then SubMode = 0
    If mode = 1 And SubMode = 2 Then SubMode = 0: vertex(SubModeData).justreleased = True
End If

DragDot = 0


End Sub

Private Sub mnuAutoRev_Click()


mnuAutoRev.Checked = Not mnuAutoRev.Checked

End Sub

Private Sub mnuConstruct_Click()

Call BTNpicCon_Click


End Sub

Private Sub mnuDeleteMuscle_Click()

Link(SelLink).pushstrength = 0


End Sub

Private Sub mnuglobal_Click()

Globals.Show


End Sub

Private Sub mnuhandl_Click()
mnuhandl.Checked = Not mnuhandl.Checked

End Sub

Private Sub mnuLink_Click()

mnuLink.Checked = Not mnuLink.Checked

End Sub

Private Sub mnuLinkDelete_Click()


'delete whatever is selected, be it a link or vertex
'if a vertex is deleted, all links it is attached to must go as well
'because a link needs 2 vertices to function

focusdummy.SetFocus



    If SelVertex > 0 Then
      For i = 1 To MaxVertices
            If vertex(i).Selected = True Then
            vari = DeleteVertex(i)
            End If
      Next i
    End If

    If SelLink > 0 Then
    Link(SelLink).used = False
   'deleting a link is much simpler than deleting a vertex
   'because a vertex does not need to be attached to a link
   'in order to function
   End If

End Sub

Private Sub mnuLinkLengthReset_Click()

   With Link(SelLink)
     T1 = .target1_id
     T2 = .target2_id
     xer = (vertex(T2).X) - (vertex(T1).X)
     yer = (vertex(T2).y) - (vertex(T1).y)
     Leng = Sqr(Abs(xer ^ 2 + yer ^ 2))
     .linklength = Leng
   End With

End Sub

Private Sub mnuNew_Click()
Call btnNew_Click
End Sub

Private Sub mnuOpen_Click()


Cmd1.InitDir = App.Path
Cmd1.DialogTitle = "Open Bot File"
Cmd1.Filter = "Botz Files|*.botz|All Files|*.*"
Cmd1.Action = 1
Dim msg, Style, Title, Response
msg = "Destroy Current Scene?"
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "New Scene"

If Cmd1.FileName <> "" Then
  If VertexCount > 0 Then
  
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then
    CycleTime = 0
    focusdummy.SetFocus
    Call BTNpicCon_Click
    DoEvents
    For i = 1 To MaxVertices
      vertex(i).used = False
    Next i
    For i = 1 To MaxLinks
      Link(i).used = False
    Next i
    VertexCount = 0
    LinkCount = 0
    mode = 0
    SubMode = 0
Else
   Exit Sub
End If

  End If
  Call File_Read(Cmd1.FileName)
End If

End Sub

Private Sub mnuSave_Click()
Call mnuSaveAs_Click
End Sub

Private Sub mnuSaveAs_Click()

Cmd1.InitDir = App.Path
Cmd1.Filter = "Botz Files|*.botz|All Files|*.*"
Cmd1.DialogTitle = "Save to Bot File"
Cmd1.ShowSave

If Cmd1.FileName = "" Then Exit Sub

File_Save (Cmd1.FileName)
  



End Sub


Private Sub mnuSimulate_Click()

Call BTNpicSim_Click


End Sub

Private Sub mnuVert_Click()

mnuVert.Checked = Not mnuVert.Checked


End Sub

Private Sub mnuVertexAddwheel_Click()


For i = 1 To MaxVertices
  With vertex(i)
    If .Selected = True Then
      AddWheel i, 25
    End If
  End With
Next i

End Sub

Private Sub mnuVertexDelete_Click()


'delete whatever is selected, be it a link or vertex
'if a vertex is deleted, all links it is attached to must go as well
'because a link needs 2 vertices to function

focusdummy.SetFocus



    If SelVertex > 0 Then
      For i = 1 To MaxVertices
            If vertex(i).Selected = True Then
            vari = DeleteVertex(i)
            End If
      Next i
    End If

    If SelLink > 0 Then
    Link(SelLink).used = False
   'deleting a link is much simpler than deleting a vertex
   'because a vertex does not need to be attached to a link
   'in order to function
   End If

End Sub

Private Sub mnuVertexDeleteWheel_Click()


For i = 1 To MaxVertices
  With vertex(i)
    If .Selected = True Then
      .wheel = False
      .Radius = False
    End If
  End With
Next i

End Sub

Private Sub muscledemo_Click()

End Sub


Public Sub Slider1_Click()
Slider1.Text = "Wind " & Slider1.Value
Let LeftWind = Slider1.Value
Label6.Caption = "Wind " & Slider1.Value

End Sub

Private Sub Slider1_Scroll()
Slider1_Click
End Sub

Public Sub Slider2_click()
Slider2.Text = "Air Friction " & Slider2.Value & "%"
Atmosphere = Slider2.Value / 100
Label7.Caption = "Air Friction " & Slider2.Value & "%"
End Sub

Private Sub Slider2_Scroll()
Slider2_click
End Sub

Public Sub Slider3_Click()
Slider3.Text = "Gravity"
Gravity = Slider3.Value / 100
Label8.Caption = "Gravity " & Gravity * 2
End Sub

Private Sub Slider3_Scroll()
Slider3_Click
End Sub

Private Sub txtX_Change()


'I don't think this needs explaining

vertex(SelVertex).X = txtX


End Sub

Private Sub txtY_Change()

'I don't think this needs explaining
vertex(SelVertex).y = txtY

End Sub

Private Sub unicycle_Click()

End Sub

Private Sub UpDown1_DownClick()

'increase clock speed
ClockSpeed = ClockSpeed - 1
INDspd = "Spd: " & ClockSpeed

End Sub

Private Sub UpDown1_UpClick()


'decrease clock speed
ClockSpeed = ClockSpeed + 1
INDspd = "Spd: " & ClockSpeed

'if the clock is negative, you'll have to click down to make it
'go faster, not up.


End Sub


Private Sub walker_Click()


End Sub
