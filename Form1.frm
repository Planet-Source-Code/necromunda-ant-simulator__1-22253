VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   476
   ScaleMode       =   0  'User
   ScaleWidth      =   460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   5520
      Width           =   1335
   End
   Begin VB.PictureBox imgApple 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6120
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   17
      Top             =   5520
      Width           =   270
   End
   Begin VB.PictureBox imgBlack 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6480
      Picture         =   "Form1.frx":09DC
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   16
      Top             =   5520
      Width           =   270
   End
   Begin VB.PictureBox imgVert 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6480
      Picture         =   "Form1.frx":10AE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   15
      Top             =   6000
      Width           =   270
   End
   Begin VB.PictureBox imgHorz 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6360
      Picture         =   "Form1.frx":1170
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   14
      Top             =   6840
      Width           =   450
   End
   Begin VB.PictureBox imgAntDeadVert 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   6120
      Picture         =   "Form1.frx":1202
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   13
      Top             =   6000
      Width           =   270
   End
   Begin VB.PictureBox imgAntDeadHorz 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   6360
      Picture         =   "Form1.frx":18D4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   12
      Top             =   6480
      Width           =   450
   End
   Begin VB.PictureBox imgAntRight 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   6000
      Picture         =   "Form1.frx":1F8E
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   11
      Top             =   6960
      Width           =   195
   End
   Begin VB.PictureBox imgAntLeft 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   5640
      Picture         =   "Form1.frx":1FF8
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   10
      Top             =   6960
      Width           =   195
   End
   Begin VB.PictureBox imgAntUp 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   6000
      Picture         =   "Form1.frx":2062
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   9
      Top             =   6600
      Width           =   120
   End
   Begin VB.PictureBox imgAntDown 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5760
      Picture         =   "Form1.frx":20E0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   8
      Top             =   6600
      Width           =   120
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "&KILL ANT!!"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdApples 
      Caption         =   "&AUTO FEED"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Timer Apples 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1920
      Top             =   6000
   End
   Begin VB.ListBox lstReport 
      Height          =   2010
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&PAUSE"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Height          =   135
      Left            =   6720
      TabIndex        =   2
      Top             =   5040
      Width           =   135
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&START"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   1440
      Top             =   6000
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   4935
      Left            =   0
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label lblReport 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   6240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'This is used for the randomly selected point
'that each ant goes to. i originally used random movement
'but it just made them dance about
Private Type pointXY
    X As Integer
    Y As Integer
End Type

Private Type ant
    Name As String
    X As Integer    'Coordinates
    Y As Integer
    Direct As Integer 'Which way its facing
    Destination As pointXY 'Where its going
    DestType As Integer 'What it's going to
    ApplesEaten As Integer 'How much its eaten
    IsDead As Boolean 'Is it dead?
End Type

Const DestApple = 0     'Constants for destination type
Const DestGotoPoint = 1

Const antNum = 50    'Changing this changes the number of ants/apples
Const sleeptime = 0

Dim ant(antNum) As ant
Dim pause As Boolean
Dim random As Integer

Dim Limit As Integer
Dim KilledCount As Integer

Dim loopval As Integer

Const goUp = 1      'constants for each direction
Const goDown = 2
Const goLeft = 3
Const goRight = 4

Dim names(10) As String

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long) 'Sleeeeeep
'used to slow it down when there are few ants


Private Sub cmdReset_Click()

    Call cmdClear_Click
    Call loadForm
        
    cmdSearch.Enabled = True
    cmdSurvive.Enabled = True
       
End Sub

Private Sub Form_Load()
    
    Call Randomize
  
    'Array storing the possible names of the ants
    names(1) = "Bob"
    names(2) = "Jim"
    names(3) = "Dan"
    names(4) = "Tim"
    names(5) = "Baz"
    names(6) = "Reg"
    names(7) = "Tom"
    names(8) = "Rob"
    names(9) = "Joe"
    names(10) = "Marmaduke"
        
    Limit = 0
    
    Me.Show
    DoEvents
    
    Call loadForm
       
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub


Private Sub cmdApples_Click()

    If Apples.Enabled = False Then Apples.Enabled = True Else Apples.Enabled = False

End Sub

Private Sub cmdClear_Click()

    Picture1.Cls

End Sub

Private Sub cmdKill_Click()
    Dim antKill As Integer
   
    DoEvents
    
    If lstReport.ListIndex = -1 Then
kill:
        DoEvents
        antKill = Int(Rnd() * antNum) + 1
        
        If ant(antKill).IsDead = False Then
            Call kill_ant(antKill)
        
        Else:
            GoTo kill
        
        End If
        
        
    Else
        antKill = lstReport.ListIndex + 1
        
        If ant(antKill).IsDead = False Then
            Call kill_ant(antKill)
        
        Else
            lblReport.Caption = UCase(Mid$(ant(antKill).Name, 3, 3) & " IS ALREADY DEAD!!")
            Timer.Enabled = True
            Exit Sub
        
        End If
        
    End If
    
End Sub

Private Sub cmdPause_Click()
        
    If Timer.Enabled = True Then Timer.Enabled = False
    If Apples.Enabled = True Then Apples.Enabled = False
    
    lblReport.ForeColor = vbRed
    If pause = True Then
        Beep
        lblReport.Caption = ""
        pause = False
    
    Else
        Beep
        pause = True
        lblReport.Caption = "PAUSED"
    
    End If

End Sub

Private Sub cmdSearch_Click()
    
    Picture1.Enabled = True
    cmdSearch.Enabled = False
    
    'The main loop for making the ants run about
    Do
        DoEvents
    
        For loopval = 1 To antNum
    
                Call gotoPoint(ant(loopval))
                                
        Next loopval
    
    Loop
    
End Sub




Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim AppleX As Integer
    Dim AppleY As Integer
    Dim IsAnt As Boolean
    
        IsAnt = False
        
        
        'Adjust the click co-ords to make the apple appear under the
        'cursors hotspot
        If 0 < X < 456 Then AppleX = X - 10 Else GoTo endstop
        If 0 < Y < 325 Then AppleY = Y - 15 Else GoTo endstop
            
        
        'Clear the way then draw an apple
        BitBlt Picture1.hDC, AppleX, AppleY, imgBlack.Width, imgBlack.Height, imgBlack.hDC, 0, 0, vbMergePaint
        BitBlt Picture1.hDC, AppleX, AppleY, imgApple.Width, imgBlack.Height, imgApple.hDC, 0, 0, vbSrcAnd
        
        'This part randomly sets ants to eat the apple if they aren't already
        For loopval = 1 To antNum
        
        random = Rnd() * 1
        If random = 1 Then
            If ant(loopval).DestType <> DestApple And ant(loopval).IsDead = False Then
                With ant(loopval)
                    .Destination.X = AppleX
                    .Destination.Y = AppleY
                    .DestType = DestApple
                End With
                
                IsAnt = True
            End If
        End If
    
        Next loopval

endstop:
    
    'If no ants have accepted the apple then blank it out and display
    'a message
    
    If IsAnt = False Then
        DoEvents
        BitBlt Picture1.hDC, AppleX, AppleY, imgBlack.Width, imgBlack.Height, imgBlack.hDC, 0, 0, vbMergePaint
        lblReport.ForeColor = vbRed
        lblReport.Caption = "Bad Apple! REJECTED!!!"
        Timer.Enabled = True
    End If


End Sub

'@@@@@@@@@@@@@@
'@PRIVATE SUBS@
'@@@@@@@@@@@@@@

Private Sub loadForm()
    'Setup the ants
        
    For loopval = 1 To antNum
        'Firstly, setup a Goto point
        ant(loopval).X = Int(Rnd() * 453)
        ant(loopval).Y = Int(Rnd() * 325)
        
        'Assign a name
        ant(loopval).Name = names(Int(Rnd() * 10) + 1)
        ant(loopval).ApplesEaten = 0
                    
        'Create the random point for the ant to go to
        Call createPoint(ant(loopval))
            
        'Then Draw the ant, and setup the graphical variables
        BitBlt Picture1.hDC, ant(loopval).X, ant(loopval).Y, imgAntUp.Width, imgAntUp.Height, imgAntUp.hDC, 0, 0, vbSrcAnd
        ant(loopval).Direct = 1
        ant(loopval).DestType = DestGotoPoint
    Next loopval
    
    'Fill the scoreboard listbox
    lstReport.Clear
    For loopval = 1 To antNum
        lstReport.AddItem ant(loopval).Name & Space(2) & ant(loopval).ApplesEaten
    
    Next loopval

End Sub


Private Sub kill_ant(number As Integer)
    
    'First set the isdead bool variable to true
    ant(number).IsDead = True
            
    'Determine the direction it's facing (horizontal or vertical) and
    'draw the dead image.
                
    If ant(number).Direct = 3 Or ant(number).Direct = 4 Then
        BitBlt Picture1.hDC, ant(number).X - 5, ant(number).Y - 5, imgHorz.Width, imgHorz.Height, imgHorz.hDC, 0, 0, vbMergePaint
        BitBlt Picture1.hDC, ant(number).X - 5, ant(number).Y - 5, imgAntDeadHorz.Width, imgAntDeadHorz.Height, imgAntDeadHorz.hDC, 0, 0, vbSrcAnd
       
    ElseIf ant(number).Direct = 1 Or ant(number).Direct = 2 Then
        BitBlt Picture1.hDC, ant(number).X - 5, ant(number).Y - 5, imgVert.Width, imgVert.Height, imgVert.hDC, 0, 0, vbMergePaint
        BitBlt Picture1.hDC, ant(number).X - 5, ant(number).Y - 5, imgAntDeadVert.Width, imgAntDeadVert.Height, imgAntDeadVert.hDC, 0, 0, vbSrcAnd
    End If
        
    'Display dead message
    lblReport.ForeColor = vbRed
    lblReport.Caption = ant(number).Name & " KILLED!!"
    ant(number).Name = "xx" & ant(number).Name & "xx"
    
    Timer.Enabled = True

End Sub
 
 
Private Sub gotoPoint(ant As ant)
'This is for searching out the random point
    
    'Pause:
    'If pause has been pressed then loop until it's not
    
    If pause = True Then
        Do
            DoEvents
        Loop Until pause = False

    End If
    
    'If its got there and it's a goto point then create a new one
    If ant.X = ant.Destination.X And ant.Y = ant.Destination.Y And ant.DestType = DestGotoPoint Then
        Call createPoint(ant)
                   
    'If it gets there and it's an apple point then eat it.
    ElseIf ant.X = ant.Destination.X And ant.Y = ant.Destination.Y And ant.DestType = DestApple Then
        Call appleEat(ant)
        
    End If

    'Makes the movement seem more natural
    'it randomly chooses between going up/down or left/right
    
    If Int(Rnd() * 2) = 1 Then
        If ant.Y < ant.Destination.Y Then
            Call moveAnt(ant, goUp)
        ElseIf ant.Y > ant.Destination.Y Then
            Call moveAnt(ant, goDown)
        End If
        
    Else
        If ant.X < ant.Destination.X Then
            Call moveAnt(ant, goRight)
        ElseIf ant.X > ant.Destination.X Then
            Call moveAnt(ant, goLeft)
        End If
 
    End If
        
End Sub

Private Sub createPoint(ant As ant)
       
    'Create the X and Y co-ords for the ant to travel to
    
    ant.Destination.X = Int(Rnd() * 453)
    ant.Destination.Y = Int(Rnd() * 325)
    
    'Set the destination type as a goto point
    ant.DestType = DestGotoPoint
    
End Sub


Private Sub moveAnt(ant As ant, direction As Integer)
    
    Sleep (sleeptime)
    'This slows down the process
    
    'Don't move the ant if it's dead.
    If ant.IsDead = True Then Exit Sub
    
    'This section demonstrates the use of BitBlt...(kinda)
    
    'First it goes through this bit below, which draws an ant
    'but reversed.  It draws only the black bits, but as white.
    'This clears the ant image out of the way before drawing the new
    'one.
    
    
    'BitBlt: A Rough Guide
    'hDestDC = the Hdc (eg. picture1.hdc) of the picture box where
    'you want to copy the image
    
    'X = Where to draw in X
    'Y = Where to draw in Y
    
    'nWidth = The width of the image being draw
    'nHeight = The height
    
    'hSrcDC = The hdc of the source picture box
    
    'xSrc = where on the source pic box X
    'ySrc = where on the source pic box Y
    
    'dwRop = how to copy the image
    
        'vbMergePaint = copy only black bits (as white)
        'vbSrcAnd = copy non white bits
        'vbNotSrcAnd = Inverse
        'vbSrcCopy = Copies image
    
    '(This wont work with image boxes as they dont have an hDC)
        
    
    Select Case ant.Direct
        Case Is = 1     'Go UP
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntUp.Width, imgAntUp.Height, imgAntUp.hDC, 0, 0, vbMergePaint
        
        Case Is = 2     'Go DOWN
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntDown.Width, imgAntDown.Height, imgAntDown.hDC, 0, 0, vbMergePaint
        
        Case Is = 3     'Go LEFT
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntLeft.Width, imgAntLeft.Height, imgAntLeft.hDC, 0, 0, vbMergePaint
        
        Case Is = 4     'Go RIGHT
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntRight.Width, imgAntRight.Height, imgAntRight.hDC, 0, 0, vbMergePaint
        
    End Select
    
    
    Select Case direction
        Case Is = goUp
            If ant.Y < 325 Then ant.Y = ant.Y + 1 Else ant.Y = ant.Y - 1
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntUp.Width, imgAntUp.Height, imgAntUp.hDC, 0, 0, vbSrcAnd
            ant.Direct = 1
        Case Is = goDown
            If ant.Y > 0 Then ant.Y = ant.Y - 1 Else ant.Y = ant.Y + 1
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntDown.Width, imgAntDown.Height, imgAntDown.hDC, 0, 0, vbSrcAnd
            ant.Direct = 2
        Case Is = goLeft
            If ant.X > 0 Then ant.X = ant.X - 1 Else ant.X = ant.X + 1
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntLeft.Width, imgAntLeft.Height, imgAntLeft.hDC, 0, 0, vbSrcAnd
            ant.Direct = 3
        Case Is = goRight
            If ant.X < 453 Then ant.X = ant.X + 1 Else ant.X = ant.X - 1
            BitBlt Picture1.hDC, ant.X, ant.Y, imgAntRight.Width, imgAntRight.Height, imgAntRight.hDC, 0, 0, vbSrcAnd
            ant.Direct = 4
    End Select
    
    DoEvents
    
End Sub


Private Sub appleEat(AnAnt As ant)
    'The movement to "eat" the apple
            
    Picture1.Enabled = False
        
    'Add to the apple tally
    AnAnt.ApplesEaten = AnAnt.ApplesEaten + 1
    
    lblReport.ForeColor = vbGreen
    lblReport.Caption = "Eaten by " & AnAnt.Name
    'lblReport.ForeColor = vbRed
    Timer.Enabled = True
    
    For loopval = 1 To antNum
        If ant(loopval).DestType = DestApple And _
            ant(loopval).Destination.X = AnAnt.Destination.X And _
            ant(loopval).Destination.Y = AnAnt.Destination.Y Then
            ant(loopval).DestType = DestGotoPoint
            
        End If
    Next loopval
        
    For loopval = 1 To 5
        Sleep (10)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Call moveAnt(AnAnt, goRight)
        Sleep (10)
        Call moveAnt(AnAnt, goUp)
        Call moveAnt(AnAnt, goUp)
        Call moveAnt(AnAnt, goUp)
        Call moveAnt(AnAnt, goUp)
        Call moveAnt(AnAnt, goUp)
        Sleep (10)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
        Call moveAnt(AnAnt, goLeft)
    Next loopval
    
    Picture1.Enabled = True
    
    lstReport.Clear
    
    For loopval = 1 To antNum
        lstReport.AddItem ant(loopval).Name & Space(2) & ant(loopval).ApplesEaten
    
    Next loopval
    
    Call createPoint(AnAnt)
    
    
End Sub

'@@@@@@@
'@TIMER@
'@@@@@@@

Private Sub Timer_Timer()

    lblReport.Caption = ""
    Timer.Enabled = False

End Sub

Private Sub Apples_Timer()

    Call Picture1_MouseDown(1, 0, Int(Rnd() * (440)) + 20, Int(Rnd() * (300)) + 20)

End Sub

