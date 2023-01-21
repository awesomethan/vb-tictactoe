VERSION 5.00
Begin VB.Form frmTrio 
   Caption         =   "Trio"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   587
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   938
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGameBoard 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   8535
      Left            =   0
      ScaleHeight     =   569
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   905
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.CommandButton cmdHome 
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9480
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5760
         MaskColor       =   &H8000000F&
         TabIndex        =   15
         Top             =   4320
         Width           =   2535
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5160
         TabIndex        =   14
         Top             =   3480
         Width           =   3735
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10920
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdStart 
         Appearance      =   0  'Flat
         Caption         =   "Play!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5160
         TabIndex        =   11
         Top             =   2640
         Width           =   3735
      End
      Begin VB.CommandButton cmdPlayAgain 
         Caption         =   "Play Again!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1800
         TabIndex        =   8
         Top             =   2640
         Width           =   3375
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame fraQuestion 
         Caption         =   "Who would you like to go first?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3480
         TabIndex        =   3
         Top             =   2040
         Width           =   6975
         Begin VB.OptionButton optComputer 
            Caption         =   "Computer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3840
            TabIndex        =   5
            Top             =   600
            Width           =   2775
         End
         Begin VB.OptionButton optUser 
            Caption         =   "You"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Label lblFirstInstructions 
         Caption         =   $"frmTrio.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   3960
         TabIndex        =   13
         Top             =   4320
         Width           =   6135
      End
      Begin VB.Label lblCompScore 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Computer: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         TabIndex        =   10
         Top             =   7560
         Width           =   2535
      End
      Begin VB.Label lblUserScore 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7920
         TabIndex        =   9
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Image imgComputer3 
         Height          =   1575
         Left            =   10680
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Image imgComputer2 
         Height          =   1575
         Left            =   10680
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgComputer1 
         Height          =   1575
         Left            =   10680
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Image imgUser3 
         DragMode        =   1  'Automatic
         Height          =   1575
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Image imgUser2 
         DragMode        =   1  'Automatic
         Height          =   1575
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Image imgUser1 
         DragMode        =   1  'Automatic
         Height          =   1575
         Left            =   7920
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Line linSeparator 
         X1              =   664
         X2              =   664
         Y1              =   80
         Y2              =   448
      End
      Begin VB.Label lblComputer 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10440
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Line linHorizontal3 
         X1              =   0
         X2              =   450
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line linHorizontal0 
         X1              =   0
         X2              =   450
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linVertical0 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   450
      End
      Begin VB.Line linVertical3 
         X1              =   450
         X2              =   450
         Y1              =   0
         Y2              =   450
      End
      Begin VB.Line linHorizontal2 
         X1              =   0
         X2              =   450
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line linHorizontal1 
         X1              =   0
         X2              =   450
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Line linVertical2 
         X1              =   300
         X2              =   300
         Y1              =   0
         Y2              =   450
      End
      Begin VB.Line linVertical1 
         X1              =   150
         X2              =   150
         Y1              =   0
         Y2              =   450
      End
      Begin VB.Label lblResults 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Good luck!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   7080
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmTrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Jonathan and Ethan
'Date: January 2019
'Purpose: Trio Game culminating assignment
Option Explicit
'Declare
Dim intUser1 As Integer
Dim intUser2 As Integer
Dim intUser3 As Integer
Dim intComp1 As Integer
Dim intComp2 As Integer
Dim intComp3 As Integer
Dim intUserScore As Integer
Dim intCompScore As Integer
Dim intAvailable As Integer
Public Function ConvertXY(X As Single, Y As Single)

Dim intMagicNum As Integer
intMagicNum = 0
If X > 0 And X < 150 Then
    If Y > 0 And Y < 150 Then          'Magic Square Position 8
         intMagicNum = 8
    ElseIf Y > 150 And Y < 300 Then    'Magic Square Position 3
         intMagicNum = 3
    Else                               'Magic Square Position 4
         intMagicNum = 4
    End If
ElseIf X > 150 And X < 300 Then
    If Y > 0 And Y < 150 Then          'Magic Square Position 1
         intMagicNum = 1
    ElseIf Y > 150 And Y < 300 Then    'Magic Square Position 5
         intMagicNum = 5
    Else                               'Magic Square Position 9
         intMagicNum = 9
    End If
Else
    If Y > 0 And Y < 150 Then          'Magic Square Position 6
         intMagicNum = 6
    ElseIf Y > 150 And Y < 300 Then    'Magic Square Position 7
         intMagicNum = 7
    Else                               'Magic Square Position 2
         intMagicNum = 2
    End If
End If
ConvertXY = intMagicNum
End Function

Private Sub cmdBack_Click()
fraQuestion.Visible = False
cmdStart.Visible = True
cmdOptions.Visible = True
cmdQuit.Visible = True
cmdBack.Visible = False
lblFirstInstructions.Visible = False
End Sub

Private Sub cmdHelp_Click()
MsgBox ("Welcome to Trio! Move your first three game pieces, as you would in Tic-Tac-Toe. If you have not beaten the computer yet, continue to move the pieces currently on the gameboard until you get three in a row - horizontally, vertically, or diagonally!")
End Sub

Private Sub cmdHome_Click()
cmdHome.Visible = False
linVertical0.Visible = False
linVertical1.Visible = False
linVertical2.Visible = False
linVertical3.Visible = False
linHorizontal0.Visible = False
linHorizontal1.Visible = False
linHorizontal2.Visible = False
linHorizontal3.Visible = False
linSeparator.Visible = False
lblResults.Visible = False
lblUser.Visible = False
lblComputer.Visible = False
imgUser1.Visible = False
imgUser2.Visible = False
imgUser3.Visible = False
imgComputer1.Visible = False
imgComputer2.Visible = False
imgComputer3.Visible = False
cmdStart.Visible = True
cmdOptions.Visible = True
cmdQuit.Visible = True
cmdPlayAgain.Visible = False
lblUserScore.Visible = False
lblCompScore.Visible = False
imgUser1.Top = 120
imgUser1.Left = 528
imgUser2.Top = 232
imgUser2.Left = 528
imgUser3.Top = 344
imgUser3.Left = 528
imgComputer1.Top = 120
imgComputer1.Left = 712
imgComputer2.Top = 232
imgComputer2.Left = 712
imgComputer3.Top = 344
imgComputer3.Left = 712
intUser1 = -99
intUser2 = -98
intUser3 = -97
intComp1 = -96
intComp2 = -95
intComp3 = -94
imgUser1.Enabled = True
imgUser2.Enabled = True
imgUser3.Enabled = True
End Sub

Private Sub cmdOptions_Click()
fraQuestion.Visible = True
cmdStart.Visible = False
cmdQuit.Visible = False
cmdBack.Visible = True
cmdOptions.Visible = False
lblFirstInstructions.Visible = True
cmdReset.Visible = False
End Sub

Private Sub cmdPlayAgain_Click()
cmdPlayAgain.Visible = False
cmdHome.Visible = True
imgUser1.Top = 120
imgUser1.Left = 528
imgUser2.Top = 232
imgUser2.Left = 528
imgUser3.Top = 344
imgUser3.Left = 528
imgComputer1.Top = 120
imgComputer1.Left = 712
imgComputer2.Top = 232
imgComputer2.Left = 712
imgComputer3.Top = 344
imgComputer3.Left = 712
intUser1 = -99
intUser2 = -98
intUser3 = -97
intComp1 = -96
intComp2 = -95
intComp3 = -94
lblResults.Caption = "Good luck!"
imgUser1.Enabled = True
imgUser2.Enabled = True
imgUser3.Enabled = True
isUserFirst
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReset_Click()
imgUser1.Top = 120
imgUser1.Left = 528
imgUser2.Top = 232
imgUser2.Left = 528
imgUser3.Top = 344
imgUser3.Left = 528
imgComputer1.Top = 120
imgComputer1.Left = 712
imgComputer2.Top = 232
imgComputer2.Left = 712
imgComputer3.Top = 344
imgComputer3.Left = 712
intUser1 = -99
intUser2 = -98
intUser3 = -97
intComp1 = -96
intComp2 = -95
intComp3 = -94
cmdPlayAgain.Visible = False
intUserScore = 0
lblUserScore.Caption = "You: " & intUserScore
intCompScore = 0
lblCompScore.Caption = "Computer: " & intCompScore
cmdStart.Visible = True
linVertical0.Visible = False
linVertical1.Visible = False
linVertical2.Visible = False
linVertical3.Visible = False
linHorizontal0.Visible = False
linHorizontal1.Visible = False
linHorizontal2.Visible = False
linHorizontal3.Visible = False
linSeparator.Visible = False
lblResults.Visible = False
lblUser.Visible = False
lblComputer.Visible = False
imgUser1.Visible = False
imgUser2.Visible = False
imgUser3.Visible = False
imgComputer1.Visible = False
imgComputer2.Visible = False
imgComputer3.Visible = False
optUser.Value = True
cmdOptions.Visible = True
cmdQuit.Visible = True
imgUser1.Enabled = True
imgUser2.Enabled = True
imgUser3.Enabled = True
cmdHome.Visible = False
lblUserScore.Visible = False
lblCompScore.Visible = False
cmdReset.Visible = False
End Sub
Private Sub cmdStart_Click()
lblResults.Caption = "Good luck!"
cmdStart.Visible = False
linVertical0.Visible = True
linVertical1.Visible = True
linVertical2.Visible = True
linVertical3.Visible = True
linHorizontal0.Visible = True
linHorizontal1.Visible = True
linHorizontal2.Visible = True
linHorizontal3.Visible = True
linSeparator.Visible = True
lblResults.Visible = True
lblUserScore.Visible = True
lblCompScore.Visible = True
lblUser.Visible = True
lblComputer.Visible = True
imgUser1.Visible = True
imgUser2.Visible = True
imgUser3.Visible = True
imgComputer1.Visible = True
imgComputer2.Visible = True
imgComputer3.Visible = True
fraQuestion.Visible = False
lblFirstInstructions.Visible = False
cmdOptions.Visible = False
cmdQuit.Visible = False
cmdHome.Visible = True
cmdReset.Visible = True
isUserFirst
End Sub
Private Sub Form_Activate()
Randomize
imgUser1.Picture = LoadPicture(App.Path & "\apple.jpg")
imgUser2.Picture = LoadPicture(App.Path & "\apple.jpg")
imgUser3.Picture = LoadPicture(App.Path & "\apple.jpg")
imgComputer1.Picture = LoadPicture(App.Path & "\orange.jfif")
imgComputer2.Picture = LoadPicture(App.Path & "\orange.jfif")
imgComputer3.Picture = LoadPicture(App.Path & "\orange.jfif")
picGameBoard.Picture = LoadPicture(App.Path & "\landscape_background.jfif")
intUser1 = -99
intUser2 = -98
intUser3 = -97
intComp1 = -96
intComp2 = -95
intComp3 = -94
intUserScore = 0
intCompScore = 0
intAvailable = 0
cmdPlayAgain.Visible = False
linVertical0.Visible = False
linVertical1.Visible = False
linVertical2.Visible = False
linVertical3.Visible = False
linHorizontal0.Visible = False
linHorizontal1.Visible = False
linHorizontal2.Visible = False
linHorizontal3.Visible = False
linSeparator.Visible = False
lblResults.Visible = False
lblUser.Visible = False
lblComputer.Visible = False
imgUser1.Visible = False
imgUser2.Visible = False
imgUser3.Visible = False
imgComputer1.Visible = False
imgComputer2.Visible = False
imgComputer3.Visible = False
lblUserScore.Visible = False
lblCompScore.Visible = False
fraQuestion.Visible = False
lblFirstInstructions.Visible = False
optUser.Value = True
cmdBack.Visible = False
cmdHome.Visible = False
cmdReset.Visible = False
End Sub

Private Sub optComputer_Click()
cmdStart.Enabled = True
End Sub
Private Sub optUser_Click()
cmdStart.Enabled = True
End Sub
Private Sub picGameBoard_DragDrop(Source As Control, X As Single, Y As Single)
'Declare
Dim intSquare As Integer
Dim intCompTarget As Integer
'Initialize
intSquare = 0
intCompTarget = 0

'Source.Top = Y
'Source.Left = X
If X > 0 And X < 150 Then
    If Y > 0 And Y < 150 Then          'Magic Square Position 8
        Source.Left = 15
        Source.Top = 15
    ElseIf Y > 150 And Y < 300 Then    'Magic Square Position 3
        Source.Left = 15
        Source.Top = 165
    Else                               'Magic Square Position 4
        Source.Left = 15
        Source.Top = 315
    End If
ElseIf X > 150 And X < 300 Then
    If Y > 0 And Y < 150 Then          'Magic Square Position 1
        Source.Left = 165
        Source.Top = 15
    ElseIf Y > 150 And Y < 300 Then    'Magic Square Position 5
        Source.Left = 165
        Source.Top = 165
    Else                               'Magic Square Position 9
        Source.Left = 165
        Source.Top = 315
    End If
Else
    If Y > 0 And Y < 150 Then          'Magic Square Position 6
        Source.Left = 315
        Source.Top = 15
    ElseIf Y > 150 And Y < 300 Then    'Magic Square Position 7
        Source.Left = 315
        Source.Top = 165
    Else                               'Magic Square Position 2
        Source.Left = 315
        Source.Top = 315
    End If
End If

intSquare = ConvertXY(X, Y)

If Source = imgUser1 Then
    intUser1 = intSquare
ElseIf Source = imgUser2 Then
    intUser2 = intSquare
Else
    intUser3 = intSquare
End If

If isGameOver Then
    cmdPlayAgain.Visible = True
    imgUser1.Enabled = False
    imgUser2.Enabled = False
    imgUser3.Enabled = False
Else
    'write the AI here (else meaning if all the comp integers are greater than 0,
    'which means they're all on the gameboard.
    'if they're all on the game board, we can write the AI here.
    MoveComputer
End If
End Sub
Public Sub MoveComputer()

'declare
Dim intCompTarget As Integer
Dim intCompPieceNum As Integer
'initialize
intCompTarget = 0
intCompPieceNum = 0
'process

    'to attack:
    If 1 <= (15 - intComp1 - intComp2) And (15 - intComp1 - intComp2) <= 9 And isSquareAvailable(15 - intComp1 - intComp2) Then
        intCompTarget = 15 - intComp1 - intComp2
        intCompPieceNum = 3
    ElseIf 1 <= (15 - intComp1 - intComp3) And (15 - intComp1 - intComp3) <= 9 And isSquareAvailable(15 - intComp1 - intComp3) Then
        intCompTarget = 15 - intComp1 - intComp3
        intCompPieceNum = 2
    ElseIf 1 <= (15 - intComp2 - intComp3) And (15 - intComp2 - intComp3) <= 9 And isSquareAvailable(15 - intComp2 - intComp3) Then
        intCompTarget = 15 - intComp2 - intComp3
        intCompPieceNum = 1
    'to defend:
    ElseIf 1 <= (15 - intUser1 - intUser2) And (15 - intUser1 - intUser2) <= 9 And isSquareAvailable(15 - intUser1 - intUser2) Then
        intCompTarget = 15 - intUser1 - intUser2
        intCompPieceNum = compPieceNum
    ElseIf 1 <= (15 - intUser1 - intUser3) And (15 - intUser1 - intUser3) <= 9 And isSquareAvailable(15 - intUser1 - intUser3) Then
        intCompTarget = 15 - intUser1 - intUser3
        intCompPieceNum = compPieceNum
    ElseIf 1 <= (15 - intUser2 - intUser3) And (15 - intUser2 - intUser3) <= 9 And isSquareAvailable(15 - intUser2 - intUser3) Then
        intCompTarget = 15 - intUser2 - intUser3
        intCompPieceNum = compPieceNum
    Else
        Do
            intCompTarget = Int(Rnd * 9) + 1
        Loop While isSquareAvailable(intCompTarget) = False
        intCompPieceNum = compPieceNum
    End If
    
    'final action of moving for the computer: the targeted square number, and the piece number
    If intCompPieceNum = 1 Then
        If intCompTarget = 8 Then
            imgComputer1.Left = 15                'Position 8
            imgComputer1.Top = 15
            intComp1 = 8
        ElseIf intCompTarget = 3 Then
            imgComputer1.Left = 15                'Position 3
            imgComputer1.Top = 165
            intComp1 = 3
        ElseIf intCompTarget = 4 Then
            imgComputer1.Left = 15                'Position 4
            imgComputer1.Top = 315
            intComp1 = 4
        ElseIf intCompTarget = 1 Then
            imgComputer1.Left = 165               'Position 1
            imgComputer1.Top = 15
            intComp1 = 1
        ElseIf intCompTarget = 5 Then
            imgComputer1.Left = 165               'Position 5
            imgComputer1.Top = 165
            intComp1 = 5
        ElseIf intCompTarget = 9 Then
            imgComputer1.Left = 165               'Position 9
            imgComputer1.Top = 315
            intComp1 = 9
        ElseIf intCompTarget = 6 Then
            imgComputer1.Left = 315               'Position 6
            imgComputer1.Top = 15
            intComp1 = 6
        ElseIf intCompTarget = 7 Then
            imgComputer1.Left = 315               'Position 7
            imgComputer1.Top = 165
            intComp1 = 7
        ElseIf intCompTarget = 2 Then
            imgComputer1.Left = 315               'Position 2
            imgComputer1.Top = 315
            intComp1 = 2
        End If
    ElseIf intCompPieceNum = 2 Then
        If intCompTarget = 8 Then
            imgComputer2.Left = 15                'Position 8
            imgComputer2.Top = 15
            intComp2 = 8
        ElseIf intCompTarget = 3 Then
            imgComputer2.Left = 15                'Position 3
            imgComputer2.Top = 165
            intComp2 = 3
        ElseIf intCompTarget = 4 Then
            imgComputer2.Left = 15                'Position 4
            imgComputer2.Top = 315
            intComp2 = 4
        ElseIf intCompTarget = 1 Then
            imgComputer2.Left = 165               'Position 1
            imgComputer2.Top = 15
            intComp2 = 1
        ElseIf intCompTarget = 5 Then
            imgComputer2.Left = 165               'Position 5
            imgComputer2.Top = 165
            intComp2 = 5
        ElseIf intCompTarget = 9 Then
            imgComputer2.Left = 165               'Position 9
            imgComputer2.Top = 315
            intComp2 = 9
        ElseIf intCompTarget = 6 Then
            imgComputer2.Left = 315               'Position 6
            imgComputer2.Top = 15
            intComp2 = 6
        ElseIf intCompTarget = 7 Then
            imgComputer2.Left = 315               'Position 7
            imgComputer2.Top = 165
            intComp2 = 7
        ElseIf intCompTarget = 2 Then
            imgComputer2.Left = 315               'Position 2
            imgComputer2.Top = 315
            intComp2 = 2
        End If
    ElseIf intCompPieceNum = 3 Then
        If intCompTarget = 8 Then
            imgComputer3.Left = 15                'Position 8
            imgComputer3.Top = 15
            intComp3 = 8
        ElseIf intCompTarget = 3 Then
            imgComputer3.Left = 15                'Position 3
            imgComputer3.Top = 165
            intComp3 = 3
        ElseIf intCompTarget = 4 Then
            imgComputer3.Left = 15                'Position 4
            imgComputer3.Top = 315
            intComp3 = 4
        ElseIf intCompTarget = 1 Then
            imgComputer3.Left = 165               'Position 1
            imgComputer3.Top = 15
            intComp3 = 1
        ElseIf intCompTarget = 5 Then
            imgComputer3.Left = 165               'Position 5
            imgComputer3.Top = 165
            intComp3 = 5
        ElseIf intCompTarget = 9 Then
            imgComputer3.Left = 165               'Position 9
            imgComputer3.Top = 315
            intComp3 = 9
        ElseIf intCompTarget = 6 Then
            imgComputer3.Left = 315               'Position 6
            imgComputer3.Top = 15
            intComp3 = 6
        ElseIf intCompTarget = 7 Then
            imgComputer3.Left = 315               'Position 7
            imgComputer3.Top = 165
            intComp3 = 7
        ElseIf intCompTarget = 2 Then
            imgComputer3.Left = 315               'Position 2
            imgComputer3.Top = 315
            intComp3 = 2
        End If
    End If
    If isGameOver Then
        cmdPlayAgain.Visible = True
        imgUser1.Enabled = False
        imgUser2.Enabled = False
        imgUser3.Enabled = False
    End If
End Sub
Public Function isGameOver()
'declare
Dim blnResult As Boolean
'initialize
blnResult = False
If intUser1 + intUser2 + intUser3 = 15 Then
    blnResult = True
    intUserScore = intUserScore + 1
    lblResults.Caption = "Congratulations! You won. Play again!"
    lblUserScore.Caption = "You: " & intUserScore
ElseIf intComp1 + intComp2 + intComp3 = 15 Then
    blnResult = True
    intCompScore = intCompScore + 1
    lblResults.Caption = "You lost. Better luck next time. Play again!"
    lblCompScore.Caption = "Computer: " & intCompScore
End If
isGameOver = blnResult
End Function
Public Function isUserFirst()
'declare
Dim intFirstX As Integer
Dim intFirstY As Integer
Dim blnUserFirst As Boolean
'initialize
intFirstX = 0
intFirstY = 0
blnUserFirst = False
If optUser.Value = True Then
    blnUserFirst = True
Else
    blnUserFirst = False
    intFirstX = Int(Rnd * 450) + 1
    intFirstY = Int(Rnd * 450) + 1
    If intFirstX > 0 And intFirstX < 150 Then
        If intFirstY > 0 And intFirstY < 150 Then          'Magic Square Position 8
            imgComputer1.Left = 15
            imgComputer1.Top = 15
            intComp1 = 8
        ElseIf intFirstY > 150 And intFirstY < 300 Then    'Magic Square Position 3
            imgComputer1.Left = 15
            imgComputer1.Top = 165
            intComp1 = 3
        Else                                               'Magic Square Position 4
            imgComputer1.Left = 15
            imgComputer1.Top = 315
            intComp1 = 4
        End If
    ElseIf intFirstX > 150 And intFirstX < 300 Then
        If intFirstY > 0 And intFirstY < 150 Then          'Magic Square Position 1
            imgComputer1.Left = 165
            imgComputer1.Top = 15
            intComp1 = 1
        ElseIf intFirstY > 150 And intFirstY < 300 Then    'Magic Square Position 5
            imgComputer1.Left = 165
            imgComputer1.Top = 165
            intComp1 = 5
        Else                                               'Magic Square Position 9
            imgComputer1.Left = 165
            imgComputer1.Top = 315
            intComp1 = 9
        End If
    Else
        If intFirstY > 0 And intFirstY < 150 Then          'Magic Square Position 6
            imgComputer1.Left = 315
            imgComputer1.Top = 15
            intComp1 = 6
        ElseIf intFirstY > 150 And intFirstY < 300 Then    'Magic Square Position 7
            imgComputer1.Left = 315
            imgComputer1.Top = 165
            intComp1 = 7
        Else                                               'Magic Square Position 2
            imgComputer1.Left = 315
            imgComputer1.Top = 315
            intComp1 = 2
        End If
    End If
End If
isUserFirst = blnUserFirst
End Function

Public Function isSquareAvailable(ByVal intAvailable As Single)
'declare
Dim blnAvailable As Boolean
'initialize
blnAvailable = False
If intUser1 > 0 And intUser1 = intAvailable Then
    blnAvailable = False
ElseIf intUser2 > 0 And intUser2 = intAvailable Then
    blnAvailable = False
ElseIf intUser3 > 0 And intUser3 = intAvailable Then
    blnAvailable = False
ElseIf intComp1 > 0 And intComp1 = intAvailable Then
    blnAvailable = False
ElseIf intComp2 > 0 And intComp2 = intAvailable Then
    blnAvailable = False
ElseIf intComp3 > 0 And intComp3 = intAvailable Then
    blnAvailable = False
Else
    If 1 <= intAvailable <= 9 Then
        blnAvailable = True
    End If
End If
isSquareAvailable = blnAvailable
End Function

Public Function compPieceNum()
'declare
Dim intGetMarker As Integer
Dim intMarkerIs2 As Integer
'initialize
intGetMarker = 0
intMarkerIs2 = 0
If intComp1 < 0 Then
    intGetMarker = 1
ElseIf intComp2 < 0 Then
    intGetMarker = 2
ElseIf intComp3 < 0 Then
    intGetMarker = 3
Else
    'this is to prevent the computer from moving a piece that is currently blocking the user
    intGetMarker = Int(Rnd * 3) + 1
    If intGetMarker = 1 Then
        If intComp1 + intUser1 + intUser2 = 15 Or intComp1 + intUser1 + intUser3 = 15 Or intComp1 + intUser2 + intUser3 Then
            intGetMarker = Int(Rnd * 2) + 2
        End If
    ElseIf intGetMarker = 2 Then
        If intComp2 + intUser1 + intUser2 = 15 Or intComp2 + intUser1 + intUser3 = 15 Or intComp2 + intUser2 + intUser3 Then
            intMarkerIs2 = Int(Rnd * 2) + 1
            If intMarkerIs2 = 1 Then
                intGetMarker = 1
            Else
                intGetMarker = 3
            End If
        End If
    Else
        If intComp3 + intUser1 + intUser2 = 15 Or intComp3 + intUser1 + intUser3 = 15 Or intComp3 + intUser2 + intUser3 Then
            intGetMarker = Int(Rnd * 2) + 1
        End If
    End If
End If
compPieceNum = intGetMarker
End Function

