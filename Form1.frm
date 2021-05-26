VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Minesweeper"
   ClientHeight    =   705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControlField UserControlField1 
      Height          =   8055
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   9495
      _ExtentX        =   10610
      _ExtentY        =   10821
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Play"
      Height          =   555
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   120
   End
   Begin VB.Label LabelMines 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuNew 
         Caption         =   "New"
      End
      Begin VB.Menu MenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuBeginner 
         Caption         =   "Beginner"
      End
      Begin VB.Menu MenuIntermediate 
         Caption         =   "Intermediate"
      End
      Begin VB.Menu MenuExpert 
         Caption         =   "Expert"
      End
      Begin VB.Menu MenuCustom 
         Caption         =   "Custom..."
      End
      Begin VB.Menu MenuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ticks As Integer
Dim CustomFieldPicked As Boolean

Private Sub StartGame()
    Ticks = 0
    Timer1.Enabled = True
    Command1.Caption = ":-)"
    Call UserControlField1.Initialize
    Call UpdateFlaggedMines
End Sub

Private Sub Command1_Click()
    Call StartGame
End Sub

Private Sub Form_Load()
    Ticks = 0
    Configuration.ShowAll = False
    Call ConfigureGame(4, 4, 2)
End Sub

Private Sub Timer1_Timer()
    Ticks = Ticks + 1
    LabelTime.Caption = Ticks
End Sub

Public Sub UpdateFlaggedMines()
    LabelMines.Caption = (Configuration.Mines - Minefield.FlaggedMines)
End Sub

Public Sub GameOver(Win As Boolean)
    Timer1.Enabled = False
    If Win Then
        Command1.Caption = "B-)"
        MsgBox ("Congratulations!")
    Else
        Command1.Caption = ":-("
    End If
End Sub

Private Sub MenuExit_Click()
    Unload Me
End Sub

Private Sub Resize()
    Dim myWidth As Integer, myHeight As Integer
    myWidth = 400
    myHeight = 1400
    Width = myWidth + (Configuration.Rows * 405)
    Height = myHeight + (Configuration.Columns * 405)
End Sub

Public Sub ConfigureGame(Rows As Integer, Columns As Integer, Mines As Integer)
    Configuration.Rows = Rows
    Configuration.Columns = Columns
    Configuration.Mines = Mines
    Call Resize
    Call StartGame
End Sub

Private Sub MenuNew_Click()
    Call StartGame
End Sub

Private Sub MenuBeginner_Click()
    Call ConfigureGame(9, 9, 10)
End Sub

Private Sub MenuIntermediate_Click()
    Call ConfigureGame(16, 16, 40)
End Sub

Private Sub MenuExpert_Click()
    Call ConfigureGame(16, 20, 99)
End Sub

Private Sub MenuCustom_Click()
    Dialog.Show vbModal
End Sub

Public Sub SetResult(CustomFieldResult As Boolean)
    CustomFieldPicked = CustomFieldResult
End Sub
