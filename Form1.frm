VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Minesweeper"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControlField UserControlField1 
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
      _extentx        =   10610
      _extenty        =   10821
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Play"
      Height          =   555
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   600
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
      Top             =   600
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
      Left            =   3600
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ticks As Integer


Private Sub Command1_Click()
    Call UserControlField1.Initialize
    Call UpdateFlaggedMines
End Sub

Private Sub Form_Load()
    Ticks = 0
    Configuration.Rows = 10
    Configuration.Columns = 10
    Configuration.Mines = 10
    Configuration.ShowAll = True
End Sub

Private Sub Timer1_Timer()
    Ticks = Ticks + 1
    LabelTime.Caption = Ticks
End Sub

Public Sub UpdateFlaggedMines()
    LabelMines.Caption = (Configuration.Mines - Minefield.FlaggedMines)
End Sub
