VERSION 5.00
Begin VB.UserControl UserControlSquare 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   ScaleHeight     =   405
   ScaleWidth      =   405
   Begin VB.CommandButton Btn 
      Caption         =   "?"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   400
   End
End
Attribute VB_Name = "UserControlSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public HasMine As Boolean
Public MineNeighborCount As Variant
Public Row As Integer
Public Column As Integer
Public FieldControl As UserControl


Enum SquareState
    stUnknown
    stRevealed
    stFlag
    stQuestion
End Enum

Dim State As SquareState

Sub RevealByFloodfill()
    State = stRevealed
    Call ShowState
End Sub

Sub Uncover()
    If HasMine Then
        MsgBox ("Game over!")
        'TODO stop the entire game
    Else
        State = stRevealed
        Enabled = False
        If MineNeighborCount = 0 Then
            Call ParentControls(0).DoFloodfill(Row, Column)
            Dim filledNeighbors() As Integer
        End If
    End If
    
End Sub

Sub Mark()
    ' Toggle between flag, question mark and empty
    If State = stUnknown Then
        State = stFlag
        Call Minefield.Flag
    ElseIf State = stFlag Then
        State = stQuestion
        Call Minefield.Unflag
    ElseIf State = stQuestion Then
        State = stUnknown
    End If
    ' Else is state stRevealed, no nothing
End Sub

Sub Colorize()
    Select Case MineNeighborCount
        Case 1:
            'Call Btn.SetForeColor(&HFF)
    End Select
    
End Sub

Public Sub ShowState()
    Select Case State
        Case stUnknown
        If Configuration.ShowAll Then
            If HasMine Then
                Btn.Caption = "."
            Else
                Btn.Caption = MineNeighborCount ' " "
            End If
        Else
            Btn.Caption = "."
        End If
        Case stRevealed
            ' blank or number, not 0
            If MineNeighborCount = 0 Then
                Btn.Caption = " "
            Else
                Btn.Caption = MineNeighborCount
            End If
            Call Colorize
        Case stFlag
            Btn.Caption = "X"
        Case stQuestion
            Btn.Caption = "?"
    End Select
End Sub

Sub Btn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Call Uncover
ElseIf Button = vbRightButton Then
    Call Mark
Else
    
End If
    Call ShowState
End Sub

Private Sub UserControl_Initialize()
    State = stUnknown
    Call ShowState
End Sub


Public Sub SetNeighborCount(ByVal Count As Integer)
    NeighborCount = Count
End Sub

