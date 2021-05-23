VERSION 5.00
Begin VB.UserControl UserControlSquare 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   405
   ScaleWidth      =   405
   Begin VB.Label Btn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   420
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
Dim State As SquareState
Public IsGenerated As Boolean

Enum SquareState
    stUnknown
    stRevealed
    stFlag
    stQuestion
    stExploded
End Enum

Public Property Get Color() As OLE_COLOR
  Color = Btn.ForeColor
End Property

Public Property Let Color(ByVal c As OLE_COLOR)
  Btn.ForeColor = c
End Property

Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Clicked(Button)
End Sub

Sub RevealByFloodfill()
    State = stRevealed
    BorderStyle = 0
    Btn.BorderStyle = 1
    Call ShowState
End Sub

Public Sub Explode()
    If State = stExploded Then
        Exit Sub
    End If
    Btn.Caption = "*"
    State = stExploded
End Sub

Sub Uncover()
    ' Cannot uncover flagged squares
    If State = stFlag Then
        Exit Sub
    End If
    If HasMine Then
        Call Explode
        Color = vbRed
        Call ParentControls(0).RevealMines
    Else
        State = stRevealed
        BorderStyle = 0
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
            Color = vbBlue
        Case 2:
            Color = &H18F01
        Case 3:
            Color = vbRed
        Case 4:
            Color = &H8F0001
        Case 5:
            Color = &H2018F
        Case 6:
            Color = &H8F8F01
        Case 7:
            Color = vbBlack
        Case 8:
            Color = &H8F8F8F
            
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
            Btn.Caption = " "
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
    Call Clicked(Button)
End Sub

Private Sub Clicked(Button As Integer)
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
