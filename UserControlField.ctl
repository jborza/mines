VERSION 5.00
Begin VB.UserControl UserControlField 
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   ScaleHeight     =   6300
   ScaleWidth      =   5010
   Begin Project1.UserControlSquare Squares 
      Height          =   375
      Index           =   32767
      Left            =   33000
      TabIndex        =   0
      Top             =   5280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
End
Attribute VB_Name = "UserControlField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Visited() As Boolean
Dim MaxLoadedSquare As Integer

Sub UnloadAll()
    For Each Square In Squares
        If Square.IsGenerated Then
            Square.Visible = False
            Unload Square
        End If
    Next
End Sub

Sub LayoutControls()
    ' Create all Squares and assign mines to them
    Dim sq As Control
    Dim Row As Integer
    Dim col As Integer
    Dim idx As Integer
    Dim mines As Integer
    Call UnloadAll
    
    For col = 0 To Configuration.columns - 1
        For Row = 0 To Configuration.rows - 1
            idx = Minefield.GetIndex(Row, col)
            Load Squares(idx)
            With Squares(idx)
                .Move col * 405, Row * 405, 405, 405
            
                If Minefield.HasMine(Row, col) Then
                    .HasMine = True
                End If
                .Visible = True
                .MineNeighborCount = Minefield.MinesWithCount(Minefield.GetIndex(Row, col))
                .Column = col
                .Row = Row
                .IsGenerated = True
                Call .ShowState
            End With
        Next
    Next
    
End Sub

Public Sub Initialize()
    Call Minefield.GenerateMines
    Call LayoutControls
    ReDim Visited(Configuration.columns * Configuration.rows) As Boolean
End Sub

Public Sub RevealMines()
    Dim index As Integer
    For i = 1 To Minefield.MineLookup.Count
        index = Minefield.MineLookup(i)
        Call Squares(index).Explode
    Next
    Call Form1.GameOver(False)
End Sub

Public Sub DoFloodfill(Row As Integer, Column As Integer)
    'Erase Visited
    'ReDim Visited(Configuration.columns * Configuration.rows) As Boolean
    Call Floodfill(Row, Column)
End Sub

Private Function GetRevealedCount() As Integer
    GetRevealedCount = 0
    For i = 0 To Configuration.columns * Configuration.rows - 1
        If Visited(i) Then
            GetRevealedCount = GetRevealedCount + 1
        End If
    Next
End Function

Public Sub MarkAsRevealed(index As Integer)
    Call Squares(index).RevealByFloodfill
    Visited(index) = True
    Minefield.RevealedMineCount = Minefield.RevealedMineCount + 1
    If Minefield.IsGameWon Then
        Call Form1.GameOver(True)
    End If
End Sub

Public Sub Floodfill(Row As Integer, Column As Integer)
    ' 8-way flood fill from this point
    ' now we can iterate over Square items, as they are indexed with GetIndex
    ' consult Minefield for neighbor counts
    Dim index As Integer
    If Not Minefield.CellExists(Row, Column) Then
        Exit Sub
    End If
    index = Minefield.GetIndex(Row, Column)
    If Visited(index) Then
        Exit Sub
    End If
    ' uncover this one
    Call MarkAsRevealed(index)
    
    ' Don't continue if a mine is adjacent
    If Minefield.GetNeighborCount(Row, Column) > 0 Then
        Exit Sub
    End If
    Call Floodfill(Row + 1, Column)
    Call Floodfill(Row + 1, Column + 1)
    Call Floodfill(Row, Column + 1)
    Call Floodfill(Row - 1, Column + 1)
    Call Floodfill(Row - 1, Column)
    Call Floodfill(Row - 1, Column - 1)
    Call Floodfill(Row, Column - 1)
    Call Floodfill(Row + 1, Column - 1)
End Sub

