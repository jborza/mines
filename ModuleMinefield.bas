Attribute VB_Name = "Minefield"
Public FlaggedMines As Integer
'stores mine neighbor counts
Public MinesWithCount() As Integer
Public MineLookup As New Collection
Public RevealedMineCount As Integer

Public Function GetIndex(Row As Integer, Column As Integer) As Integer
     GetIndex = Row * Configuration.Columns + Column
End Function

Public Function Exists(ByRef col As Collection, ByVal key As String) As Boolean
    On Error GoTo DoesntExist
    col.Item (key)
    Exists = True
    Exit Function
DoesntExist:
    Exists = False
End Function

Sub GenerateMines()
    
    Dim pt As Point
    Dim i As Integer
    Dim key As Integer
    ' reset MineLookup
    Set MineLookup = Nothing
    Set MineLookup = New Collection
    Do While MineLookup.Count < Configuration.Mines
    
        pt.x = Int((Configuration.Columns * Rnd))
        pt.y = Int((Configuration.Rows * Rnd))
        key = Minefield.GetIndex(pt.y, pt.x)
        ' don't insert the same mine twice
        If Exists(MineLookup, Str(key)) Then
            i = i - 1
        Else
            Call MineLookup.Add(key, Str(key))
        End If
    Loop
    Call PopulateNeighbors
End Sub

Sub PopulateNeighbors()
    ReDim MinesWithCount(Configuration.Columns * Configuration.Rows)

    ' insert mines first
    Dim index As Integer
    For i = 1 To MineLookup.Count
        index = MineLookup(i)
        MinesWithCount(index) = -1
    Next
    Dim Row As Integer
    Dim col As Integer
    For col = 0 To Configuration.Columns - 1
        For Row = 0 To Configuration.Rows - 1
            index = GetIndex(Row, col)
            If MinesWithCount(index) <> -1 Then
                MinesWithCount(index) = CalculateNeighborCount(Row, col)
            End If
        Next Row
    Next col
End Sub

Public Function HasMine(Row As Integer, Column As Integer) As Boolean
    HasMine = MinesWithCount(GetIndex(Row, Column)) = -1
End Function

Public Function CellExists(Row As Integer, Column As Integer) As Boolean
    CellExists = True
    If Column < 0 Or Column >= Configuration.Columns Then
        CellExists = False
    End If
    If Row < 0 Or Row >= Configuration.Rows Then
        CellExists = False
    End If
End Function

Public Function GetNeighborCount(Row As Integer, Column As Integer) As Integer
    GetNeighborCount = MinesWithCount(GetIndex(Row, Column))
End Function

Function CalculateNeighborCount(Row As Integer, Column As Integer) As Integer
    Dim Count As Integer
    Dim x As Integer
    Dim y As Integer
    For x = Column - 1 To Column + 1
        For y = Row - 1 To Row + 1
            If x < 0 Or x >= Configuration.Columns Then
                GoTo Continue
            End If
            If y < 0 Or y >= Configuration.Rows Then
                GoTo Continue
            End If
            If HasMine(y, x) Then
                Count = Count + 1
            End If
Continue:
        Next y
    Next x
    CalculateNeighborCount = Count
End Function

Public Sub Flag()
    FlaggedMines = FlaggedMines + 1
    Call Form1.UpdateFlaggedMines
End Sub

Public Sub Unflag()
    FlaggedMines = FlaggedMines - 1
    Call Form1.UpdateFlaggedMines
End Sub

Public Function IsGameWon() As Boolean
    IsGameWon = RevealedMineCount + Configuration.Mines = Configuration.Rows * Configuration.Columns
End Function
