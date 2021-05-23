Attribute VB_Name = "Minefield"
Public MineLocations() As Point

'stores mine neighbor counts
Public MinesWithCount() As Integer

Public Function GetIndex(Row As Integer, Column As Integer) As Integer
     GetIndex = Row * Configuration.Columns + Column
End Function

Sub GenerateMines()
    ReDim MineLocations(Configuration.Mines - 1) As Point
    Dim pt As Point
    
    Dim i As Integer
    For i = 0 To Configuration.Mines - 1
        pt.x = Int((Configuration.Columns * Rnd))
        pt.y = Int((Configuration.Rows * Rnd))
        MineLocations(i) = pt
    Next
    Call PopulateNeighbors
End Sub

Sub PopulateNeighbors()
    ReDim MinesWithCount(Configuration.Columns * Configuration.Rows)
    ' insert mines first
    Dim index As Integer
    For i = LBound(MineLocations) To UBound(MineLocations)
        index = GetIndex(MineLocations(i).y, MineLocations(i).x)
        MinesWithCount(index) = -1
    Next
    Dim Row As Integer
    Dim col As Integer
    For col = 0 To Configuration.Columns - 1
        For Row = 0 To Configuration.Rows - 1
            index = GetIndex(Row, col)
            If MinesWithCount(index) <> -1 Then
                MinesWithCount(index) = GetNeighborCount(Row, col)
            End If
        Next Row
    Next col
End Sub

Public Function HasMine(Row As Integer, Column As Integer) As Boolean
    HasMine = MinesWithCount(GetIndex(Row, Column)) = -1
End Function


'TODO precalculate into an array
Public Function HasMineOld(Row As Integer, Column As Integer) As Boolean
    HasMineOld = False
    Dim i As Integer
    For i = LBound(MineLocations) To UBound(MineLocations)
        If MineLocations(i).x = Column And MineLocations(i).y = Row Then
            HasMineOld = True
            Exit Function
        End If
    Next
End Function

'TODO precalculate this
Public Function GetNeighborCount(Row As Integer, Column As Integer) As Integer
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
    GetNeighborCount = Count
End Function

