VERSION 5.00
Begin VB.UserControl UserControlField 
   BackColor       =   &H80000018&
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   ScaleHeight     =   6300
   ScaleWidth      =   5010
   Begin Project1.UserControlSquare Squares 
      Height          =   495
      Index           =   32767
      Left            =   4080
      TabIndex        =   0
      Top             =   5280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "UserControlField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim Visited() As Boolean

Sub LayoutControls()
    ' Create all Squares and assign mines to them
    Dim sq As Control
    Dim Row As Integer
    Dim col As Integer
    Dim idx As Integer
    Dim Mines As Integer
    For col = 0 To Configuration.Columns - 1
        For Row = 0 To Configuration.Rows - 1
            idx = Minefield.GetIndex(Row, col)
            Load Squares(idx)
            With Squares(idx)
                .Move col * 400, Row * 400, 400, 400
            
                If Minefield.HasMine(Row, col) Then
                    .HasMine = True
                End If
                .Visible = True
                .MineNeighborCount = Minefield.MinesWithCount(Minefield.GetIndex(Row, col))
                .Column = col
                .Row = Row
                Call .ShowState
            End With
        Next
    Next
    
End Sub

Public Sub Initialize()
    Call Minefield.GenerateMines
    Call LayoutControls
End Sub


Public Sub DoFloodfill(Row As Integer, Column As Integer)
    Erase Visited
    ReDim Visited(Configuration.Columns * Configuration.Rows) As Boolean
    Call Floodfill(Row, Column)
    
End Sub

Public Sub Floodfill(Row As Integer, Column As Integer)
    '4-way flood fill from this point
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
    Call Squares(index).RevealByFloodfill
    Visited(index) = True
    
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

