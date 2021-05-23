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
Sub LayoutControls()
    ' Create all Squares and assign mines to them
    Dim sq As Control
    Dim Row As Integer
    Dim col As Integer
    Dim idx As Integer
    Dim Mines As Integer
    For col = 0 To Configuration.Columns - 1
        For Row = 0 To Configuration.Rows - 1
            idx = Minefield.GetIndex(col, Row)
            Load Squares(idx)
            
            Squares(idx).Move col * 400, Row * 400, 400, 400
            
            If Minefield.HasMine(Row, col) Then
                Squares(idx).HasMine = True
            End If
            Squares(idx).Visible = True
            'Mines = Minefield.GetNeighborCount(Row, col)
            Mines = Minefield.MinesWithCount(Minefield.GetIndex(Row, col))
            Squares(idx).MineNeighborCount = Mines
            Call Squares(idx).ShowState
        Next
    Next
    
End Sub

Public Sub Initialize()
    Call Minefield.GenerateMines
    Call LayoutControls
End Sub

Public Sub Test()
    MsgBox ("Boo")
End Sub

Public Sub Floodfill(Row As Integer, Column As Integer)
    '4-way flood fill from this point
    ' now we can iterate over Square items, as they are indexed with GetIndex
    ' consult Minefield for neighbor counts
    Dim index As Integer
    index = Minefield.GetIndex(Row + 1, Column)
    If Minefield.MinesWithCount(index) = 0 Then
        'Reveal this mine
        Call Squares(index).RevealByFloodfill
    End If
End Sub

