VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Field "
   ClientHeight    =   1725
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextMines 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Text            =   "15"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox TextRows 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "10"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox TextColumns 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "10"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Mines"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Height"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Width"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Form1.SetResult False
    Unload Me
End Sub

Private Sub OKButton_Click()
    Dim rows As Integer, columns As Integer, mines As Integer
    rows = Val(TextRows.Text)
    columns = Val(TextColumns.Text)
    mines = Val(TextMines.Text)
    
    Call Form1.ConfigureGame(rows, columns, mines)
    Form1.SetResult True
    Unload Me
End Sub
