VERSION 5.00
Object = "*\AprjSymbolSelector.vbp"
Begin VB.Form frmMain 
   Caption         =   "TEST Symbol Selector"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      Begin VB.Label lblSymbol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   900
      End
   End
   Begin vb6projectSymbolSelector.SymbolSelector SymbolSelector1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SymbolSelector1_SelectionChange(FontName As String, Item As Integer)
   With lblSymbol
      .FontName = SymbolSelector1.SymbolFont
      .Caption = Chr(Item)
      .Left = (picCanvas.Width \ 2) - (.Width \ 2)
      .Top = (picCanvas.Height \ 2) - (.Height \ 2)
   End With
End Sub
