VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form mdiGraph 
   Caption         =   "Graph"
   ClientHeight    =   4785
   ClientLeft      =   2730
   ClientTop       =   1695
   ClientWidth     =   5085
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Graph_fo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4785
   ScaleWidth      =   5085
   Begin GraphLib.Graph graSample 
      Height          =   4005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3510
      _Version        =   65536
      _ExtentX        =   6191
      _ExtentY        =   7064
      _StockProps     =   96
      BorderStyle     =   1
      AutoInc         =   0
      GraphType       =   0
      RandomData      =   0
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
End
Attribute VB_Name = "mdiGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  graSample.RandomData = 1
  graSample.GraphType = 1
  graSample.GraphStyle = 1

  graSample.DrawMode = 2
  i% = DoEvents()
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'''If the user double clicks on the control box minimize the form
'''instead of closing it
  If UnloadMode = 0 Then
    Cancel = True
    Me.WindowState = 1
  End If

End Sub

Private Sub Form_Resize()
  graSample.Width = Me.ScaleWidth
  graSample.Height = Me.ScaleHeight
End Sub

