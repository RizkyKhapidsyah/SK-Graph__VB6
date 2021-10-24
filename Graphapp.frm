VERSION 5.00
Begin VB.MDIForm frmGraphApp 
   BackColor       =   &H8000000C&
   Caption         =   "GraphApp 95 "
   ClientHeight    =   8490
   ClientLeft      =   780
   ClientTop       =   1590
   ClientWidth     =   12060
   Icon            =   "Graphapp.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Begin VB.PictureBox PicTabContainer 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9060
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   12060
      Begin VB.Frame fratabs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Miscellaneous Properties"
         Height          =   1605
         Index           =   5
         Left            =   45
         TabIndex        =   88
         Top             =   7185
         Width           =   12960
         Begin VB.CheckBox chkThln 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Thick Lines"
            Height          =   210
            Left            =   10920
            TabIndex        =   44
            Top             =   615
            Width           =   1800
         End
         Begin VB.CheckBox chkPtln 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Patterened Lines"
            Height          =   210
            Left            =   9075
            TabIndex        =   43
            Top             =   615
            Width           =   1800
         End
         Begin VB.ComboBox cboLineStats 
            Height          =   300
            Left            =   105
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   630
            Width           =   5730
         End
         Begin VB.ComboBox cboPrintStyle 
            Height          =   300
            Left            =   5820
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   630
            Width           =   2280
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Line Stats"
            Height          =   210
            Left            =   180
            TabIndex        =   89
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label Label36 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Print Style"
            Height          =   210
            Left            =   5820
            TabIndex        =   90
            Top             =   375
            Width           =   1065
         End
      End
      Begin VB.Frame fratabs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   1695
         Index           =   1
         Left            =   60
         TabIndex        =   61
         Top             =   3735
         Width           =   13335
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Single Dimension Array"
            Height          =   1470
            Left            =   4440
            TabIndex        =   62
            Top             =   165
            Width           =   7215
            Begin VB.HScrollBar hsbPositionBar 
               Height          =   270
               Index           =   2
               Left            =   120
               Min             =   1
               TabIndex        =   12
               Top             =   555
               Value           =   1
               Width           =   1560
            End
            Begin VB.TextBox txtPointData 
               Height          =   285
               Index           =   3
               Left            =   5130
               TabIndex        =   17
               Top             =   855
               Width           =   2000
            End
            Begin VB.TextBox txtPointData 
               Height          =   285
               Index           =   2
               Left            =   2130
               TabIndex        =   16
               Top             =   870
               Width           =   2000
            End
            Begin VB.ComboBox cboEnumArrays 
               Height          =   300
               Index           =   0
               Left            =   2145
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   405
               Width           =   1650
            End
            Begin VB.ComboBox cboEnumArrays 
               Height          =   300
               Index           =   1
               Left            =   3795
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   405
               Width           =   1650
            End
            Begin VB.ComboBox cboEnumArrays 
               Height          =   300
               Index           =   2
               Left            =   5445
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   405
               Width           =   1650
            End
            Begin VB.Label Label5 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Remember:  One dimensional array properties are indexed by ThisPoint and NOT ThisSet."
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   495
               TabIndex        =   67
               Top             =   1185
               Width           =   6375
            End
            Begin VB.Label Label4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "This Point"
               Height          =   225
               Left            =   105
               TabIndex        =   63
               Top             =   315
               Width           =   1560
            End
            Begin VB.Label Label9 
               BackColor       =   &H00C0C0C0&
               Caption         =   "LegendText"
               Height          =   240
               Left            =   4245
               TabIndex        =   64
               Top             =   930
               Width           =   1005
            End
            Begin VB.Label Label8 
               BackColor       =   &H00C0C0C0&
               Caption         =   "LabelText"
               Height          =   240
               Left            =   1380
               TabIndex        =   66
               Top             =   930
               Width           =   870
            End
            Begin VB.Label Label11 
               BackColor       =   &H00C0C0C0&
               Caption         =   "SymbolData"
               Height          =   240
               Left            =   3795
               TabIndex        =   79
               Top             =   165
               Width           =   1575
            End
            Begin VB.Label Label10 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PatternData"
               Height          =   240
               Left            =   5490
               TabIndex        =   78
               Top             =   165
               Width           =   1560
            End
            Begin VB.Label Label6 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ExtraData"
               Height          =   240
               Left            =   2145
               TabIndex        =   77
               Top             =   165
               Width           =   1515
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2D Arrays"
            Height          =   1470
            Left            =   525
            TabIndex        =   75
            Top             =   150
            Width           =   3735
            Begin VB.TextBox txtPointData 
               Height          =   285
               Index           =   4
               Left            =   2865
               TabIndex        =   10
               Top             =   540
               Width           =   700
            End
            Begin VB.TextBox txtPointData 
               Height          =   285
               Index           =   5
               Left            =   2865
               TabIndex        =   11
               Top             =   1035
               Width           =   700
            End
            Begin VB.HScrollBar hsbPositionBar 
               Height          =   270
               Index           =   1
               Left            =   165
               Max             =   2
               Min             =   1
               TabIndex        =   9
               Top             =   1095
               Value           =   1
               Width           =   1500
            End
            Begin VB.HScrollBar hsbPositionBar 
               Height          =   270
               Index           =   0
               Left            =   150
               Max             =   1
               Min             =   1
               TabIndex        =   8
               Top             =   525
               Value           =   1
               Width           =   1500
            End
            Begin VB.Label Label12 
               BackColor       =   &H00C0C0C0&
               Caption         =   "XPosData"
               Height          =   240
               Left            =   1935
               TabIndex        =   65
               Top             =   615
               Width           =   1005
            End
            Begin VB.Label Label7 
               BackColor       =   &H00C0C0C0&
               Caption         =   "GraphData"
               Height          =   240
               Left            =   1905
               TabIndex        =   68
               Top             =   1110
               Width           =   915
            End
            Begin VB.Label lblPlab 
               BackColor       =   &H00C0C0C0&
               Caption         =   "This set ( 1 to 1 )"
               Height          =   240
               Index           =   0
               Left            =   165
               TabIndex        =   69
               Top             =   285
               Width           =   1320
            End
            Begin VB.Label lblPlab 
               BackColor       =   &H00C0C0C0&
               Caption         =   "This point (1 to 2)"
               Height          =   240
               Index           =   1
               Left            =   150
               TabIndex        =   76
               Top             =   855
               Width           =   1395
            End
         End
      End
      Begin VB.Frame fratabs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Graph"
         Height          =   1005
         Index           =   0
         Left            =   90
         TabIndex        =   58
         Top             =   1125
         Width           =   11790
         Begin VB.TextBox txtPointData 
            Height          =   285
            Index           =   1
            Left            =   6210
            TabIndex        =   3
            Text            =   "2"
            Top             =   585
            Width           =   525
         End
         Begin VB.TextBox txtPointData 
            Height          =   285
            Index           =   0
            Left            =   6210
            TabIndex        =   2
            Text            =   "1"
            Top             =   210
            Width           =   525
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Grid"
            Height          =   630
            Left            =   7050
            TabIndex        =   73
            Top             =   225
            Width           =   4545
            Begin VB.OptionButton optGridP 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Both"
               Height          =   285
               Index           =   3
               Left            =   3690
               TabIndex        =   7
               Top             =   240
               Width           =   675
            End
            Begin VB.OptionButton optGridP 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Vertical"
               Height          =   285
               Index           =   2
               Left            =   2775
               TabIndex        =   6
               Top             =   240
               Width           =   870
            End
            Begin VB.OptionButton optGridP 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Horizontal"
               Height          =   285
               Index           =   1
               Left            =   1665
               TabIndex        =   5
               Top             =   240
               Width           =   1050
            End
            Begin VB.OptionButton optGridP 
               BackColor       =   &H00C0C0C0&
               Caption         =   "(Default) None"
               Height          =   285
               Index           =   0
               Left            =   240
               TabIndex        =   4
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.ComboBox cboGRTS 
            Height          =   300
            Index           =   0
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   525
            Width           =   1000
         End
         Begin VB.ComboBox cboGRTS 
            Height          =   300
            Index           =   1
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   525
            Width           =   2450
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Number of points in each set"
            Height          =   240
            Left            =   3885
            TabIndex        =   70
            Top             =   615
            Width           =   2130
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Number of Sets in the graph"
            Height          =   240
            Left            =   3885
            TabIndex        =   71
            Top             =   270
            Width           =   2115
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Graph Type"
            Height          =   240
            Left            =   90
            TabIndex        =   60
            Top             =   300
            Width           =   930
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Graph Style"
            Height          =   240
            Left            =   1080
            TabIndex        =   59
            Top             =   300
            Width           =   930
         End
      End
      Begin VB.Frame fratabs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Axis Properties"
         Height          =   1605
         Index           =   4
         Left            =   45
         TabIndex        =   53
         Top             =   5505
         Width           =   9030
         Begin VB.TextBox txtYAxv 
            Height          =   285
            Index           =   4
            Left            =   8130
            TabIndex        =   40
            Top             =   1080
            Width           =   660
         End
         Begin VB.TextBox txtYAxv 
            Height          =   285
            Index           =   3
            Left            =   8115
            TabIndex        =   39
            Top             =   600
            Width           =   660
         End
         Begin VB.ComboBox cboAx_Props 
            Height          =   300
            Index           =   2
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   690
            Width           =   1320
         End
         Begin VB.TextBox txtYAxv 
            Height          =   285
            Index           =   2
            Left            =   4395
            TabIndex        =   37
            Top             =   1065
            Width           =   660
         End
         Begin VB.TextBox txtYAxv 
            Height          =   285
            Index           =   1
            Left            =   4395
            TabIndex        =   36
            Top             =   630
            Width           =   660
         End
         Begin VB.TextBox txtYAxv 
            Height          =   285
            Index           =   0
            Left            =   4395
            TabIndex        =   35
            Top             =   270
            Width           =   660
         End
         Begin VB.ComboBox cboAx_Props 
            Height          =   300
            Index           =   1
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   630
            Width           =   1320
         End
         Begin VB.ComboBox cboAx_Props 
            Height          =   300
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   630
            Width           =   1320
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Label Every"
            Height          =   210
            Left            =   7080
            TabIndex        =   87
            Top             =   1140
            Width           =   1065
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tick Every"
            Height          =   210
            Left            =   7065
            TabIndex        =   86
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ticks"
            Height          =   210
            Left            =   5715
            TabIndex        =   85
            Top             =   390
            Width           =   630
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y Axis Ticks"
            Height          =   210
            Left            =   3315
            TabIndex        =   84
            Top             =   1170
            Width           =   960
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y Axis Max"
            Height          =   210
            Left            =   3315
            TabIndex        =   83
            Top             =   735
            Width           =   960
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y Axis Min"
            Height          =   210
            Left            =   3345
            TabIndex        =   82
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y Axis Style"
            Height          =   210
            Left            =   1845
            TabIndex        =   81
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label label60 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Y Axis Position"
            Height          =   210
            Left            =   180
            TabIndex        =   80
            Top             =   375
            Width           =   1065
         End
      End
      Begin VB.Frame fratabs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Captions & Fonts"
         Height          =   1020
         Index           =   3
         Left            =   0
         TabIndex        =   48
         Top             =   60
         Width           =   13440
         Begin VB.TextBox txtFontsize 
            Height          =   300
            Left            =   12390
            TabIndex        =   32
            Top             =   525
            Width           =   705
         End
         Begin VB.ComboBox cboFontInfo 
            Height          =   300
            Index           =   2
            Left            =   10575
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   525
            Width           =   1820
         End
         Begin VB.ComboBox cboFontInfo 
            Height          =   300
            Index           =   1
            Left            =   9075
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   525
            Width           =   1500
         End
         Begin VB.ComboBox cboFontInfo 
            Height          =   300
            Index           =   0
            Left            =   7275
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   525
            Width           =   1800
         End
         Begin VB.TextBox txtTitles 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   25
            Top             =   540
            Width           =   1740
         End
         Begin VB.TextBox txtTitles 
            Height          =   285
            Index           =   1
            Left            =   1815
            TabIndex        =   26
            Top             =   540
            Width           =   1740
         End
         Begin VB.TextBox txtTitles 
            Height          =   285
            Index           =   2
            Left            =   3570
            TabIndex        =   27
            Top             =   540
            Width           =   1740
         End
         Begin VB.TextBox txtTitles 
            Height          =   285
            Index           =   3
            Left            =   5325
            TabIndex        =   28
            Top             =   540
            Width           =   1740
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Font Size"
            Height          =   240
            Left            =   12375
            TabIndex        =   54
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Font Style"
            Height          =   240
            Left            =   10560
            TabIndex        =   55
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Font Family"
            Height          =   240
            Left            =   9090
            TabIndex        =   56
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Font Use"
            Height          =   240
            Left            =   7275
            TabIndex        =   57
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Left Title"
            Height          =   240
            Left            =   3600
            TabIndex        =   52
            Top             =   300
            Width           =   810
         End
         Begin VB.Label label20 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Graph Title"
            Height          =   240
            Index           =   0
            Left            =   1860
            TabIndex        =   51
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bottom Title"
            Height          =   240
            Left            =   75
            TabIndex        =   50
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Graph Caption"
            Height          =   240
            Left            =   5325
            TabIndex        =   49
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame fratabs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Color"
         Height          =   1470
         Index           =   2
         Left            =   75
         TabIndex        =   46
         Top             =   2220
         Width           =   10665
         Begin VB.PictureBox picColorPallete 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   90
            ScaleHeight     =   49
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   204
            TabIndex        =   91
            Top             =   285
            Width           =   3120
         End
         Begin VB.HScrollBar hsbColorBar 
            Height          =   270
            Left            =   4860
            Max             =   2
            Min             =   1
            TabIndex        =   21
            Top             =   660
            Value           =   1
            Width           =   1500
         End
         Begin VB.Frame frame5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Palette"
            Height          =   600
            Left            =   6660
            TabIndex        =   72
            Top             =   315
            Width           =   3885
            Begin VB.OptionButton optPalt 
               BackColor       =   &H00C0C0C0&
               Caption         =   "(Default) Solid"
               Height          =   285
               Index           =   0
               Left            =   195
               TabIndex        =   22
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optPalt 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Pastel"
               Height          =   285
               Index           =   1
               Left            =   1605
               TabIndex        =   23
               Top             =   225
               Width           =   825
            End
            Begin VB.OptionButton optPalt 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Grayscale"
               Height          =   285
               Index           =   2
               Left            =   2565
               TabIndex        =   24
               Top             =   225
               Width           =   1110
            End
         End
         Begin VB.OptionButton optColor_Item 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Colordata "
            Height          =   315
            Index           =   3
            Left            =   3300
            TabIndex        =   20
            Top             =   750
            Width           =   1250
         End
         Begin VB.OptionButton optColor_Item 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Foreground"
            Height          =   315
            Index           =   2
            Left            =   3300
            TabIndex        =   19
            Top             =   435
            Width           =   1250
         End
         Begin VB.OptionButton optColor_Item 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Background"
            Height          =   315
            Index           =   1
            Left            =   3300
            TabIndex        =   18
            Top             =   150
            Width           =   1250
         End
         Begin VB.Label lblClrlab 
            BackColor       =   &H00C0C0C0&
            Caption         =   "This point (1 to 2)"
            Height          =   240
            Left            =   4845
            TabIndex        =   47
            Top             =   405
            Width           =   1395
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Remember:  One dimensional array properties are indexed by ThisPoint and NOT ThisSet.  eg. ( ColorData )"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   135
            TabIndex        =   74
            Top             =   1095
            Width           =   7665
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveBitmap 
         Caption         =   "Save Graph as &Bitmap"
      End
      Begin VB.Menu mnuFileSaveMetafile 
         Caption         =   "Save Graph as &Metafile"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print Graph"
      End
      Begin VB.Menu FileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear Graph"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmGraphApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuHelpAbout_Click()
  
  About.Show

End Sub

Private Sub cboAx_Props_Click(Index As Integer)

'-------------------------------------------------------
'  Here, axis properties are set to the combo boxes
'  Listindex property when the selected item in the
'  combo box gets changed.
'-------------------------------------------------------

  Select Case Index
    Case 0
      '-------------------------------------------------
      ' Change the Axis position  (left, right, default)
      '-------------------------------------------------
      mdiGraph.graSample.YAxisPos = cboAx_Props(0).ListIndex
    Case 1
      '-------------------------------------------------
      ' Change the Axis style
      '   (variable, user defined or default origin
      '-------------------------------------------------
      mdiGraph.graSample.YAxisStyle = cboAx_Props(1).ListIndex
    Case 2
      '-------------------------------------------------
      ' Select from no ticks, all ticks, x tick only
      '   y ticks only.
      '-------------------------------------------------
      mdiGraph.graSample.Ticks = cboAx_Props(2).ListIndex
  End Select
  '-----------------------------------------------------
  ' If this event is not intentionally skipped then
  ' draw the graph
  '-----------------------------------------------------
  If Not blnSkipEvent Then
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub chkPtln_Click()
  
'-------------------------------------------------------
' Set the patterned lines property to either true or
' false dependeing on the value of the check box
' called chkPtln.  Then draw the graph.
'-------------------------------------------------------

  mdiGraph.graSample.PatternedLines = chkPtln.Value
  mdiGraph.graSample.DrawMode = 2

End Sub

Private Sub chkThln_Click()
  
'-------------------------------------------------------
' Set the thick lines property to either true or
' false dependeing on the value of the check box
' called chkThln.  Then draw the graph, if the event
' was not chosen to be skipped.
'-------------------------------------------------------

  mdiGraph.graSample.ThickLines = chkThln.Value
  If Not blnSkipEvent Then
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub mnuEditClear_Click()
  
'-------------------------------------------------------
' Clear the graph
'-------------------------------------------------------

  mdiGraph.graSample.DrawMode = 1     'clear the graph
  mdiGraph.graSample.DataReset = 9    'reset all data in
                                     'the graph
  mdiGraph.graSample.RandomData = 0   'set randomdata in
                                     'the graph to
                                     'off

End Sub

Private Sub hsbColorBar_Change()

'-------------------------------------------------------
' Set colors of individual points or sets depending on
' the type of graph.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' Since the Colordata array property is a single
  ' dimension array, it is indexed by THISPOINT NOT
  ' THISSET.  Therefore, we need to set thispoint to
  ' the value in hsbColorBar so we change the correct
  ' color.
  '-----------------------------------------------------
  mdiGraph.graSample.ThisPoint = hsbColorBar.Value
  '-----------------------------------------------------
  ' If the colordata option is set, then let's                need to look at some more!!!!!
  'If (optColor_Item(3).Value) Then
  '  X% = mdiGraph.graSample.ColorData
  'End If

End Sub

Private Sub mnuEditCopy_Click()
  
'-------------------------------------------------------
' Copy the graph to the clipboard by setting drawmode
' equal to four.  Then inform the user the graph was
' copied
'-------------------------------------------------------

  mdiGraph.graSample.DrawMode = 4
  MsgBox "Graph was copied to the clipboard"

End Sub

Private Sub Draw_ColorPalette()

'-------------------------------------------------------
'  Here, we draw the color palette which is used to
'  select various colors for the graph.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' Since the colors the graph control accepts are
  ' identical for QBcolors, it works out very nicely.
  '-----------------------------------------------------

  Dim intRows As Integer
  Dim intCols As Integer
  Dim intRowoffset As Integer
  Dim intColoffset As Integer

  Dim intX As Integer
  Dim intY As Integer
  Dim intXPos As Integer
  Dim intYPos As Integer
  
  Dim lngColor As Long
  
  intRows = 2  'Number of rows in the palette
  intCols = 8  'Number of columns in the palette

  intRowoffset = 5
  intColoffset = 5
  
  For intX = 0 To intRows - 1
    For intY = 0 To intCols - 1
      '-------------------------------------------------
      ' Get the positions for the box to be drawn.
      '-------------------------------------------------
      intXPos = (intY * 15) + ((intY + 1) * intColoffset)
      intYPos = (intX * 15) + ((intX + 1) * intRowoffset)
      '-------------------------------------------------
      ' Choose the color for the box to be drawn.
      '-------------------------------------------------
      lngColor = QBColor((intX * intCols) + (intY))
      '-------------------------------------------------
      ' Draw the actual color box.
      '-------------------------------------------------
      picColorPallete.Line (intXPos, intYPos)-(intXPos + 15, intYPos + 15), lngColor, BF
    Next intY
  Next intX

End Sub



Private Sub cboEnumArrays_Click(Index As Integer)
  
'-------------------------------------------------------
' This event fills in some of the enumerated array
' properties for the graph control.  Then it draws
' the graph, if the event is not being skipped.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' Since the Extradata, Symboldata, and Patterndata
  ' array properties are single dimension arrays, they
  ' are indexed by THISPOINT NOT THISSET.  Therefore, we
  ' need to set thispoint to the value in hsbPositionBar(2)
  ' so we change the correct part of the garph.
  '-----------------------------------------------------
  mdiGraph.graSample.ThisPoint = hsbPositionBar(2).Value
  Select Case Index
    Case 0
      '-------------------------------------------------
      ' Change any extra data for the graph
      ' (Explode/ Not Explode for pie charts)
      ' (Specify the color of sids for 3d bar charts)
      '-------------------------------------------------
      mdiGraph.graSample.ExtraData = cboEnumArrays(Index).ListIndex
    Case 1
      '-------------------------------------------------
      ' Set syboldata for line, log/lin, scatter, and
      ' polar graphs.
      '-------------------------------------------------
      mdiGraph.graSample.SymbolData = cboEnumArrays(Index).ListIndex
    Case 2
      '-------------------------------------------------
      ' Select a pattern for solid fills, a line pattern
      ' for patterened lines, or a line width for thick
      ' lines.
      '-------------------------------------------------
      mdiGraph.graSample.PatternData = CInt(cboEnumArrays(Index).Text)
  End Select
  '-----------------------------------------------------
  ' Draw the graph if we have chosen not to skip this
  ' event, otherwise set the blnSkipEvent flag to false
  '-----------------------------------------------------
  If (Not blnSkipEvent) Then
    mdiGraph.graSample.DrawMode = 2
  Else
    blnSkipEvent = False
  End If

End Sub

Private Sub mnuFileExit_Click()
  
'-----------------------------------------------------
' End the program!
'-----------------------------------------------------
  
  End

End Sub

Private Sub cboFontInfo_Click(Index As Integer)
  
'-------------------------------------------------------
' In this event we choose which fonts to use with
' each type of text in the graph.  Then if the event is
' not being skipped draw the graph.
'-------------------------------------------------------

Static dontdo As Boolean

  Select Case Index
    '---------------------------------------------------
    ' Select which font attribute of the graph we are
    ' going to modify.  Done by setting fontuse equal
    ' to the listindex of the combo box. ie the current
    ' selected font use.
    '---------------------------------------------------
    Case 0
      mdiGraph.graSample.FontUse = cboFontInfo(Index).ListIndex
      '-------------------------------------------------
      ' Set dontdo to true to prevent graph redrawing
      ' when we change the values of the cboFontInfo
      ' combo boxes for font family and font style.
      '-------------------------------------------------
      dontdo = True
      cboFontInfo(1).ListIndex = mdiGraph.graSample.FontFamily
      cboFontInfo(2).ListIndex = mdiGraph.graSample.FontStyle
      '-------------------------------------------------
      ' Set back to false since we are done changing
      '-------------------------------------------------
      dontdo = False
      txtFontsize.Text = Str$(mdiGraph.graSample.FontSize)
    Case 1
      '-------------------------------------------------
      ' Change the FontFamily by setting it equal to
      ' the current family selected in the combo box
      ' by setting it equal to the combo box's listindex
      '-------------------------------------------------
      mdiGraph.graSample.FontFamily = cboFontInfo(Index).ListIndex
    Case 2
      '-------------------------------------------------
      ' Change the FontStyle by setting it equal to
      ' the current Style selected in the combo box
      ' by setting it equal to the combo box's listindex
      '-------------------------------------------------
      mdiGraph.graSample.FontStyle = cboFontInfo(Index).ListIndex
    Case 3
      '-------------------------------------------------
      ' Set the Fontsize to whatever the user types in.
      '-------------------------------------------------
      mdiGraph.graSample.FontSize = CInt(txtFontsize.Text)
  End Select
  '-----------------------------------------------------
  ' If this event is not skipped then draw the graph
  '-----------------------------------------------------
  If Not blnSkipEvent Then
    If Not dontdo Then
      mdiGraph.graSample.DrawMode = 2
    End If
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'''If the user double clicks on the control box minimize the form
'''instead of closing it
  If UnloadMode = 0 Then
    Cancel = True
    Me.WindowState = 1
  End If

End Sub

Private Sub optGridP_Click(Index As Integer)
  
'-------------------------------------------------------
' In this event we set the grid style equal to the index
' of the currently selected grid style option button.
' Then we draw the graph by setting drawmode equal to
' two.
'-------------------------------------------------------

  mdiGraph.graSample.GridStyle = Index
  mdiGraph.graSample.DrawMode = 2

End Sub

Private Sub cboGRTS_Click(Index As Integer)
  
'-------------------------------------------------------
' In this event we do the following....
'
'  If a graphtype was chosen we have to reset the
'  items in the graph style combo box.  If a graph
'  style was chosen, we set the graph style to the
'  one currently selected in the graph style combo box
'  by setting the graph style property equal to the
'  combo box's listindex.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' If the GraphType property was chosen then...
  '-----------------------------------------------------
  If (Index = 0) Then
    '---------------------------------------------------
    ' Set the graph type equal to the listindex of the
    ' combo box and then depending on the list indexes
    ' value, reload the graph style combo box with new
    ' settings.
    '---------------------------------------------------
    mdiGraph.graSample.GraphType = cboGRTS(0).ListIndex
    Select Case cboGRTS(0).ListIndex
      '-------------------------------------------------
      ' No graph
      '-------------------------------------------------
      Case 0
        cboGRTS(1).Clear
      
      '-------------------------------------------------
      ' Load in Pie graph styles
      '-------------------------------------------------
      Case 1, 2
        Load_PieStylesList
      
      '-------------------------------------------------
      ' Load in Bar Graph Styles, 2D
      '-------------------------------------------------
      Case 3
        Load_BarStylesList 0
      
      '-------------------------------------------------
      ' Load in Bar Graph Styles, 3D
      '-------------------------------------------------
      Case 4
        Load_BarStylesList 1
      
      '-------------------------------------------------
      ' Load in Gantt Chart Styles
      '-------------------------------------------------
      Case 5
        Load_GantStylesList
      
      '-------------------------------------------------
      ' Load in Line, Log/Lin, Polar Graph Styles
      '-------------------------------------------------
      Case 6, 7, 10
        Load_LinLogPolStylesList
      
      '-------------------------------------------------
      ' Load in Area Graph Styles
      '-------------------------------------------------
      Case 8
        Load_AreaStylesList
      
      '-------------------------------------------------
      ' Load in Scatter Graph Styles
      '-------------------------------------------------
      Case 9
        Load_ScatterStylesList
      
      '-------------------------------------------------
      ' Load in High Low Close Graph Styles
      '-------------------------------------------------
      Case 11
        Load_HLCStylesList
    End Select
  Else
    '---------------------------------------------------
    ' If the GraphStyle combo box was selected, set the
    ' graph style equal to it's combo box's listindex
    '---------------------------------------------------
    mdiGraph.graSample.GraphStyle = cboGRTS(1).ListIndex
  End If
  '-----------------------------------------------------
  ' Make sure we have the correct data in the enumerated
  ' property array combo boxes.
  '-----------------------------------------------------
  Dim intCurrentIndex As Integer
  
  intCurrentIndex = cboGRTS(0).ListIndex
  Load_ExtraDataList intCurrentIndex
  Load_PatternDataList intCurrentIndex
  Load_SymbolDataList intCurrentIndex
  
  '-----------------------------------------------------
  ' If we are not skipping the event, then draw the
  ' graph, by setting drawmode equal to two.
  '-----------------------------------------------------
  If (Not blnSkipEvent) Then
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub cboLineStats_Click()
  
'-------------------------------------------------------
' In this event we change the Lines stats property of
' the graph control to the listindex of the cboLineStats
' combo box.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' Change the cboLineStats property to whatever was
  ' selected in the combo box.
  '-----------------------------------------------------
  mdiGraph.graSample.LineStats = cboLineStats.ListIndex
  '-----------------------------------------------------
  ' If this event is not skipped, then draw the graph.
  '-----------------------------------------------------
  If Not blnSkipEvent Then
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub Load_AreaStylesList()

'-------------------------------------------------------
'  Subroutine to Load the Area Graph Styles List
'  Called when the Graph type becomes an area graph
'  cboGRTS--GRaphTypeStyles(1)
'-------------------------------------------------------
  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"
  cboGRTS(1).AddItem "Absolute"
  cboGRTS(1).AddItem "Percentage"

End Sub

Private Sub Load_BarStylesList(intBarType As Integer)

'-------------------------------------------------------
'  Subroutine to Load Bar Styles List
'  Called when the Graph Type becomes a bar graph
'  cboGRTS--GRaphTypeStyles(1)
'-------------------------------------------------------
  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"
  cboGRTS(1).AddItem "Horizontal"
  cboGRTS(1).AddItem "Stacked"
  cboGRTS(1).AddItem "Horizontal Stacked"
  cboGRTS(1).AddItem "Stacked %"
  cboGRTS(1).AddItem "Horizontal Stacked %"
  '-----------------------------------------------------
  ' If the bartype is 3D then we have a few other styles
  '-----------------------------------------------------
  If (intBarType = 1) Then
    cboGRTS(1).AddItem "Z Clustered"
    cboGRTS(1).AddItem "Horizontal Z Clustered"
  End If

End Sub

Private Sub Load_ExtraDataList(intGraphType As Integer)

'-------------------------------------------------------
'  Subroutine to fill the Extra data list
'  Called when the Graph Type changes
'  Enumerated Arrays--cboEnumArrays(0)
'-------------------------------------------------------
  cboEnumArrays(0).Clear
  Select Case intGraphType
    '---------------------------------------------------
    '  If the GraphType is a pie chart then load the
    '  extra data list with these options...
    '---------------------------------------------------
    Case 1, 2
      cboEnumArrays(0).Enabled = True
      cboEnumArrays(0).AddItem "(Default) Not Exploded"
      cboEnumArrays(0).AddItem "Exploded"
    '---------------------------------------------------
    '  If the GraphType is a 3D bar then load the
    '  extra data list with these options...
    '---------------------------------------------------
    Case 4
      cboEnumArrays(0).Enabled = True
      cboEnumArrays(0).AddItem "(Default) Black"
      cboEnumArrays(0).AddItem "Blue"
      cboEnumArrays(0).AddItem "Green"
      cboEnumArrays(0).AddItem "Cyan"
      cboEnumArrays(0).AddItem "Red"
      cboEnumArrays(0).AddItem "Magenta"
      cboEnumArrays(0).AddItem "Brown"
      cboEnumArrays(0).AddItem "Light Gray"
      cboEnumArrays(0).AddItem "Dark Gray"
      cboEnumArrays(0).AddItem "Light Blue"
      cboEnumArrays(0).AddItem "Light Green"
      cboEnumArrays(0).AddItem "Light Cyan"
      cboEnumArrays(0).AddItem "Light Red"
      cboEnumArrays(0).AddItem "Light Magenta"
      cboEnumArrays(0).AddItem "Yellow"
      cboEnumArrays(0).AddItem "White"
    '---------------------------------------------------
    '  If the GraphType is anything else then
    '  disable the extra data list...
    '---------------------------------------------------
    Case Else
      cboEnumArrays(0).Enabled = False
  End Select

End Sub

Private Sub Load_cboFontInfoList()

'-------------------------------------------------------
'  Subroutine to fill the font information List
'  Called when the Application begins
'-------------------------------------------------------
  
  '-----------------------------------------------------
  'This combo box designates what the font is used for..
  '-----------------------------------------------------
  cboFontInfo(0).AddItem "(Default) Graph Title"
  cboFontInfo(0).AddItem "Other txtTitles"
  cboFontInfo(0).AddItem "Labels"
  cboFontInfo(0).AddItem "Legend"
  cboFontInfo(0).AddItem "All Text"

  '-----------------------------------------------------
  'This combo box designates what font family to apply
  '-----------------------------------------------------
  cboFontInfo(1).AddItem "(Default) Roman"
  cboFontInfo(1).AddItem "Swiss"
  cboFontInfo(1).AddItem "Modern"

  '-----------------------------------------------------
  'This combo box designates what font style to apply
  '-----------------------------------------------------
  cboFontInfo(2).AddItem "(Default)"
  cboFontInfo(2).AddItem "Italic"
  cboFontInfo(2).AddItem "Bold"
  cboFontInfo(2).AddItem "Bold Italic"
  cboFontInfo(2).AddItem "Underlined"
  cboFontInfo(2).AddItem "Underlined Italic"
  cboFontInfo(2).AddItem "Underlined Bold"
  cboFontInfo(2).AddItem "Underlined Bold Italic"

End Sub

Private Sub Load_GantStylesList()

'-------------------------------------------------------
'  Subroutine to fill the gant style List
'  Called when the Graph Type becomes a gant style graph
'-------------------------------------------------------
  
  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"
  cboGRTS(1).AddItem "Spaced Bars"

End Sub

Private Sub Load_GraphTypesList()

'-------------------------------------------------------
'  Subroutine to fill the Graph Types list
'  Called when the applicationi begins
'  Graph Types and Styles   cboGRTS
'-------------------------------------------------------

  cboGRTS(0).AddItem "None", 0
  cboGRTS(0).AddItem "2d Pie", 1
  cboGRTS(0).AddItem "3d Pie", 2
  cboGRTS(0).AddItem "2d Bar", 3
  cboGRTS(0).AddItem "3d Bar", 4
  cboGRTS(0).AddItem "Gantt", 5
  cboGRTS(0).AddItem "Line", 6
  cboGRTS(0).AddItem "Log/Lin", 7
  cboGRTS(0).AddItem "Area", 8
  cboGRTS(0).AddItem "Scatter", 9
  cboGRTS(0).AddItem "Polar", 10
  cboGRTS(0).AddItem "HLC", 11
  


End Sub

Private Sub Load_HLCStylesList()

'-------------------------------------------------------
'  Subroutine to fill the High Low Close styles list
'  Called when the graph type changes to HLC
'  Graph Types and Styles   cboGRTS(1) used for styles
'-------------------------------------------------------

  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"
  cboGRTS(1).AddItem "No Close Bar"
  cboGRTS(1).AddItem "No High-Low Bars"
  cboGRTS(1).AddItem "No Bars"

End Sub

Private Sub Load_cboLineStatsList()

'-------------------------------------------------------
'  Subroutine to fill the Line Stats list
'  Called when the application begins
'  cboLineStats, a combo box containing the stats
'-------------------------------------------------------

  cboLineStats.AddItem "None"
  cboLineStats.AddItem "Mean"
  cboLineStats.AddItem "Minimum and Maximum"
  cboLineStats.AddItem "Mean and Minimum and Maximum"
  cboLineStats.AddItem "Standard Deviation"
  cboLineStats.AddItem "Standard Deviation and Mean."
  cboLineStats.AddItem "Standard Deviation and Minimum and Maximum"
  cboLineStats.AddItem "Standard Deviation and Minimum and Maximum and Mean"
  cboLineStats.AddItem "Best Fit"
  cboLineStats.AddItem "Best Fit and Mean"
  cboLineStats.AddItem "Best Fit and Minimum and Maximum"
  cboLineStats.AddItem "Best Fit and Minimum and Maximum and Mean"
  cboLineStats.AddItem "Best Fit and Standard Deviation"
  cboLineStats.AddItem "Best Fit and Standard Deviation and Mean"
  cboLineStats.AddItem "Best Fit and Standard Deviation and Minimum and Maximum"
  cboLineStats.AddItem "All"

End Sub

Private Sub Load_LinLogPolStylesList()

'-------------------------------------------------------
'  Subroutine to fill the Line Styles list
'  Called when the graph type becomes Line, log, Polar
'  cboGRTS(1), used for styles
'-------------------------------------------------------

  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"
  cboGRTS(1).AddItem "Symbols"
  cboGRTS(1).AddItem "Sticks"
  cboGRTS(1).AddItem "Sticks and Symbols"
  cboGRTS(1).AddItem "Lines"
  cboGRTS(1).AddItem "Lines and Symbols"
  cboGRTS(1).AddItem "Lines and Sticks"
  cboGRTS(1).AddItem "Lines and Sticks and Symbols"

End Sub

Private Sub Load_PatternDataList(intGraphType As Integer)
  
'-------------------------------------------------------
'  Subroutine to fill the Pattern Data list
'  Called when the application begins
'  cboLineStats, a combo box containing the stats
'  cboEnumArrays(2) is used for the Pattern Data list
'-------------------------------------------------------
  Dim intX As Integer
  
  Select Case intGraphType
    '---------------------------------------------------
    ' If the graph type is Line, Log, Polar or HLC then
    ' use these values for pattern data.
    '---------------------------------------------------
    Case 6, 7, 10, 11
      cboEnumArrays(2).Clear
      '-------------------------------------------------
      ' If the graph uses thick lines then the patterns
      ' are slightly different.
      '-------------------------------------------------
      If (mdiGraph.graSample.ThickLines) Then
        cboEnumArrays(2).AddItem "1"
        cboEnumArrays(2).AddItem "2"
        cboEnumArrays(2).AddItem "3"
        cboEnumArrays(2).AddItem "4"
        cboEnumArrays(2).AddItem "5"
      Else
        cboEnumArrays(2).AddItem "0"
        cboEnumArrays(2).AddItem "1"
        cboEnumArrays(2).AddItem "2"
        cboEnumArrays(2).AddItem "3"
        cboEnumArrays(2).AddItem "4"
      End If
      cboEnumArrays(2).Enabled = True
    '---------------------------------------------------
    ' If the graph type is Pie, Bar, Gantt, or Area then
    ' use these values for pattern data.
    '---------------------------------------------------
    Case 1 To 5, 8
      cboEnumArrays(2).Clear
      For intX = 0 To 7
        cboEnumArrays(2).AddItem Str$(intX)
      Next intX
      For intX = 16 To 31
        cboEnumArrays(2).AddItem Str$(intX)
      Next intX
      cboEnumArrays(2).Enabled = True
    '---------------------------------------------------
    ' With anything else just disable the pattern data
    ' list
    '---------------------------------------------------
    Case Else
      cboEnumArrays(2).Enabled = False
  End Select

End Sub

Private Sub Load_PieStylesList()

'-------------------------------------------------------
'  Subroutine to fill the Pie Styles list
'  Called when the graph type changes to pie
'  cboGRTS(1) used for Styles
'-------------------------------------------------------

  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"
  cboGRTS(1).AddItem "No Label Lines"
  cboGRTS(1).AddItem "Colored Labels"
  cboGRTS(1).AddItem "Colored Labels without Lines"
  cboGRTS(1).AddItem "% Labels"
  cboGRTS(1).AddItem "% Labels without Lines"
  cboGRTS(1).AddItem "% Colored Labels"
  cboGRTS(1).AddItem "% Colored Labels without Lines"

End Sub

Private Sub Load_cboPrintStyleList()

'-------------------------------------------------------
'  Subroutine to fill the Print Styles list
'  Called when the application begins
'  cboPrintStyle, a combo box containing print styles
'-------------------------------------------------------

  cboPrintStyle.AddItem "Default Monochrome"
  cboPrintStyle.AddItem "Color"
  cboPrintStyle.AddItem "Monochrome with border"
  cboPrintStyle.AddItem "Color with border."

End Sub

Private Sub Load_ScatterStylesList()

'-------------------------------------------------------
'  Subroutine to fill the Scatter Styles list
'  Called when the graph type changes to Scatter plot
'  cboGRTS(1) contains graph styles
'-------------------------------------------------------

  cboGRTS(1).Clear
  cboGRTS(1).AddItem "Default"

End Sub

Private Sub Load_SymbolDataList(intGraphType As Integer)

'-------------------------------------------------------
'  Subroutine to fill the Symbol Data list
'  Called when the application begins
'  cboEnumArrays(1) contains Symbol Data
'-------------------------------------------------------

Static blnPerformedFill As Boolean   'used to prevent
                                   'multiple fills.
  
  Select Case (intGraphType)
    '---------------------------------------------------
    ' If the graph type is Line, Log, Scatter, or Polar
    ' then fill the list and enable it.
    '---------------------------------------------------
    Case 6, 7, 9, 10
      cboEnumArrays(1).Enabled = True
      If (Not (blnPerformedFill)) Then
        cboEnumArrays(1).AddItem "+"
        cboEnumArrays(1).AddItem "x"
        cboEnumArrays(1).AddItem "Empty triangle up"
        cboEnumArrays(1).AddItem "Filled triangle up"
        cboEnumArrays(1).AddItem "Empty triangle down"
        cboEnumArrays(1).AddItem "Filled triangle down"
        cboEnumArrays(1).AddItem "Empty Square"
        cboEnumArrays(1).AddItem "Filled Square"
        cboEnumArrays(1).AddItem "Empty Diamond"
        cboEnumArrays(1).AddItem "Filled Diamond"
        blnPerformedFill = True
      End If
    '--------------------------------------------------
    ' If any other graph type than symbol data isn't
    ' useful, so just disable the combo box.
    '--------------------------------------------------
    Case Else
      cboEnumArrays(1).Enabled = False
  End Select

End Sub

Private Sub Load_TicksList()

'-------------------------------------------------------
'  Subroutine to fill the Tick list
'  Called when the application begins
'  cboAx_Props(2) contains Axis tick info
'-------------------------------------------------------

  cboAx_Props(2).AddItem "Default (off)"
  cboAx_Props(2).AddItem "On"
  cboAx_Props(2).AddItem "X Ticks"
  cboAx_Props(2).AddItem "Y Ticks"

End Sub

Private Sub Load_YAxisPosList()

'-------------------------------------------------------
'  Subroutine to fill the Y Axis Position list
'  Called when the application begins
'  cboAx_Props(0) contains axis position information
'-------------------------------------------------------

  cboAx_Props(0).AddItem "Default"
  cboAx_Props(0).AddItem "Left"
  cboAx_Props(0).AddItem "Right"

End Sub

Private Sub Load_YAxisStyleList()

'-------------------------------------------------------
'  Subroutine to fill the Y Axis Origin list
'  Called when the application begins
'  cboAx_Props(1) contains axis origin information
'-------------------------------------------------------

  cboAx_Props(1).AddItem "Default"
  cboAx_Props(1).AddItem "Variable Origin"
  cboAx_Props(1).AddItem "User Defined Origin"

End Sub

Private Sub MDIForm_Load()

Dim intX As Integer

'-------------------------------------------------------
'  Form Load event, handle things which we need done
'  prior to the user using the application.
'-------------------------------------------------------

  '-------------------------------------------------------
  'Bring the form up maximized.
  '-------------------------------------------------------
  frmGraphApp.WindowState = 2
  
  '-------------------------------------------------------
  'Set the height of our tabs...
  '-------------------------------------------------------
  PicTabContainer.Height = 2200
  intX = DoEvents()
  
  '-------------------------------------------------------
  'Load the child form which contains the graph
  '-------------------------------------------------------
  Load mdiGraph

  '-------------------------------------------------------
  'Set up the tabs on the MDI form and draw
  '-------------------------------------------------------
  SetupTabs Me, 6

  '-------------------------------------------------------
  'Show the MDI Parent form
  '-------------------------------------------------------
  frmGraphApp.Show

  '-------------------------------------------------------
  ' Now Perform some functions which need to be done
  ' before using the application.

  '-------------------------------------------------------
  Draw_ColorPalette
  Load_cboFontInfoList
  Load_GraphTypesList
  intX = mdiGraph.graSample.GraphType
  Load_ExtraDataList intX
  Load_PatternDataList intX
  Load_SymbolDataList intX
  Load_YAxisPosList
  Load_YAxisStyleList
  Load_TicksList
  Load_cboLineStatsList
  Load_cboPrintStyleList
  
  '-------------------------------------------------------
  ' Since the application starts off with the graph
  ' using random data, we need all of the settings to
  ' reflect whatever the graph is showing.
  '-------------------------------------------------------
  For intX = 0 To 5
    Realize_GraphSettings intX
  Next intX

  intX = DoEvents()
  '-------------------------------------------------------
  ' Now show the tabs
  '-------------------------------------------------------
  PicTabContainer.Visible = True

End Sub

Private Sub optPalt_Click(Index As Integer)
  
'-------------------------------------------------------
' In this event we select which palette we are using for
' the graph.  (Solid[0], Pastel[1], Grayscale[2])
'-------------------------------------------------------

  mdiGraph.graSample.Palette = Index
  '-----------------------------------------------------
  'Now draw the graph.
  '-----------------------------------------------------
  mdiGraph.graSample.DrawMode = 2

End Sub

Private Sub PicTabContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'-------------------------------------------------------
'  When the user clicks on a new tab, the new tab needs
'  to be drawn.
'-------------------------------------------------------
On Error GoTo Error_Handler

  fratabs(DrawTabs(Me, X, Y) - 1).ZOrder

Error_Handler:
  Exit Sub

End Sub

Private Sub PicTabContainer_Resize()
  SetupTabs Me, 6
End Sub

Private Sub picColorPallete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim intRemainder As Integer
Dim intRow As Integer
Dim intColumn As Integer
Dim intIndex As Integer

'-------------------------------------------------------
'  This event determines which color was selected when
'  the user clicks on the picture.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' Make sure the user clicked in the color region, not
  ' on the outer edges of the picture.
  '-----------------------------------------------------
  If (Y < 40) And (Y > 5) And (X < 160) And (X > 5) Then
      '-------------------------------------------------
      ' See if the user actually clicked on a color area
      ' of the picture.  If intRemainder is less than 5
      ' then the user clicked in whitespace between the
      ' color regions.  (HORIZONTALY)
      '-------------------------------------------------
      intRemainder = X Mod 20
      
      If (intRemainder > 5) Then
        '-------------------------------------------------
        ' See if the user actually clicked on a color area
        ' of the picture.  If intRemainder is less than 5
        ' then the user clicked in whitespace between the
        ' color regions.  (VERTICALLY)
        '-------------------------------------------------
        intRemainder = Y Mod 20
        If (intRemainder > 5) Then
          '-----------------------------------------------
          ' Determine exactly which block was chosen.
          '-----------------------------------------------
          intRow = Y \ 20
          intColumn = X \ 20
          intIndex = ((intRow * 8) + intColumn)
        End If
      End If
      '---------------------------------------------------
      'Now find out which color item we are dealing with
      'and set it to the index we calculated.
      '---------------------------------------------------
      If (optColor_Item(1).Value = True) Then
        '-------------------------------------------------
        ' Set the background color to the index
        '-------------------------------------------------
        mdiGraph.graSample.Background = intIndex
      Else
        If (optColor_Item(2).Value = True) Then
          '-------------------------------------------------
          ' Set the Foreground color to the index
          '-------------------------------------------------
          mdiGraph.graSample.Foreground = intIndex
        Else
          '-------------------------------------------------
          ' Set the Colordata to the calculated index.
          '-------------------------------------------------
          mdiGraph.graSample.ColorData = intIndex
        End If
      End If
      '-----------------------------------------------------
      ' Draw the graph.
      '-----------------------------------------------------
      mdiGraph.graSample.DrawMode = 2
    End If

End Sub




Private Sub txtPointData_KeyPress(Index As Integer, KeyAscii As Integer)
  
'-------------------------------------------------------
'  Set values of various properties of the graph on the
'  pressing of <Enter>.
'-------------------------------------------------------
Dim intPosition As Integer
  
  If (KeyAscii = 13) Then
    If ((Index = 0) Or (Index = 1)) Then
      hsbPositionBar(Index).Max = CInt(txtPointData(Index).Text)
      If (CSng(txtPointData(0).Text) >= CSng(txtPointData(1).Text)) Then
        '-------------------------------------------------------
        ' Note, the colordata scrollbar needs resized because
        '       it is indexed by the ThisPoint property.
        '-------------------------------------------------------
        hsbColorBar.Max = CSng(txtPointData(0).Text)
        hsbColorBar.Value = mdiGraph.graSample.ThisPoint
        '-------------------------------------------------------
        ' hsbPositionBar(1), used for the number of points needs to
        ' be set to the NumPoints property of the graph.
        '-------------------------------------------------------
        hsbPositionBar(1).Max = CInt(txtPointData(0).Text)
        hsbPositionBar(1).Value = mdiGraph.graSample.ThisPoint
        '-------------------------------------------------------
        ' hsbPositionBar(2), used to index the single dimension
        ' enumerated arrays needs it's max value set to the
        ' value of NumPoints.
        '-------------------------------------------------------
        hsbPositionBar(2).Max = CSng(txtPointData(0).Text)
      Else
        '-------------------------------------------------------
        ' If the number of points are not greater than the
        ' number of sets then we set the scrollbar max values
        ' differently.  The code looks weird, but this is the
        ' way the control works!
        '-------------------------------------------------------
        hsbPositionBar(1).Max = CSng(txtPointData(1).Text)
        hsbPositionBar(2).Max = CSng(txtPointData(1).Text)
      End If
      '---------------------------------------------------------
      '  Now set the captions for the scrollbars so they report
      '  the correct values.
      '---------------------------------------------------------
      intPosition = InStr(lblPlab(Index).Caption, "(")
      lblPlab(Index).Caption = Left$(lblPlab(Index).Caption, intPosition - 1)
      lblPlab(Index).Caption = lblPlab(Index).Caption & "( 1 to " & txtPointData(Index).Text & " )"
      mdiGraph.graSample.NumSets = CInt(txtPointData(0).Text)
      mdiGraph.graSample.NumPoints = CInt(txtPointData(1).Text)
      Label4.Caption = lblPlab(1).Caption
    Else
      '---------------------------------------------------------
      ' If we are not dealing with Numpoints or Numsets then we
      ' are dealing with one of the following...
      '---------------------------------------------------------
      mdiGraph.graSample.RandomData = 0  'turn off so we can add data
      Select Case Index
        Case 5
          '-----------------------------------------------------
          ' Set the value of the graph for the current point and
          ' the current set.
          '-----------------------------------------------------
          mdiGraph.graSample.ThisPoint = hsbPositionBar(1).Value
          mdiGraph.graSample.GraphData = CSng(txtPointData(Index).Text)
        Case 2
          '-----------------------------------------------------
          ' Set the label value for the current point.
          '-----------------------------------------------------
          mdiGraph.graSample.ThisPoint = hsbPositionBar(2).Value
          mdiGraph.graSample.LabelText = txtPointData(2).Text
        Case 3
          '-----------------------------------------------------
          ' Set the string for the legend.
          '-----------------------------------------------------
          mdiGraph.graSample.ThisPoint = hsbPositionBar(2).Value
          mdiGraph.graSample.LegendText = txtPointData(3).Text
       End Select
    End If
    '-----------------------------------------------------------
    ' Draw the graph.
    '-----------------------------------------------------------
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub hsbPositionBar_Change(Index As Integer)

'-------------------------------------------------------
'  Set the current point and set of the graph using the
'  value of the hsbPositionBar scrollbars.
'-------------------------------------------------------

  '-----------------------------------------------------
  ' This must be done in order to prevent some very
  ' eratic behavior.
  '-----------------------------------------------------
  hsbPositionBar(Index).Enabled = 0
  mdiGraph.graSample.ThisSet = hsbPositionBar(0).Value
  '-----------------------------------------------------
  ' Depending on which position bar was changed,
  ' Thispoint will get a different value.
  '-----------------------------------------------------
  If (Index = 1) Or (Index = 0) Then
    mdiGraph.graSample.ThisPoint = hsbPositionBar(1).Value
  Else
    mdiGraph.graSample.ThisPoint = hsbPositionBar(2).Value
  End If
  i% = DoEvents()
  
  '-----------------------------------------------------
  ' Show the user the value of Xposdata and GraphData
  ' for the point in which they selected.
  '-----------------------------------------------------
  txtPointData(4).Text = CStr(mdiGraph.graSample.XPosData)
  txtPointData(5).Text = CStr(mdiGraph.graSample.GraphData)
  
  '-----------------------------------------------------
  ' Re-enable the position bar scrollbar and set focus
  ' to it.
  '-----------------------------------------------------
  hsbPositionBar(Index).Enabled = -1
'  hsbPositionBar(Index).SetFocus

End Sub

Private Sub mnuFilePrint_Click()

'-------------------------------------------------------
'  Send the graph to the printer
'-------------------------------------------------------

  mdiGraph.graSample.DrawMode = 5
  
End Sub

Private Sub cboPrintStyle_Click()

'-------------------------------------------------------
'  Set the graphs cboPrintStyle to the listindex property
'  of the cboPrintStyle combo box.
'-------------------------------------------------------

  mdiGraph.graSample.PrintStyle = cboPrintStyle.ListIndex
  '----------------------------------------------------
  ' If we are not skipping this event then draw the
  ' graph.
  '----------------------------------------------------
  If Not blnSkipEvent Then
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub mnuFileSaveBitmap_Click()

'-------------------------------------------------------
'  Save the current image of the graph as a bitmap
'-------------------------------------------------------
 
Dim strName As String

  '-----------------------------------------------------
  '  Ask the user for the name of the file to save as
  '  until they enter a name.
  '-----------------------------------------------------
  strName = InputBox$("Select the name you would like to save the graph as...")
  Do While strName = ""

    strName = InputBox$("Select the name you would like to save the graph as...")
  Loop
  '-----------------------------------------------------
  ' Now set drawmode to 3, in order to get a bitmap,
  ' set the Imagefile property to the name of the
  ' file to save, and then save the file.  Drawmode = 6
  '-----------------------------------------------------
  mdiGraph.graSample.DrawMode = 3
  mdiGraph.graSample.ImageFile = strName
  mdiGraph.graSample.DrawMode = 6

End Sub

Private Sub mnuFileSaveMetafile_Click()

'-------------------------------------------------------
'  Save the current image of the graph as a bitmap
'-------------------------------------------------------

Dim strName As String
  '-----------------------------------------------------
  '  Ask the user for the name of the file to save as
  '  until they enter a name.
  '-----------------------------------------------------
  strName = InputBox$("Select the name you would like to save the graph as...")
  Do While strName = ""
    strName = InputBox$("Select the name you would like to save the graph as...")
  Loop
  '-----------------------------------------------------
  ' Now set drawmode to 2, in order to get a bitmap,
  ' set the Imagefile property to the name of the
  ' file to save, and then save the file.  Drawmode = 6
  '-----------------------------------------------------
  mdiGraph.graSample.DrawMode = 2
  mdiGraph.graSample.ImageFile = strName
  mdiGraph.graSample.DrawMode = 6

End Sub

Private Sub txtFontsize_KeyPress(KeyAscii As Integer)
  
'-----------------------------------------------------
' Sets the size of the selected type of font for
' the graph.
'-----------------------------------------------------

  If (KeyAscii = 13) Then
    mdiGraph.graSample.FontSize = CInt(txtFontsize.Text)
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub txtTitles_KeyPress(Index As Integer, KeyAscii As Integer)
  
'-----------------------------------------------------
' Upon hiting enter on the txtTitles textboxes, depending
' on the index of the textbox a property will be set
' for the graph
'-----------------------------------------------------

  '---------------------------------------------------
  ' Make sure <Enter> was hit.
  '---------------------------------------------------
  If (KeyAscii = 13) Then
    Select Case Index
      '-----------------------------------------------
      ' Set the text for a title at the bottom of
      ' the graph.
      '-----------------------------------------------
      Case 0
        mdiGraph.graSample.BottomTitle = txtTitles(Index).Text
      '-----------------------------------------------
      ' Set the text for a title of the graph.
      '-----------------------------------------------
      Case 1
        mdiGraph.graSample.GraphTitle = txtTitles(Index).Text
      '-----------------------------------------------
      ' Set the text for a title at the left of
      ' the graph.
      '-----------------------------------------------
      Case 2
        mdiGraph.graSample.LeftTitle = txtTitles(Index).Text
      '-----------------------------------------------
      ' Set the text the caption of the graph.
      '-----------------------------------------------
      Case 3
        mdiGraph.graSample.GraphCaption = txtTitles(Index).Text
    End Select
    '-------------------------------------------------
    ' Draw the graph itself.
    '-------------------------------------------------
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

Private Sub txtYAxv_KeyPress(Index As Integer, KeyAscii As Integer)

'-----------------------------------------------------
' Upon hiting enter on the txtYAxv textboxes, depending
' on the index of the textbox a property will be set
' for the graph
'-----------------------------------------------------
  
  '---------------------------------------------------
  ' Make sure <Enter> was hit.
  '---------------------------------------------------
  If (KeyAscii = 13) Then
    Select Case Index
      '-----------------------------------------------
      ' Set the minimum value for the Y Axis.
      '-----------------------------------------------
      Case 0
        mdiGraph.graSample.YAxisMin = CSng(txtYAxv(Index).Text)
      '-----------------------------------------------
      ' Set the maximum value for the Y Axis.
      '-----------------------------------------------
      Case 1
        mdiGraph.graSample.YAxisMax = CSng(txtYAxv(Index).Text)
      '-----------------------------------------------
      ' Set the number of Y Axis tick marks.
      '-----------------------------------------------
      Case 2
        mdiGraph.graSample.YAxisTicks = CSng(txtYAxv(Index).Text)
      '-----------------------------------------------
      ' Set the value for the interval in which you
      ' want to show ticks.
      '-----------------------------------------------
      Case 3
        mdiGraph.graSample.TickEvery = CSng(txtYAxv(Index).Text)
      '-----------------------------------------------
      ' Set the value for the interval in which you
      ' want to display labels.
      '-----------------------------------------------
      Case 4
        mdiGraph.graSample.LabelEvery = CSng(txtYAxv(Index).Text)
    End Select
    '-------------------------------------------------
    ' Draw the Graph.
    '-------------------------------------------------
    mdiGraph.graSample.DrawMode = 2
  End If

End Sub

