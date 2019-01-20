VERSION 5.00
Begin VB.Form frmCardGameSevens 
   Caption         =   "Card Game Sevens"
   ClientHeight    =   8985
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picKingS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   9120
      Picture         =   "frmJuliaRenSevensGraphical.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   60
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picQueenS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8640
      Picture         =   "frmJuliaRenSevensGraphical.frx":04CB
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   59
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picJackS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8160
      Picture         =   "frmJuliaRenSevensGraphical.frx":09AD
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   58
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picTenS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7680
      Picture         =   "frmJuliaRenSevensGraphical.frx":0E39
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   57
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picNineS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7200
      Picture         =   "frmJuliaRenSevensGraphical.frx":1141
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   56
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picEightS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   6720
      Picture         =   "frmJuliaRenSevensGraphical.frx":141E
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   55
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picKingH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   9120
      Picture         =   "frmJuliaRenSevensGraphical.frx":16E5
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   48
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picQueenH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8640
      Picture         =   "frmJuliaRenSevensGraphical.frx":1BD5
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   47
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picJackH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8160
      Picture         =   "frmJuliaRenSevensGraphical.frx":20E2
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   46
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picTenH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7680
      Picture         =   "frmJuliaRenSevensGraphical.frx":25C3
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   45
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picNineH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7200
      Picture         =   "frmJuliaRenSevensGraphical.frx":28C8
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   44
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picEightH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   6720
      Picture         =   "frmJuliaRenSevensGraphical.frx":2B9A
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   43
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picKingC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   9120
      Picture         =   "frmJuliaRenSevensGraphical.frx":2E45
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   36
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picQueenC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8640
      Picture         =   "frmJuliaRenSevensGraphical.frx":3356
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   35
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picJackC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8160
      Picture         =   "frmJuliaRenSevensGraphical.frx":385F
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   34
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picTenC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7680
      Picture         =   "frmJuliaRenSevensGraphical.frx":3D57
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   33
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picNineC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7200
      Picture         =   "frmJuliaRenSevensGraphical.frx":405E
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   32
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picEightC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   6720
      Picture         =   "frmJuliaRenSevensGraphical.frx":4348
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   31
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picKingD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   9000
      Picture         =   "frmJuliaRenSevensGraphical.frx":461D
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   19
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picQueenD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8520
      Picture         =   "frmJuliaRenSevensGraphical.frx":4AFF
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   20
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picJackD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   8160
      Picture         =   "frmJuliaRenSevensGraphical.frx":4FF6
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   21
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picTenD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7680
      Picture         =   "frmJuliaRenSevensGraphical.frx":54D8
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   22
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picNineD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   7200
      Picture         =   "frmJuliaRenSevensGraphical.frx":5780
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   23
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picEightD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   6720
      Picture         =   "frmJuliaRenSevensGraphical.frx":59F9
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   24
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picSevens 
      AutoSize        =   -1  'True
      Height          =   1515
      Index           =   1
      Left            =   6240
      Picture         =   "frmJuliaRenSevensGraphical.frx":5C58
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   18
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picSevens 
      AutoSize        =   -1  'True
      Height          =   1515
      Index           =   3
      Left            =   6240
      Picture         =   "frmJuliaRenSevensGraphical.frx":5EFD
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   17
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picSevens 
      AutoSize        =   -1  'True
      Height          =   1515
      Index           =   2
      Left            =   6240
      Picture         =   "frmJuliaRenSevensGraphical.frx":6199
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   16
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picSevens 
      AutoSize        =   -1  'True
      Height          =   1515
      Index           =   0
      Left            =   6240
      Picture         =   "frmJuliaRenSevensGraphical.frx":642D
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   15
      Top             =   600
      Width           =   1155
   End
   Begin VB.CommandButton cmdHowToPlay 
      Caption         =   "How To Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlayCard 
      Caption         =   "Play Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   11
      Top             =   8040
      Width           =   1815
   End
   Begin VB.TextBox txtSelectedCard 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   6
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Timer tmrBlinkSevens 
      Interval        =   700
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdStartNewGame 
      Caption         =   "Start New Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.PictureBox picSixD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5760
      Picture         =   "frmJuliaRenSevensGraphical.frx":6671
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   25
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picFiveD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5280
      Picture         =   "frmJuliaRenSevensGraphical.frx":68A0
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   26
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picFourD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4800
      Picture         =   "frmJuliaRenSevensGraphical.frx":6AC2
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   27
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picThreeD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4320
      Picture         =   "frmJuliaRenSevensGraphical.frx":6CD2
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   30
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picTwoD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3840
      Picture         =   "frmJuliaRenSevensGraphical.frx":6EC3
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   28
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picAceD 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3360
      Picture         =   "frmJuliaRenSevensGraphical.frx":7099
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   29
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox picSixC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5760
      Picture         =   "frmJuliaRenSevensGraphical.frx":7260
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   37
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picFiveC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5280
      Picture         =   "frmJuliaRenSevensGraphical.frx":74E8
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   38
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picFourC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4800
      Picture         =   "frmJuliaRenSevensGraphical.frx":7741
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   39
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picThreeC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4320
      Picture         =   "frmJuliaRenSevensGraphical.frx":797E
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   40
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picTwoC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3840
      Picture         =   "frmJuliaRenSevensGraphical.frx":7B99
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   41
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picAceC 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3360
      Picture         =   "frmJuliaRenSevensGraphical.frx":7D95
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   42
      Top             =   2280
      Width           =   1155
   End
   Begin VB.PictureBox picSixH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5760
      Picture         =   "frmJuliaRenSevensGraphical.frx":7F72
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   49
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picFiveH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5280
      Picture         =   "frmJuliaRenSevensGraphical.frx":81EA
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   50
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picFourH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4800
      Picture         =   "frmJuliaRenSevensGraphical.frx":843F
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   51
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picThreeH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4320
      Picture         =   "frmJuliaRenSevensGraphical.frx":867B
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   52
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picTwoH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3840
      Picture         =   "frmJuliaRenSevensGraphical.frx":8893
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   53
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picAceH 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3360
      Picture         =   "frmJuliaRenSevensGraphical.frx":8A90
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   54
      Top             =   3960
      Width           =   1155
   End
   Begin VB.PictureBox picSixS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5760
      Picture         =   "frmJuliaRenSevensGraphical.frx":8C6A
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   61
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picFiveS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   5280
      Picture         =   "frmJuliaRenSevensGraphical.frx":8EEA
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   62
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picFourS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4800
      Picture         =   "frmJuliaRenSevensGraphical.frx":9147
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   63
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picThreeS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   4320
      Picture         =   "frmJuliaRenSevensGraphical.frx":9380
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   64
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picTwoS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3840
      Picture         =   "frmJuliaRenSevensGraphical.frx":95A7
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   65
      Top             =   5640
      Width           =   1155
   End
   Begin VB.PictureBox picAceS 
      AutoSize        =   -1  'True
      Height          =   1515
      Left            =   3360
      Picture         =   "frmJuliaRenSevensGraphical.frx":97A3
      ScaleHeight     =   1455
      ScaleWidth      =   1095
      TabIndex        =   66
      Top             =   5640
      Width           =   1155
   End
   Begin VB.Label lblLegend 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   75
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblUserCardNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12840
      TabIndex        =   74
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblUserCardRemaining 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your remaining cards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   73
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblOpponentLastCard 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   72
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label lblOpponentCard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Card your opponent just went with:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   71
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label lblSpadesDone 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   70
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label lblHeartsDone 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   69
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblCloversDone 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   68
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblDiamondsDone 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   67
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblOpponentCardNumber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12840
      TabIndex        =   13
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label lblOpponentRemainingCard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your opponent's remaining cards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   12
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label lblSpades 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblHearts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblClovers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblDiamonds 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblSelectCard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the card you want to play:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label lblCards 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   7440
      Width           =   9135
   End
   Begin VB.Label lblYourCards 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Cards:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lblGameRule 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sevens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmCardGameSevens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer Name: Julia Ren
'Date: Dec 16, 2015
'Course: ICS 2O1
'
'User Input: select a sequence of the card already put down or another 7
'Process: input chosen card according to game rule, or skip the turn if no moves are possible
'Output: play the game and diaplay result

Const intStart As Integer = 1 'Take card out from the 1st card on
Const intStep As Integer = 2 'Skip 2 string characters each time
Private strPlayer1Cards As String 'Computer's card
Private strPlayer2Cards As String 'User's card
Private strPlayer1LeftCards As String 'Computer's card on the left
Private strPlayer2LeftCards As String 'User's card on the left
Private strPlayer1RightCards As String 'Computer's card on the right
Private strPlayer2RightCards As String 'User's card on the right
Private intPlayer1Length As Integer 'Number of card the computer has
Private intPlayer2Length As Integer 'Number of cards the user has
Private intIncrement As Integer 'Picture box number


Private Sub cmdEndGame_Click()
    'Exit game
    Unload Me
End Sub

Private Sub cmdHowToPlay_Click()
    'To read the instructions again if user forgets or didn't pay attention
    MsgBox ("The player who has 7 of diamonds starts the game. Each player, if they can, add a sequel card to the existing sequence. This can either go up (8, then 9, then 10 etc) or down (6, then 5, then 4 etc). A player can also start a new sequence in a different suit by placing any of the other 7s. If a player can do neither, they skip a turn. You can skip a turn by just clicking the Play Card button without inputting a card.The winner is the first player to use up all his cards. The winner will be displayed at the game of each game and you can choose to start a new round or exit this game."), , ("Card Game Instructions")
End Sub

Private Sub cmdPlayCard_Click()
    Dim blnSevenOfDiamonds As Boolean 'Search for 7 of diamonds
    Dim blnSevenOfClovers As Boolean 'Search for 7 of clovers
    Dim blnSevenOfHearts As Boolean 'Search for 7 of hearts
    Dim blnSevenOfSpades As Boolean 'Search for 7 of spades
    Dim blnCompareCard As Boolean 'Compare user input to see if card is inside user's card
    Dim strUserInput As String 'User input
    Dim strUserInputNumber As String 'Card number in user input
    Dim strUserInputSuit As String 'Card suit in user input
    Dim strDiamonds As String 'Diamond cards already put down
    Dim strClovers As String 'Clover cards already put down
    Dim strHearts As String 'Heart cards already put down
    Dim strSpades As String 'Spade cards already put down
    Dim strSeperateCard As String 'Each of computer's card
    Dim strSevenOfDiamonds As String '7 of diamonds
    Dim strSevenOfClovers As String '7 of clovers
    Dim strSevenOfHearts As String '7 of hearts
    Dim strSevenOfSpades As String '7 of spades
    Dim strSeperateCardNumber As String 'Card number of the computer card taken out
    Dim strSeperateCardSuit As String 'Card suit of the computer card taken out
    Dim strDiamondsLowest As String 'Lowest card in the diamond cards already put down
    Dim strCloversLowest As String 'Lowest card in the clover cards already put down
    Dim strHeartsLowest As String 'Lowest card in the heart cards already put down
    Dim strSpadesLowest As String 'Lowest card in the spade cards already put down
    Dim strDiamondsHighest As String 'Highest card in the spade cards already put down
    Dim strCloversHighest As String 'Highest card in the spade cards already put down
    Dim strHeartsHighest As String 'Highest card in the spade cards already put down
    Dim strSpadesHighest As String 'Highest card in the spade cards already put down
    Dim strCompareCard As String 'Card taken out to compare if user input is part of user's card
    Dim intUserInputLength As Integer 'Length of user input
    Dim intDiamondsLowest As Integer 'Lowest card number in the diamond cards already put down
    Dim intCloversLowest As Integer 'Lowest card number in the clover cards already put down
    Dim intHeartsLowest As Integer 'Lowest card number in the heart cards already put down
    Dim intSpadesLowest As Integer 'Lowest card number in the spade cards already put down
    Dim intDiamondsHighest As Integer 'Highest card number in the spade cards already put down
    Dim intCloversHighest As Integer 'Highest card number in the spade cards already put down
    Dim intHeartsHighest As Integer 'Highest card number in the spade cards already put down
    Dim intSpadesHighest As Integer 'Highest card number in the spade cards already put down
    Dim intSelectedCardStart As Integer 'Where does user input card start in user's cards
    Dim intSeperateCard As Integer 'Take each computer card out to compare
    Dim intDiamondsLength As Integer 'Length of diamond cards already put down
    Dim intCloversLength As Integer 'Length of clover cards already put down
    Dim intHeartsLength As Integer 'Length of heart cards already put down
    Dim intSpadesLength As Integer 'Length of spade cards already put down
    Dim intSeperateCardNumber As Integer 'Card number of the card selected from computer's cards
    Dim intCompareCard As Integer 'Take user's card to compare with user input
    Dim intUserInputNumber As Integer 'Card number of user input
 
    strDiamonds = lblDiamonds.Caption
    strClovers = lblClovers.Caption
    strHearts = lblHearts.Caption
    strSpades = lblSpades.Caption

    intPlayer1Length = Len(strPlayer1Cards)
    intPlayer2Length = Len(strPlayer2Cards)
        
    strUserInput = txtSelectedCard.Text
    
    intUserInputLength = Len(strUserInput)
    
    strSevenOfDiamonds = "7d"
    strSevenOfClovers = "7c"
    strSevenOfHearts = "7h"
    strSevenOfSpades = "7s"
    
    'Determines if user input is an appropriate card to put down
        If strUserInput <> "" Then
            If strUserInput <> "" Then
                For intCompareCard = intStart To intPlayer2Length Step intStep
                    strCompareCard = Mid(strPlayer2Cards, intCompareCard, 2)
                        
                        If strUserInput = strCompareCard Then
                            blnCompareCard = True
                            Exit For
                        Else
                            blnCompareCard = False
                        End If
                Next intCompareCard
            End If
        
     'User input is not a card or does not exist in user's cards
            If intUserInputLength <> 2 Or blnCompareCard = False Then
                MsgBox "Please enter a valid card."
                txtSelectedCard.Text = ""
                Exit Sub
            End If
                          
            blnSevenOfDiamonds = strUserInput Like strSevenOfDiamonds
            blnSevenOfClovers = strUserInput Like strSevenOfClovers
            blnSevenOfHearts = strUserInput Like strSevenOfHearts
            blnSevenOfSpades = strUserInput Like strSevenOfSpades
                        
            strUserInputNumber = Left(strUserInput, 1)
            strUserInputSuit = Right(strUserInput, 1)
                    
     'If user input contains ace,ten,jack,queen or king
                If strUserInputNumber = "A" Then
                    strUserInputNumber = "1"
                ElseIf strUserInputNumber = "T" Then
                    strUserInputNumber = "10"
                ElseIf strUserInputNumber = "J" Then
                    strUserInputNumber = "11"
                ElseIf strUserInputNumber = "Q" Then
                    strUserInputNumber = "12"
                ElseIf strUserInputNumber = "K" Then
                    strUserInputNumber = "13"
                End If
                
            intUserInputNumber = Val(strUserInputNumber)
          
     'If user inputs a card of diamonds
                If strUserInputSuit = "d" Then
                    If strDiamonds = "" Then
                        If blnSevenOfDiamonds = True Then
                            lblDiamonds.Caption = strUserInput
                        Else
                            MsgBox "This card cannot be put down."
                            Exit Sub
                        End If
                    Else
                        strDiamondsLowest = Left(strDiamonds, 1)
                
                            If strDiamondsLowest = "A" Then
                                strDiamondsLowest = "1"
                            End If
                                
                        intDiamondsLowest = Val(strDiamondsLowest)
                            
                        intDiamondsLength = Len(strDiamonds)
                            
                        strDiamondsHighest = Mid(strDiamonds, intDiamondsLength - 1, 1)
                            
                            If strDiamondsHighest = "T" Then
                                strDiamondsHighest = "10"
                            ElseIf strDiamondsHighest = "J" Then
                                strDiamondsHighest = "11"
                            ElseIf strDiamondsHighest = "Q" Then
                                strDiamondsHighest = "12"
                            ElseIf strDiamondsHighest = "K" Then
                                strDiamondsHighest = "13"
                            End If
                                
                        intDiamondsHighest = Val(strDiamondsHighest)
                        
                            If intDiamondsLowest - intUserInputNumber = 1 Then
                                lblDiamonds.Caption = strUserInput + strDiamonds
                            ElseIf intUserInputNumber - intDiamondsHighest = 1 Then
                                lblDiamonds.Caption = strDiamonds + strUserInput
                            Else
                                MsgBox "This card cannot be put down."
                                Exit Sub
                            End If
                    End If
                End If
    
    'If user inputs a card of clovers
                If strUserInputSuit = "c" Then
                    If strClovers = "" Then
                        If blnSevenOfClovers = True Then
                            lblClovers.Caption = strUserInput
                        Else
                            MsgBox "This card cannot be put down."
                            Exit Sub
                        End If
                    Else
                        strCloversLowest = Left(strClovers, 1)
                        
                            If strCloversLowest = "A" Then
                                strCloversLowest = "1"
                            End If
                            
                        intCloversLowest = Val(strCloversLowest)
                        
                        intCloversLength = Len(strClovers)
                        
                        strCloversHighest = Mid(strClovers, intCloversLength - 1, 1)
                        
                            If strCloversHighest = "T" Then
                                strCloversHighest = "10"
                            ElseIf strCloversHighest = "J" Then
                                strCloversHighest = "11"
                            ElseIf strCloversHighest = "Q" Then
                                strCloversHighest = "12"
                            ElseIf strCloversHighest = "K" Then
                                strCloversHighest = "13"
                            End If
                            
                        intCloversHighest = Val(strCloversHighest)
                    
                            If intCloversLowest - intUserInputNumber = 1 Then
                                lblClovers.Caption = strUserInput + strClovers
                            ElseIf intUserInputNumber - intCloversHighest = 1 Then
                                lblClovers.Caption = strClovers + strUserInput
                            Else
                                MsgBox "This card cannot be put down."
                                Exit Sub
                            End If
                    End If
                End If
                
    'If user inputs a card of hearts
                If strUserInputSuit = "h" Then
                    If strHearts = "" Then
                        If blnSevenOfHearts = True Then
                            lblHearts.Caption = strUserInput
                        Else
                            MsgBox "This card cannot be put down."
                            Exit Sub
                        End If
                    Else
                        strHeartsLowest = Left(strHearts, 1)
                            
                            If strHeartsLowest = "A" Then
                                strHeartsLowest = "1"
                            End If
                                
                        intHeartsLowest = Val(strHeartsLowest)
                                
                        intHeartsLength = Len(strHearts)
                                
                        strHeartsHighest = Mid(strHearts, intHeartsLength - 1, 1)
                                
                            If strHeartsHighest = "T" Then
                                strHeartsHighest = "10"
                            ElseIf strHeartsHighest = "J" Then
                                strHeartsHighest = "11"
                            ElseIf strHeartsHighest = "Q" Then
                                strHeartsHighest = "12"
                            ElseIf strHeartsHighest = "K" Then
                                strHeartsHighest = "13"
                            End If
                                
                        intHeartsHighest = Val(strHeartsHighest)
                            
                            If intHeartsLowest - intUserInputNumber = 1 Then
                                lblHearts.Caption = strUserInput + strHearts
                            ElseIf intUserInputNumber - intHeartsHighest = 1 Then
                                lblHearts.Caption = strHearts + strUserInput
                            Else
                                MsgBox "This card cannot be put down."
                                Exit Sub
                            End If
                    End If
                End If
                
    'If user inputs a card of spades
                If strUserInputSuit = "s" Then
                    If strSpades = "" Then
                        If blnSevenOfSpades = True Then
                            lblSpades.Caption = strUserInput
                        Else
                            MsgBox "This card cannot be put down."
                            Exit Sub
                        End If
                    Else
                        strSpadesLowest = Left(strSpades, 1)
                            
                            If strSpadesLowest = "A" Then
                                strSpadesLowest = "1"
                            End If
                            
                        intSpadesLowest = Val(strSpadesLowest)
                                   
                        intSpadesLength = Len(strSpades)
                                        
                        strSpadesHighest = Mid(strSpades, intSpadesLength - 1, 1)
                            
                            If strSpadesHighest = "T" Then
                                strSpadesHighest = "10"
                            ElseIf strSpadesHighest = "J" Then
                                strSpadesHighest = "11"
                            ElseIf strSpadesHighest = "Q" Then
                                strSpadesHighest = "12"
                            ElseIf strSpadesHighest = "K" Then
                                strSpadesHighest = "13"
                            End If
                            
                        intSpadesHighest = Val(strSpadesHighest)
                                 
                            If intSpadesLowest - intUserInputNumber = 1 Then
                                lblSpades.Caption = strUserInput + strSpades
                            ElseIf intUserInputNumber - intSpadesHighest = 1 Then
                                    lblSpades.Caption = strSpades + strUserInput
                            Else
                                MsgBox "This card cannot be put down."
                                Exit Sub
                            End If
                    End If
                End If
                        
            txtSelectedCard.Text = ""
            
    'Delete the card from user's deck
                If strUserInput <> "" Then
                    intSelectedCardStart = InStr(strPlayer2Cards, strUserInput)
                    Debug.Print "Selected Card start: "; intSelectedCardStart
                        
                    strPlayer2LeftCards = Left(strPlayer2Cards, intSelectedCardStart - 1)
                    strPlayer2RightCards = Right(strPlayer2Cards, intPlayer2Length - intSelectedCardStart - 1)
                    strPlayer2Cards = strPlayer2LeftCards + strPlayer2RightCards
                End If
                       
            lblCards.Caption = strPlayer2Cards
            
                If strUserInput = "Ad" Then
                    picAceD.Visible = True
                ElseIf strUserInput = "Ac" Then
                    picAceC.Visible = True
                ElseIf strUserInput = "Ah" Then
                    picAceH.Visible = True
                ElseIf strUserInput = "As" Then
                    picAceS.Visible = True
                ElseIf strUserInput = "2d" Then
                    picTwoD.Visible = True
                ElseIf strUserInput = "2c" Then
                    picTwoC.Visible = True
                ElseIf strUserInput = "2h" Then
                    picTwoH.Visible = True
                ElseIf strUserInput = "2s" Then
                    picTwoS.Visible = True
                ElseIf strUserInput = "3d" Then
                    picThreeD.Visible = True
                ElseIf strUserInput = "3c" Then
                    picThreeC.Visible = True
                ElseIf strUserInput = "3h" Then
                    picThreeH.Visible = True
                ElseIf strUserInput = "3s" Then
                    picThreeS.Visible = True
                ElseIf strUserInput = "4d" Then
                    picFourD.Visible = True
                ElseIf strUserInput = "4c" Then
                    picFourC.Visible = True
                ElseIf strUserInput = "4h" Then
                    picFourH.Visible = True
                ElseIf strUserInput = "4s" Then
                    picFourS.Visible = True
                ElseIf strUserInput = "5d" Then
                    picFiveD.Visible = True
                ElseIf strUserInput = "5c" Then
                    picFiveC.Visible = True
                ElseIf strUserInput = "5h" Then
                    picFiveH.Visible = True
                ElseIf strUserInput = "5s" Then
                    picFiveS.Visible = True
                ElseIf strUserInput = "6d" Then
                    picSixD.Visible = True
                ElseIf strUserInput = "6c" Then
                    picSixC.Visible = True
                ElseIf strUserInput = "6h" Then
                    picSixH.Visible = True
                ElseIf strUserInput = "6s" Then
                    picSixS.Visible = True
                ElseIf strUserInput = "7d" Then
                    picSevens(0).Visible = True
                ElseIf strUserInput = "7c" Then
                    picSevens(1).Visible = True
                ElseIf strUserInput = "7h" Then
                    picSevens(2).Visible = True
                ElseIf strUserInput = "7s" Then
                    picSevens(3).Visible = True
                ElseIf strUserInput = "8d" Then
                    picEightD.Visible = True
                ElseIf strUserInput = "8c" Then
                    picEightC.Visible = True
                ElseIf strUserInput = "8h" Then
                    picEightH.Visible = True
                ElseIf strUserInput = "8s" Then
                    picEightS.Visible = True
                ElseIf strUserInput = "9d" Then
                    picNineD.Visible = True
                ElseIf strUserInput = "9c" Then
                    picNineC.Visible = True
                ElseIf strUserInput = "9h" Then
                    picNineH.Visible = True
                ElseIf strUserInput = "9s" Then
                    picNineS.Visible = True
                ElseIf strUserInput = "Td" Then
                    picTenD.Visible = True
                ElseIf strUserInput = "Tc" Then
                    picTenC.Visible = True
                ElseIf strUserInput = "Th" Then
                    picTenH.Visible = True
                ElseIf strUserInput = "Ts" Then
                    picTenS.Visible = True
                ElseIf strUserInput = "Jd" Then
                    picJackD.Visible = True
                ElseIf strUserInput = "Jc" Then
                    picJackC.Visible = True
                ElseIf strUserInput = "Jh" Then
                    picJackH.Visible = True
                ElseIf strUserInput = "Js" Then
                    picJackS.Visible = True
                ElseIf strUserInput = "Qd" Then
                    picQueenD.Visible = True
                ElseIf strUserInput = "Qc" Then
                    picQueenC.Visible = True
                ElseIf strUserInput = "Qh" Then
                    picQueenH.Visible = True
                ElseIf strUserInput = "Qs" Then
                    picQueenS.Visible = True
                ElseIf strUserInput = "Kd" Then
                    picKingD.Visible = True
                ElseIf strUserInput = "Kc" Then
                    picKingC.Visible = True
                ElseIf strUserInput = "Kh" Then
                    picKingH.Visible = True
                ElseIf strUserInput = "Ks" Then
                    picKingS.Visible = True
                End If
        End If
    
    'Check computer's and user's number of remaining cards
    intPlayer1Length = Len(strPlayer1Cards)
    intPlayer2Length = Len(strPlayer2Cards)
    
    'Display user's number of remaining cards
    lblUserCardNumber.Caption = (intPlayer2Length - 12) / 2
    
    'If either player has no cards left, display result
        If intPlayer1Length = 0 Then
            MsgBox "You Lost!"
            Exit Sub
        ElseIf intPlayer2Length = 0 Then
            MsgBox "You Won!"
            Exit Sub
        End If
        
    strDiamonds = lblDiamonds.Caption
    strClovers = lblClovers.Caption
    strHearts = lblHearts.Caption
    strSpades = lblSpades.Caption
    
    'Compare each of computer's card to see if it can be put down
    For intSeperateCard = intStart To intPlayer1Length Step intStep
        strSeperateCard = Mid(strPlayer1Cards, intSeperateCard, 2)
        Debug.Print "Seperate Card:"; strSeperateCard
       
        blnSevenOfDiamonds = strSeperateCard Like strSevenOfDiamonds
        blnSevenOfClovers = strSeperateCard Like strSevenOfClovers
        blnSevenOfHearts = strSeperateCard Like strSevenOfHearts
        blnSevenOfSpades = strSeperateCard Like strSevenOfSpades
                    
        strSeperateCardNumber = Left(strSeperateCard, 1)
        strSeperateCardSuit = Right(strSeperateCard, 1)
            
    'If the card contains a ace,ten,jack,queen or king
            If strSeperateCardNumber = "A" Then
                strSeperateCardNumber = "1"
            ElseIf strSeperateCardNumber = "T" Then
                strSeperateCardNumber = "10"
            ElseIf strSeperateCardNumber = "J" Then
                strSeperateCardNumber = "11"
            ElseIf strSeperateCardNumber = "Q" Then
                strSeperateCardNumber = "12"
            ElseIf strSeperateCardNumber = "K" Then
                strSeperateCardNumber = "13"
            End If
        
        intSeperateCardNumber = Val(strSeperateCardNumber)
        
    'If the card is a card of diamonds
            If strSeperateCardSuit = "d" Then
                strDiamondsLowest = Left(strDiamonds, 1)
                
                    If strDiamondsLowest = "A" Then
                        strDiamondsLowest = "1"
                    End If
                    
                intDiamondsLowest = Val(strDiamondsLowest)
                
                intDiamondsLength = Len(strDiamonds)
                
                strDiamondsHighest = Mid(strDiamonds, intDiamondsLength - 1, 1)
                
                    If strDiamondsHighest = "T" Then
                        strDiamondsHighest = "10"
                    ElseIf strDiamondsHighest = "J" Then
                        strDiamondsHighest = "11"
                    ElseIf strDiamondsHighest = "Q" Then
                        strDiamondsHighest = "12"
                    ElseIf strDiamondsHighest = "K" Then
                        strDiamondsHighest = "13"
                    End If
                    
                intDiamondsHighest = Val(strDiamondsHighest)
                
                    If intDiamondsLowest - intSeperateCardNumber = 1 Then
                        lblDiamonds.Caption = strSeperateCard + strDiamonds
                        Exit For
                    ElseIf intSeperateCardNumber - intDiamondsHighest = 1 Then
                        lblDiamonds.Caption = strDiamonds + strSeperateCard
                        Exit For
                    End If
            End If
        
    'If the card is a card of clovers
            If strSeperateCardSuit = "c" Then
                If strClovers = "" Then
                    If blnSevenOfClovers = True Then
                        lblClovers.Caption = strSeperateCard
                        Exit For
                    End If
                Else
                    strCloversLowest = Left(strClovers, 1)
                
                        If strCloversLowest = "A" Then
                            strCloversLowest = "1"
                        End If
                    
                    intCloversLowest = Val(strCloversLowest)
                
                    intCloversLength = Len(strClovers)
                
                    strCloversHighest = Mid(strClovers, intCloversLength - 1, 1)
                
                        If strCloversHighest = "T" Then
                            strCloversHighest = "10"
                        ElseIf strCloversHighest = "J" Then
                            strCloversHighest = "11"
                        ElseIf strCloversHighest = "Q" Then
                            strCloversHighest = "12"
                        ElseIf strCloversHighest = "K" Then
                            strCloversHighest = "13"
                        End If
                    
                    intCloversHighest = Val(strCloversHighest)
                    
                        If intCloversLowest - intSeperateCardNumber = 1 Then
                            lblClovers.Caption = strSeperateCard + strClovers
                            Exit For
                        ElseIf intSeperateCardNumber - intCloversHighest = 1 Then
                            lblClovers.Caption = strClovers + strSeperateCard
                            Exit For
                        End If
                End If
            End If
            
    'If the card is a card of hearts
            If strSeperateCardSuit = "h" Then
                If strHearts = "" Then
                    If blnSevenOfHearts = True Then
                        lblHearts.Caption = strSeperateCard
                        Exit For
                    End If
                Else
                    strHeartsLowest = Left(strHearts, 1)
                    
                        If strHeartsLowest = "A" Then
                            strHeartsLowest = "1"
                        End If
                        
                    intHeartsLowest = Val(strHeartsLowest)
                        
                    intHeartsLength = Len(strHearts)
                        
                    strHeartsHighest = Mid(strHearts, intHeartsLength - 1, 1)
                        
                        If strHeartsHighest = "T" Then
                            strHeartsHighest = "10"
                        ElseIf strHeartsHighest = "J" Then
                            strHeartsHighest = "11"
                        ElseIf strHeartsHighest = "Q" Then
                            strHeartsHighest = "12"
                        ElseIf strHeartsHighest = "K" Then
                            strHeartsHighest = "13"
                        End If
                        
                    intHeartsHighest = Val(strHeartsHighest)
                    
                        If intHeartsLowest - intSeperateCardNumber = 1 Then
                            lblHearts.Caption = strSeperateCard + strHearts
                            Exit For
                        ElseIf intSeperateCardNumber - intHeartsHighest = 1 Then
                            lblHearts.Caption = strHearts + strSeperateCard
                            Exit For
                        End If
                End If
            End If
                
    'If the card is a card of spades
            If strSeperateCardSuit = "s" Then
                If strSpades = "" Then
                    If blnSevenOfSpades = True Then
                        lblSpades.Caption = strSeperateCard
                        Exit For
                    End If
                Else
                    strSpadesLowest = Left(strSpades, 1)
                    
                        If strSpadesLowest = "A" Then
                            strSpadesLowest = "1"
                        End If
                    
                    intSpadesLowest = Val(strSpadesLowest)
                           
                    intSpadesLength = Len(strSpades)
                                
                    strSpadesHighest = Mid(strSpades, intSpadesLength - 1, 1)
                    
                        If strSpadesHighest = "T" Then
                            strSpadesHighest = "10"
                        ElseIf strSpadesHighest = "J" Then
                            strSpadesHighest = "11"
                        ElseIf strSpadesHighest = "Q" Then
                            strSpadesHighest = "12"
                        ElseIf strSpadesHighest = "K" Then
                            strSpadesHighest = "13"
                        End If
                    
                    intSpadesHighest = Val(strSpadesHighest)
                         
                        If intSpadesLowest - intSeperateCardNumber = 1 Then
                            lblSpades.Caption = strSeperateCard + strSpades
                            Exit For
                        ElseIf intSeperateCardNumber - intSpadesHighest = 1 Then
                            lblSpades.Caption = strSpades + strSeperateCard
                            Exit For
                        End If
                End If
            End If
        
        strSeperateCard = ""
    Next intSeperateCard
    
    'If computer did not put down a card,user continues
        If strSeperateCard = "" Then
            MsgBox "Your opponent skipped a turn."
        Else
        'If computer put down a card, delete the card from computer's deck
            intSelectedCardStart = InStr(strPlayer1Cards, strSeperateCard)
            Debug.Print "Selected Card start: "; intSelectedCardStart
                
            strPlayer1LeftCards = Left(strPlayer1Cards, intSelectedCardStart - 1)
            strPlayer1RightCards = Right(strPlayer1Cards, intPlayer1Length - intSelectedCardStart - 1)
            strPlayer1Cards = strPlayer1LeftCards + strPlayer1RightCards

                If strSeperateCard = "Ad" Then
                    picAceD.Visible = True
                ElseIf strSeperateCard = "Ac" Then
                    picAceC.Visible = True
                ElseIf strSeperateCard = "Ah" Then
                    picAceH.Visible = True
                ElseIf strSeperateCard = "As" Then
                    picAceS.Visible = True
                ElseIf strSeperateCard = "2d" Then
                    picTwoD.Visible = True
                ElseIf strSeperateCard = "2c" Then
                    picTwoC.Visible = True
                ElseIf strSeperateCard = "2h" Then
                    picTwoH.Visible = True
                ElseIf strSeperateCard = "2s" Then
                    picTwoS.Visible = True
                ElseIf strSeperateCard = "3d" Then
                    picThreeD.Visible = True
                ElseIf strSeperateCard = "3c" Then
                    picThreeC.Visible = True
                ElseIf strSeperateCard = "3h" Then
                    picThreeH.Visible = True
                ElseIf strSeperateCard = "3s" Then
                    picThreeS.Visible = True
                ElseIf strSeperateCard = "4d" Then
                    picFourD.Visible = True
                ElseIf strSeperateCard = "4c" Then
                    picFourC.Visible = True
                ElseIf strSeperateCard = "4h" Then
                    picFourH.Visible = True
                ElseIf strSeperateCard = "4s" Then
                    picFourS.Visible = True
                ElseIf strSeperateCard = "5d" Then
                    picFiveD.Visible = True
                ElseIf strSeperateCard = "5c" Then
                    picFiveC.Visible = True
                ElseIf strSeperateCard = "5h" Then
                    picFiveH.Visible = True
                ElseIf strSeperateCard = "5s" Then
                    picFiveS.Visible = True
                ElseIf strSeperateCard = "6d" Then
                    picSixD.Visible = True
                ElseIf strSeperateCard = "6c" Then
                    picSixC.Visible = True
                ElseIf strSeperateCard = "6h" Then
                    picSixH.Visible = True
                ElseIf strSeperateCard = "6s" Then
                    picSixS.Visible = True
                ElseIf strSeperateCard = "7c" Then
                    picSevens(1).Visible = True
                ElseIf strSeperateCard = "7h" Then
                    picSevens(2).Visible = True
                ElseIf strSeperateCard = "7s" Then
                    picSevens(3).Visible = True
                ElseIf strSeperateCard = "8d" Then
                    picEightD.Visible = True
                ElseIf strSeperateCard = "8c" Then
                    picEightC.Visible = True
                ElseIf strSeperateCard = "8h" Then
                    picEightH.Visible = True
                ElseIf strSeperateCard = "8s" Then
                    picEightS.Visible = True
                ElseIf strSeperateCard = "9d" Then
                    picNineD.Visible = True
                ElseIf strSeperateCard = "9c" Then
                    picNineC.Visible = True
                ElseIf strSeperateCard = "9h" Then
                    picNineH.Visible = True
                ElseIf strSeperateCard = "9s" Then
                    picNineS.Visible = True
                ElseIf strSeperateCard = "Td" Then
                    picTenD.Visible = True
                ElseIf strSeperateCard = "Tc" Then
                    picTenC.Visible = True
                ElseIf strSeperateCard = "Th" Then
                    picTenH.Visible = True
                ElseIf strSeperateCard = "Ts" Then
                    picTenS.Visible = True
                ElseIf strSeperateCard = "Jd" Then
                    picJackD.Visible = True
                ElseIf strSeperateCard = "Jc" Then
                    picJackC.Visible = True
                ElseIf strSeperateCard = "Jh" Then
                    picJackH.Visible = True
                ElseIf strSeperateCard = "Js" Then
                    picJackS.Visible = True
                ElseIf strSeperateCard = "Qd" Then
                    picQueenD.Visible = True
                ElseIf strSeperateCard = "Qc" Then
                    picQueenC.Visible = True
                ElseIf strSeperateCard = "Qh" Then
                    picQueenH.Visible = True
                ElseIf strSeperateCard = "Qs" Then
                    picQueenS.Visible = True
                ElseIf strSeperateCard = "Kd" Then
                    picKingD.Visible = True
                ElseIf strSeperateCard = "Kc" Then
                    picKingC.Visible = True
                ElseIf strSeperateCard = "Kh" Then
                    picKingH.Visible = True
                ElseIf strSeperateCard = "Ks" Then
                    picKingS.Visible = True
                End If
        End If
       
    lblOpponentLastCard.Caption = strSeperateCard

    Debug.Print "Player 1 Cards: "; strPlayer1Cards
    Debug.Print "Player 1 Played Card: "; strSeperateCard
              
    'Check computer's and user's number of remaining cards
    intPlayer1Length = Len(strPlayer1Cards)
    intPlayer2Length = Len(strPlayer2Cards) - 12
    
    'Display computer's number of remaining card
    lblOpponentCardNumber.Caption = intPlayer1Length / 2
    
    'Display if a suit is finished
        If lblDiamonds.Caption = "Ad2d3d4d5d6d7d8d9dTdJdQdKd" Then
            lblDiamondsDone.Caption = "Diamonds are finished!"
        End If
        
        If lblClovers.Caption = "Ac2c3c4c5c6c7c8c9cTcJcQcKc" Then
            lblCloversDone.Caption = "Clovers are finished!"
        End If
        
        If lblHearts.Caption = "Ah2h3h4h5h6h7h8h9hThJhQhKh" Then
            lblHeartsDone.Caption = "Hearts are finished!"
        End If
        
        If lblSpades.Caption = "As2s3s4s5s6s7s8s9sTsJsQsKs" Then
            lblSpadesDone.Caption = "Spades are finished!"
        End If
        
    'If either player has no cards left, display result
        If intPlayer1Length = 0 Then
            MsgBox "You Lost!"
        ElseIf intPlayer2Length = 0 Then
            MsgBox "You Won!"
        End If
End Sub
        

Private Sub cmdStartNewGame_Click()
    Randomize                                  'Generates a different random number every time
    
    Const intHighNumShuffle As Integer = 52 'Highest number the random number can go to
    Const intLowNumShuffle As Integer = 1 'Lowest number the random number can go to
    Dim blnSearchPlayer1 As Boolean 'Search computer for 7 of diamonds to start the game
    Dim blnSearchPlayer2 As Boolean 'Search user for 7 of diamonds to start the game
    Dim blnDiamondsCard As Boolean 'Search if a card is diamond
    Dim blnCloversCard As Boolean 'Search if a card is clovers
    Dim blnHeartsCard As Boolean 'Search if a card is hearts
    Dim blnSpadesCard As Boolean 'Search if a card is spades
    Dim strDeck As String 'The full deck
    Dim strCard As String 'Card taken out to shuffle
    Dim strRightCards As String 'Cards on the right of the card taken out to shuffle
    Dim strLeftCards As String 'Cards on the left of the card taken out to shuffle
    Dim strSearchFirstCard As String 'Search for 7 of diamonds
    Dim strUserDiamonds As String 'Diamond cards in user's cards
    Dim strUserClovers As String 'Clover cards in user's cards
    Dim strUserHearts As String 'Heart cards in user's cards
    Dim strUserSpades As String 'Spade cards in user's cards
    Dim strDiamondsCard As String 'A diamond card
    Dim strCloversCard As String 'A clovers card
    Dim strHeartsCard As String 'A hearts card
    Dim strSpadesCard As String 'A spades card
    Dim strEachUserCard As String 'Each of user's card
    Dim intLengthDeck As Integer 'Length of the full deck
    Dim intRandomNumberShuffle As Integer 'Take a random card out
    Dim intCounter As Integer 'Number of times to shuffle
    Dim intPlayerStart As Integer 'Where is 7 of diamonds
    Dim intEachUserCard As Integer 'Where each of user's card start

    Debug.Print " ---- NEW PROGRAM ----"

    lblDiamonds.Caption = ""
    lblClovers.Caption = ""
    lblHearts.Caption = ""
    lblSpades.Caption = ""
    
    lblDiamondsDone.Caption = ""
    lblCloversDone.Caption = ""
    lblHeartsDone.Caption = ""
    lblSpadesDone.Caption = ""
    
    lblOpponentCardNumber.Caption = ""
    lblOpponentLastCard.Caption = ""
    
    
    tmrBlinkSevens.Enabled = False
    picSevens(0).Visible = False
    picSevens(1).Visible = False
    picSevens(2).Visible = False
    picSevens(3).Visible = False
    
    picAceD.Visible = False
    picAceC.Visible = False
    picAceH.Visible = False
    picAceS.Visible = False
    picTwoD.Visible = False
    picTwoC.Visible = False
    picTwoH.Visible = False
    picTwoS.Visible = False
    picThreeD.Visible = False
    picThreeC.Visible = False
    picThreeH.Visible = False
    picThreeS.Visible = False
    picFourD.Visible = False
    picFourC.Visible = False
    picFourH.Visible = False
    picFourS.Visible = False
    picFiveD.Visible = False
    picFiveC.Visible = False
    picFiveH.Visible = False
    picFiveS.Visible = False
    picSixD.Visible = False
    picSixC.Visible = False
    picSixH.Visible = False
    picSixS.Visible = False
    picEightD.Visible = False
    picEightC.Visible = False
    picEightH.Visible = False
    picEightS.Visible = False
    picNineD.Visible = False
    picNineC.Visible = False
    picNineH.Visible = False
    picNineS.Visible = False
    picTenD.Visible = False
    picTenC.Visible = False
    picTenH.Visible = False
    picTenS.Visible = False
    picJackD.Visible = False
    picJackC.Visible = False
    picJackH.Visible = False
    picJackS.Visible = False
    picQueenD.Visible = False
    picQueenC.Visible = False
    picQueenH.Visible = False
    picQueenS.Visible = False
    picKingD.Visible = False
    picKingC.Visible = False
    picKingH.Visible = False
    picKingS.Visible = False
    
    strPlayer1Cards = ""
    strPlayer2Cards = ""
    
    strUserDiamonds = ""
    strUserClovers = ""
    strUserHearts = ""
    strUserSpades = ""
    
    strDiamondsCard = "?d"
    strCloversCard = "?c"
    strHeartsCard = "?h"
    strSpadesCard = "?s"
    
    strDeck = "Ad2d3d4d5d6d7d8d9dTdJdQdKdAc2c3c4c5c6c7c8c9cTcJcQcKcAh2h3h4h5h6h7h8h9hThJhQhKhAs2s3s4s5s6s7s8s9sTsJsQsKs"
    Debug.Print "Deck: "; strDeck
    
    intLengthDeck = Len(strDeck)
    Debug.Print "Length of Deck: "; intLengthDeck
    
    'Shuffle the full deck
    For intCounter = 1 To 200
        intRandomNumberShuffle = Int((intHighNumShuffle - intLowNumShuffle + 1) * Rnd + intLowNumShuffle)
        
        If intRandomNumberShuffle <> intLengthDeck And intRandomNumberShuffle Mod 2 <> 0 Then
            'Randomly take a card out of the deck and put it as the last card of the deck

            strCard = Mid(strDeck, intRandomNumberShuffle, 2)
    
            strLeftCards = Left(strDeck, intRandomNumberShuffle - 1)
    
            strRightCards = Right(strDeck, intLengthDeck - intRandomNumberShuffle - 1)
    
            strDeck = strLeftCards + strRightCards + strCard
        End If
    Next intCounter
    
    Debug.Print "Random Number: "; intRandomNumber
    Debug.Print "Card: "; strCard
    Debug.Print "Left cards: "; strLeftCards
    Debug.Print "Right cards: "; strRightCards
    Debug.Print "Shuffled Deck: "; strDeck
    
    'Deal cards
    Do While Len(strDeck) > 0 ' Repeat loop if the deck still has card
        strPlayer1Cards = strPlayer1Cards + Left(strDeck, 2)
        
        strDeck = Right(strDeck, Len(strDeck) - 2)
        
        'If the last card was dealt to Player 1, exit the Do loop
        If Len(strDeck) = 0 Then
            Exit Do
        End If
                    
        strPlayer2Cards = strPlayer2Cards + Left(strDeck, 2)
        
        strDeck = Right(strDeck, Len(strDeck) - 2)
    Loop
    
    Debug.Print "Player 1 cards: "; strPlayer1Cards
    Debug.Print "Player 2 cards: "; strPlayer2Cards
    Debug.Print "Deck: "; strDeck
    
    'Search which player has 7 of diamonds
    strSearchFirstCard = "*7d*"
    blnSearchPlayer1 = strPlayer1Cards Like strSearchFirstCard
    blnSearchPlayer2 = strPlayer2Cards Like strSearchFirstCard
    
    Debug.Print "Search player 1: "; blnSearchPlayer1
    Debug.Print "Search player 2: "; blnSearchPlayer2
    
    'If computer has 7 of diamonds
    If blnSearchPlayer1 = True Then
        intPlayer1Length = Len(strPlayer1Cards)
        
        intPlayerStart = InStr(strPlayer1Cards, "7d")
        Debug.Print "Player start: "; intPlayerStart
        
        strPlayer1LeftCards = Left(strPlayer1Cards, intPlayerStart - 1)
        
        strPlayer1RightCards = Right(strPlayer1Cards, intPlayer1Length - intPlayerStart - 1)
        
        strPlayer1Cards = strPlayer1LeftCards + strPlayer1RightCards
        
        lblDiamonds.Caption = "7d"
        
        lblOpponentLastCard.Caption = "7d"

        picSevens(0).Visible = True
    End If
    
    'If user has 7 of diamonds
    If blnSearchPlayer2 = True Then
        MsgBox "You have 7 of Diamonds, please take your first turn with this card."
    End If
        
    intPlayer1Length = Len(strPlayer1Cards)
    intPlayer2Length = Len(strPlayer2Cards)
    
    'Organize each of user's cards into suits
    For intEachUserCard = intStart To intPlayer2Length Step intStep
        strEachUserCard = Mid(strPlayer2Cards, intEachUserCard, 2)
        Debug.Print "Each User Card:"; strEachUserCard
        
        blnDiamondsCard = strEachUserCard Like strDiamondsCard
        blnCloversCard = strEachUserCard Like strCloversCard
        blnHeartsCard = strEachUserCard Like strHeartsCard
        blnSpadesCard = strEachUserCard Like strSpadesCard
        
            If blnDiamondsCard = True Then
                strUserDiamonds = strUserDiamonds + strEachUserCard
            End If
            
            If blnCloversCard = True Then
                strUserClovers = strUserClovers + strEachUserCard
            End If
            
            If blnHeartsCard = True Then
                strUserHearts = strUserHearts + strEachUserCard
            End If
            
            If blnSpadesCard = True Then
                strUserSpades = strUserSpades + strEachUserCard
            End If
    Next intEachUserCard
    
    'Display user's cards
    strPlayer2Cards = strUserDiamonds & Space(4) & strUserClovers & Space(4) & strUserHearts & Space(4) & strUserSpades
    lblCards.Caption = strPlayer2Cards
 
    intPlayer1Length = Len(strPlayer1Cards)
    intPlayer2Length = Len(strPlayer2Cards)
    
    'Display computer's number of remaining card
    lblOpponentCardNumber.Caption = intPlayer1Length / 2
    
    'Display user's number of remaining card
    lblUserCardNumber.Caption = (intPlayer2Length - 12) / 2
    
    Debug.Print "Player 1 cards: "; strPlayer1Cards
    Debug.Print "Player 2 cards: "; strPlayer2Cards
End Sub
Private Sub Form_Load()
    Me.BackColor = vbWhite
    
    intIncrement = 0
    
    picSevens(0).Visible = True
    picSevens(1).Visible = False
    picSevens(2).Visible = False
    picSevens(3).Visible = False
    
    picAceD.Visible = False
    picAceC.Visible = False
    picAceH.Visible = False
    picAceS.Visible = False
    picTwoD.Visible = False
    picTwoC.Visible = False
    picTwoH.Visible = False
    picTwoS.Visible = False
    picThreeD.Visible = False
    picThreeC.Visible = False
    picThreeH.Visible = False
    picThreeS.Visible = False
    picFourD.Visible = False
    picFourC.Visible = False
    picFourH.Visible = False
    picFourS.Visible = False
    picFiveD.Visible = False
    picFiveC.Visible = False
    picFiveH.Visible = False
    picFiveS.Visible = False
    picSixD.Visible = False
    picSixC.Visible = False
    picSixH.Visible = False
    picSixS.Visible = False
    picEightD.Visible = False
    picEightC.Visible = False
    picEightH.Visible = False
    picEightS.Visible = False
    picNineD.Visible = False
    picNineC.Visible = False
    picNineH.Visible = False
    picNineS.Visible = False
    picTenD.Visible = False
    picTenC.Visible = False
    picTenH.Visible = False
    picTenS.Visible = False
    picJackD.Visible = False
    picJackC.Visible = False
    picJackH.Visible = False
    picJackS.Visible = False
    picQueenD.Visible = False
    picQueenC.Visible = False
    picQueenH.Visible = False
    picQueenS.Visible = False
    picKingD.Visible = False
    picKingC.Visible = False
    picKingH.Visible = False
    picKingS.Visible = False
    
    'To read the instructions again if user forgets or didn't pay attention
    MsgBox ("The player who has 7 of diamonds starts the game. Each player, if they can, add a sequel card to the existing sequence. This can either go up (8, then 9, then 10 etc) or down (6, then 5, then 4 etc). A player can also start a new sequence in a different suit by placing any of the other 7s. If a player can do neither, they skip a turn. You can skip a turn by just clicking the Play Card button without inputting a card.The winner is the first player to use up all his cards. The winner will be displayed at the game of each game and you can choose to start a new round or exit this game."), , ("Card Game Instructions")
    
    lblLegend.Caption = "d - Diamonds" & vbCrLf & "c - Clovers" & vbCrLf & "h - Hearts" & vbCrLf & "s - Spades"
End Sub

Private Sub tmrBlinkSevens_Timer()
    picSevens(intIncrement).Visible = False
    
        If intIncrement = 3 Then
            intIncrement = 0
        Else
            intIncrement = intIncrement + 1
        End If
    
    picSevens(intIncrement).Visible = True
End Sub
