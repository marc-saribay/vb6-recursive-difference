VERSION 5.00
Begin VB.Form frmGraph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graph"
   ClientHeight    =   7365
   ClientLeft      =   285
   ClientTop       =   585
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGraph.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12.991
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   20.294
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGraph 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   720
      ScaleHeight     =   11.245
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   18.018
      TabIndex        =   0
      Top             =   600
      Width           =   10215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " oo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.Line Line6 
      X1              =   8.89
      X2              =   9.102
      Y1              =   0.635
      Y2              =   0.423
   End
   Begin VB.Line Line5 
      X1              =   9.102
      X2              =   8.89
      Y1              =   0.423
      Y2              =   0.212
   End
   Begin VB.Line Line4 
      X1              =   9.102
      X2              =   8.467
      Y1              =   0.423
      Y2              =   0.423
   End
   Begin VB.Line Line3 
      X1              =   0.635
      X2              =   0.847
      Y1              =   6.35
      Y2              =   6.138
   End
   Begin VB.Line Line2 
      X1              =   0.635
      X2              =   0.423
      Y1              =   6.35
      Y2              =   6.138
   End
   Begin VB.Line Line1 
      X1              =   0.635
      X2              =   0.635
      Y1              =   6.35
      Y2              =   5.715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub picGraph_Paint()
  frmGraph.Caption = "Generating Graph..."
  Dim X(100) As Double
  Dim K As Double, I As Integer, KX As Integer, KY As Double, XX As Double, XY As Integer
  For KX = 0 To 20
    For KY = 0 To 12 Step 0.05
      frmGraph.picGraph.PSet (KX, KY)    'Generate grid lines from 0 to 20 on K axis
    Next
  Next
  For XY = 0 To 12
    For XX = 0 To 20 Step 0.05
      frmGraph.picGraph.PSet (XX, XY)    'Generate grid lines from 0 to 12 on X axis
    Next
  Next
'--------------------------------------------
'  Plotting the graph
'--------------------------------------------
  For K = 0.001 To 20 Step 0.01
    I = 1
    X(I - 1) = frmMain.txtX0.Text    'Getting the value of X0
    X(I) = frmMain.txtX1.Text        'Getting the value of X1
    For I = 1 To 99                  'Generate the sequence from X sub 2 to X sub 100
      X(I + 1) = Max(X(I), X(I - 1)) / (K * X(I))            'The given sequence
      frmGraph.picGraph.PSet (K, X(I + 1)), RGB(0, 0, 255)   'Plotting the points
    Next
  Next
'--------------------------------------------
  frmGraph.Caption = "Graph [Done]"
End Sub
