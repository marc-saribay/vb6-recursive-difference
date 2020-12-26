VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recursive Difference Analysis"
   ClientHeight    =   4215
   ClientLeft      =   720
   ClientTop       =   1005
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Generate &Graph"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame fraConds 
      Caption         =   "Required Conditions"
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   3855
      Begin VB.Label lblConds 
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSequence 
      Caption         =   "Generate &Sequence"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   2400
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame fraEquation 
      Caption         =   "Recursive Difference Equation"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " n"
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   270
      End
      Begin VB.Line linEquation 
         BorderWidth     =   2
         X1              =   1200
         X2              =   3360
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " n         n-1"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "max { x  , x    }"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " n+1"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblEquation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x    ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   705
      End
   End
   Begin VB.TextBox txtK 
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtX1 
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtX0 
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter X"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter  K"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   5
      Left            =   720
      TabIndex        =   11
      Top             =   2280
      Width           =   120
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   120
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter X"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DisplayConditions()
  Dim strConds As String
  strConds = "Each box must contain a valid numeric value where K > 0"
  lblConds.Caption = strConds    'Display required conditions
End Sub

Private Sub InputChecker()
  'Text boxes must contain a value
  If txtX0.Text <> "" And txtX1.Text <> "" And txtK.Text <> "" Then
    'Text boxes must contain a numeric value
    If IsNumeric(txtX0.Text) = True And IsNumeric(txtX1.Text) = True And IsNumeric(txtK.Text) = True Then
      If txtX1.Text * txtK.Text <> 0 Then    'Division by zero is not allowed
        If txtK.Text > 0 Then                'Required conditions of K
          cmdSequence.Enabled = True         'Enables the Commands if the boxes (x0, x1, k, n) contains valid numeric values
          cmdPlot.Enabled = True
        Else
          cmdSequence.Enabled = False        'Disables the Commands if the boxes (x0, x1, k, n) contains invalid values
          cmdPlot.Enabled = False
        End If
      Else
        cmdSequence.Enabled = False          'Disables the Commands if the boxes (x0, x1, k, n) contains invalid values
        cmdPlot.Enabled = False
      End If
    Else
      cmdSequence.Enabled = False            'Disables the Commands if the boxes (x0, x1, k, n) contains invalid values
      cmdPlot.Enabled = False
    End If
  Else
    cmdSequence.Enabled = False              'Disables the Commands if the boxes (x0, x1, k, n) contains invalid values
    cmdPlot.Enabled = False
  End If
End Sub

Private Sub cmdAbout_Click()
  frmAbout.Show    'Calls the About Dialog Box
End Sub


Private Sub cmdPlot_Click()
  frmGraph.Show    'Calls the Graph Form
End Sub

Private Sub cmdSequence_Click()
 Unload frmSequence
 frmSequence.Show    'Calls the Sequence Form
End Sub

Private Sub Form_Load()
  DisplayConditions
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End    'Ends the program
End Sub

Private Sub txtK_GotFocus()
  txtK.SelStart = 0
  txtK.SelLength = Len(txtK.Text)    'Selects the whole text in the box for quick value replacement whenever you click on it
End Sub

Private Sub txtX0_Change()
  InputChecker    'Calls Input Checker
End Sub

Private Sub txtX0_GotFocus()
  txtX0.SelStart = 0
  txtX0.SelLength = Len(txtX0.Text)    'Selects the whole text in the box for quick value replacement whenever you click on it
End Sub

Private Sub txtX1_Change()
  InputChecker    'Calls Input Checker
End Sub

Private Sub txtK_Change()
  InputChecker    'Calls Input Checker
End Sub

Private Sub txtX1_GotFocus()
  txtX1.SelStart = 0
  txtX1.SelLength = Len(txtX1.Text)    'Selects the whole text in the box for quick value replacement whenever you click on it
End Sub
