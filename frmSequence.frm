VERSION 5.00
Begin VB.Form frmSequence 
   Caption         =   "Sequence Generated"
   ClientHeight    =   6315
   ClientLeft      =   5265
   ClientTop       =   1470
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSequence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   3525
   Begin VB.ListBox lstXns 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Details here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3180
   End
End
Attribute VB_Name = "frmSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(100) As Double, N As Double, K As Double

Private Sub Form_Load()
  N = 1
  X(N - 1) = frmMain.txtX0.Text
  X(N) = frmMain.txtX1.Text
  K = frmMain.txtK.Text
  lblCaption.Caption = "K   = " & Str(K) & Chr(10) & "X0 = " & Str(X(N - 1)) & Chr(10) & "X1 = " & Str(X(N))    'Displays the values of X0, X1, and K
  '----------------------------------------------
  '  Generates the list of X sub 2 to X sub 100
  '----------------------------------------------
  For N = 1 To 99
    X(N + 1) = Max(X(N), X(N - 1)) / (K * X(N))
    lstXns.AddItem ("X" & Trim(Str((N + 1))) & " = " & Str(X(N + 1)))
  Next
  '----------------------------------------------
End Sub
