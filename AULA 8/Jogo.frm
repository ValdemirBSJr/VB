VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FormJogo"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   600
      TabIndex        =   9
      Top             =   2880
      Width           =   4215
      Begin VB.CommandButton Mesa8 
         Height          =   615
         Index           =   8
         Left            =   2520
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Mesa7 
         Height          =   615
         Index           =   7
         Left            =   1680
         TabIndex        =   17
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Mesa6 
         Height          =   615
         Index           =   6
         Left            =   840
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Mesa5 
         Height          =   615
         Index           =   5
         Left            =   2520
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Mesa4 
         Height          =   615
         Index           =   4
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Mesa3 
         Height          =   615
         Index           =   3
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Mesa2 
         Height          =   615
         Index           =   2
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Mesa1 
         Height          =   615
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Mesa0 
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton BtnIniciar 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Placar"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5295
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Começa jogando o nome que tem o X"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu Sair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

End Sub

Private Sub Sair_Click()

End Sub
