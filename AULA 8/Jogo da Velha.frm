VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Velha"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnEntrar 
      Caption         =   "Entrar"
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox TxtSegundo 
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox TxtPrimeiro 
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Text            =   "Segundo Jogador"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Text            =   "Primeiro Jogador"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Text            =   "Entre  com os nomes dos jogadores"
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEntrar_Click()
    Primeiro = TxtPrimeiro.Text
    Segundo = TxtSegundo.Text
    Unload Me
    FrmJogo.Show
End Sub

Private Sub BtnSair_Click()
    End
End Sub

Private Sub TxtPrimeiro_Change()
    If TxtPrimeiro.Text <> "" And TxtSegundo.Text <> "" Then
       BtnEntrar.Enabled = True
       
    Else
        BtnEntrar.Enabled = False
    End If
    
End Sub

Private Sub TxtSegundo_Change()
    If TxtSegundo.Text <> "" And TxtPrimeiro.Text <> "" Then
        BtnEntrar.Enabled = True
    Else
        BtnEntrar.Enabled = False
    End If
    
End Sub
