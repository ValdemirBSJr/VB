VERSION 5.00
Begin VB.Form FrmJogo 
   Caption         =   "Jogo da velha"
   ClientHeight    =   5460
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6345
   LinkTopic       =   "Form2"
   ScaleHeight     =   5460
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1200
      TabIndex        =   9
      Top             =   3360
      Width           =   3855
      Begin VB.CommandButton Mesa8 
         Height          =   375
         Index           =   8
         Left            =   2400
         TabIndex        =   18
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Mesa7 
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Mesa6 
         Height          =   375
         Index           =   6
         Left            =   960
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Mesa5 
         Height          =   375
         Index           =   5
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Mesa4 
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Mesa3 
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Mesa2 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Mesa1 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Mesa0 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton BtnIniciar 
      Caption         =   "Iniciar"
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Placar"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6015
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lb1PlacarSegundo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Lb1PlacarPrimeiro 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Lb1Segundo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Lb1Primeiro 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Começar jogando o nome que tem o X"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Menu Sair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "FrmJogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Verifica As Boolean
Public Cont, X, y As Byte
Public ganhador, ganhador1 As String

Private Sub BtnIniciarClick()
    Ativa Botoes
    Limpa Mesa
    Cont = 0
    Label1.Caption = "Começa jogando o nome que tem o " & ganhador1
    
End Sub

Private Sub Form_Load()
    Lb1Primeiro.Caption = Primeiro
    Lb1Segundo.Caption = Segundo
    Verifica = True
    Cont = 0
    
End Sub

Private Sub VerificarGanhador()
    If (Mesa0.Caption = "X" And Mesa1.Caption = "X" And Mesa2.Caption = "X") Or (Mesa0.Caption = "0" And Mesa1.Caption = 0 And Mesa2.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        DesativaBotoes
        Somaponto
        ExitSub
    ElseIf (Mesa3.Caption = "X" And Mesa4.Caption = "X" And Mesa5.Caption = "X") Or (Mesa3.Caption = "0" And Mesa4.Caption = "0" And Mesa5.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
    End If
    If Cont = 9 Then
        MsgBox "Não ouve ganhador", vbInformation, "jogo da Velha"
            ganhador1 = "0"
            Verifica = False
        End If
    End Sub
    
    Private Sub Mnujogosair_click()
    End
    End Sub
    
    Private Sub DesativaBotoes()
        Mesa0.Enabled = False
        Mesa1.Enabled = False
        Mesa2.Enabled = False
        Mesa3.Enabled = False
        Mesa4.Enabled = False
        Mesa5.Enabled = False
        Mesa6.Enabled = False
        Mesa7.Enabled = False
        Mesa8.Enabled = False
End Sub
Private Sub ativaBotoes()
        Mesa0.Enabled = True
        Mesa1.Enabled = True
        Mesa2.Enabled = True
        Mesa3.Enabled = True
        Mesa4.Enabled = True
        Mesa5.Enabled = True
        Mesa6.Enabled = True
        Mesa7.Enabled = True
        Mesa8.Enabled = True
End Sub
Private Sub QuemGanhou()
    If X > y Then
        ganhador = "X"
        ganhador1 = "0"
    Else
        ganhador = "0"
        ganhador1 = "X"
    End If
End Sub
Private Sub Somaponto()
    If X > y Then
        Lb1PlacarPrimeiro.Caption = Val(Lb1PlacarPrimeiro.Caption) + 1
    End If
End Sub
Private Sub LimpaMesa()
        Mesa0.Caption = ""
        Mesa1.Caption = ""
        Mesa2.Caption = ""
        Mesa3.Caption = ""
        Mesa4.Caption = ""
        Mesa5.Caption = ""
        Mesa6.Caption = ""
        Mesa7.Caption = ""
        Mesa8.Caption = ""
End Sub
 
    
    
        
        
        
     ElseIf (Mesa6.Caption = "X" And Mesa7.Caption = "X" And Mesa8.Caption = "X") Or (Mesa6.Caption = "0" And Mesa7.Caption = "0" And Mesa8.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
     ElseIf (Mesa0.Caption = "X" And Mesa3.Caption = "X" And Mesa6.Caption = "X") Or (Mesa0.Caption = "0" And Mesa3.Caption = "0" And Mesa6.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
     ElseIf (Mesa1.Caption = "X" And Mesa4.Caption = "X" And Mesa7.Caption = "X") Or (Mesa1.Caption = "0" And Mesa4.Caption = "0" And Mesa7.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
     ElseIf (Mesa2.Caption = "X" And Mesa5.Caption = "X" And Mesa8.Caption = "X") Or (Mesa2.Caption = "0" And Mesa5.Caption = "0" And Mesa8.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
     ElseIf (Mesa0.Caption = "X" And Mesa4.Caption = "X" And Mesa8.Caption = "X") Or (Mesa0.Caption = "0" And Mesa4.Caption = "0" And Mesa8.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
     ElseIf (Mesa2.Caption = "X" And Mesa4.Caption = "X" And Mesa6.Caption = "X") Or (Mesa2.Caption = "0" And Mesa4.Caption = "0" And Mesa6.Caption = "0") Then
        QuemGanhou
        MsgBox "O ganhador foi:" & ganhador
        Desativa Botoes
        Somaponto
        ExitSub
End Sub


Private Sub Mesa0_Click(Index As Integer)
    If Verifica = True Then
        Mesa0.Caption = "X"
        Verifica = False
        Mesa0.Enabled = False
        X = X + 1
    Else
        Mesa0.Caption = "0"
        Verifica = True
        Mesa0.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa1_Click(Index As Integer)
    If Verifica = True Then
        Mesa1.Caption = "X"
        Verifica = False
        Mesa1.Enabled = False
        X = X + 1
    Else
        Mesa1.Caption = "0"
        Verifica = True
        Mesa1.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa2_Click(Index As Integer)
    If Verifica = True Then
        Mesa2.Caption = "X"
        Verifica = False
        Mesa2.Enabled = False
        X = X + 1
    Else
        Mesa2.Caption = "0"
        Verifica = True
        Mesa2.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa3_Click(Index As Integer)
    If Verifica = True Then
        Mesa3.Caption = "X"
        Verifica = False
        Mesa3.Enabled = False
        X = X + 1
    Else
        Mesa3.Caption = "0"
        Verifica = True
        Mesa3.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa4_Click(Index As Integer)
    If Verifica = True Then
        Mesa4.Caption = "X"
        Verifica = False
        Mesa4.Enabled = False
        X = X + 1
    Else
        Mesa4.Caption = "0"
        Verifica = True
        Mesa4.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa5_Click(Index As Integer)
    If Verifica = True Then
        Mesa5.Caption = "X"
        Verifica = False
        Mesa5.Enabled = False
        X = X + 1
    Else
        Mesa5.Caption = "0"
        Verifica = True
        Mesa5.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa6_Click(Index As Integer)
    If Verifica = True Then
        Mesa6.Caption = "X"
        Verifica = False
        Mesa6.Enabled = False
        X = X + 1
    Else
        Mesa6.Caption = "0"
        Verifica = True
        Mesa6.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa7_Click(Index As Integer)
    If Verifica = True Then
        Mesa7.Caption = "X"
        Verifica = False
        Mesa7.Enabled = False
        X = X + 1
    Else
        Mesa7.Caption = "0"
        Verifica = True
        Mesa7.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Mesa8_Click(Index As Integer)
    If Verifica = True Then
        Mesa8.Caption = "X"
        Verifica = False
        Mesa8.Enabled = False
        X = X + 1
    Else
        Mesa8.Caption = "0"
        Verifica = True
        Mesa8.Enabled = False
        y = y + 1
    End If
        Cont = Cont + 1
        VerificaGanhador
End Sub

Private Sub Sair_Click()
    End
End Sub
