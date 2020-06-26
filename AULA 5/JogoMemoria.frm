VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Memoria"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5160
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   3600
   End
   Begin VB.Label Lb1Tempo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Você term 45 s para vencer"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   15
      Left            =   5040
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   14
      Left            =   3480
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   13
      Left            =   1920
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   12
      Left            =   360
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   11
      Left            =   5040
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   10
      Left            =   3480
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   9
      Left            =   1920
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   8
      Left            =   360
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   7
      Left            =   5040
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   6
      Left            =   3480
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   5
      Left            =   1920
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   4
      Left            =   360
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   3
      Left            =   5040
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   2
      Left            =   3360
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   1
      Left            =   1920
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   0
      Left            =   360
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Byte
Dim Primeiro As Boolean
Dim Clicou, Ganhou As Byte
Const NumPecas = 16

Private Sub Form_Load()
    IniciaJogo
End Sub

Private Sub ImgMemo_Click(Index As Integer)
    If ImgMemo(Index).Tag = "Word" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\word.ico")
    ElseIf ImgMemo(Index).Tag = "Tabela" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\table.ico")
    ElseIf ImgMemo(Index).Tag = "Cd" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\cdrom01.ico")
    ElseIf ImgMemo(Index).Tag = "Copa" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc34.ico")
    ElseIf ImgMemo(Index).Tag = "Paus" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc35.ico")
    ElseIf ImgMemo(Index).Tag = "Ouro" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc36.ico")
    ElseIf ImgMemo(Index).Tag = "Preto" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc37.ico")
    ElseIf ImgMemo(Index).Tag = "Abrir" Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\folder03.ico")
    End If
      
    If Primeiro = True Then
        Clicou = ImgMemo(Index).Index
    End If
    If ImgMemo(Clicou).Tag <> ImgMemo(Index).Tag And Primeiro = False Then
        ImgMemo(Index).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\delphi.ico")
 
    ElseIf ImgMemo(Clicou).Tag = ImgMemo(Index).Tag And Primeiro = False Then
        Ganhou = Ganhou + 1
        Primeiro = True
    Else
        Primeiro = False
    End If
    If Ganhou = 8 Then
        MsgBox "Você venceu"
        Timer1.Enabled = False
        IniciaJogo
    End If
End Sub

Private Sub IniciaJogo()
    Primeiro = True
    Ganhou = 0
    Randomize
    For I = 0 To 15
        ImgMemo(I).Tag = "Vazio"
        ImgMemo(I).Enabled = True
        ImgMemo(I).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\delphi.ico")
        Next I
        I = 0
        Do While I < 2
            posicao = Int(Rnd * NumPecas)
            If ImgMemo(posicao).Tag = "Vazio" Then
                ImgMemo(posicao).Tag = "Word"
                ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\word.ico")
            
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Tabela"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\table.ico")
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
         posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Cd"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\cdrom01.ico")
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
         posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Copa"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc34.ico")
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
         posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Paus"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc35.ico")
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
         posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Ouro"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc36.ico")
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
         posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Preto"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\misc37.ico")
        I = I + 1
        End If
        Loop
        I = 0
        Do While I < 2
         posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Abrir"
            ImgMemo(posicao).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\folder03.ico")
        I = I + 1
        End If
        Loop
        Lb1Tempo.Caption = "0"
        Timer2.Enabled = True
    End Sub
    
    Private Sub Timer1_Timer()
        Lb1Tempo.Caption = Val(Lb1Tempo.Caption) + 1
        If Lb1Tempo.Caption = "45" Then
            MsgBox "Você Perdeu"
            IniciaJogo
        End If
        
    End Sub
    
    Private Sub Timer2_Timer()
        For I = 0 To 15
            ImgMemo(I).Picture = LoadPicture("D:\Carlos\Cursos Prepara\Programador\VB6\AULA 5\Memória\delphi.ico")
            Next I
            Timer1.Enabled = True
            Timer2.Enabled = False
        End Sub


