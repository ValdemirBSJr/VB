VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jogo da Memória"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4440
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   3960
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   15
      Left            =   4080
      Tag             =   "Vazio"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   14
      Left            =   2760
      Tag             =   "Vazio"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   13
      Left            =   1440
      Tag             =   "Vazio"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   12
      Left            =   120
      Tag             =   "Vazio"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   11
      Left            =   4080
      Tag             =   "Vazio"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   10
      Left            =   2760
      Tag             =   "Vazio"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   9
      Left            =   1440
      Tag             =   "Vazio"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   8
      Left            =   120
      Tag             =   "Vazio"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   7
      Left            =   4080
      Tag             =   "Vazio"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   6
      Left            =   2760
      Tag             =   "Vazio"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   5
      Left            =   1440
      Tag             =   "Vazio"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   4
      Left            =   120
      Tag             =   "Vazio"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Você tem 60 s para vencer"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label LblTempo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   3
      Left            =   4080
      Tag             =   "Vazio"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   2
      Left            =   1440
      Tag             =   "Vazio"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   1
      Left            =   2760
      Tag             =   "Vazio"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image ImgMemo 
      Height          =   495
      Index           =   0
      Left            =   120
      Tag             =   "Vazio"
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Byte
Dim Primeiro As Boolean
Dim Primeiro1 As Boolean
Dim Clicou, Ganhou As Byte
Const NumPecas = 16

Private Sub Form_Load()
    IniciaJogo
End Sub

Private Sub ImgMemo_Click(Index As Integer)
    If ImgMemo(Index).Tag = "Word" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\word.ico")
    ElseIf ImgMemo(Index).Tag = "Tabela" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\table.ico")
    ElseIf ImgMemo(Index).Tag = "Cd" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\cdrom01.ico")
    ElseIf ImgMemo(Index).Tag = "Copa" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\misc34.ico")
    ElseIf ImgMemo(Index).Tag = "Paus" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\misc35.ico")
    ElseIf ImgMemo(Index).Tag = "Ouro" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\misc36.ico")
    ElseIf ImgMemo(Index).Tag = "Preto" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\misc37.ico")
    ElseIf ImgMemo(Index).Tag = "Abrir" Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\folder03.ico")
    End If
    If Primeiro = True Then
        Clicou = ImgMemo(Index).Index
    End If
    If ImgMemo(Clicou).Tag <> ImgMemo(Index).Tag And Primeiro = False Then
        ImgMemo(Index).Picture = LoadPicture(App.Path & "\delphi.ico")
        ImgMemo(Clicou).Picture = LoadPicture(App.Path & "\delphi.ico")
        Primeiro = True
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
    Primeiro1 = True
    Ganhou = 0
    Randomize
    For I = 0 To 15
        ImgMemo(I).Tag = "Vazio"
        ImgMemo(I).Picture = LoadPicture(App.Path & "\delphi.ico")
    Next I
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Word"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\word.ico")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Tabela"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\table.ico")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Cd"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\cdrom01.ico")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Copa"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\misc34.ico")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Paus"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\misc35.ico")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Ouro"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\misc36.ico")
            I = I + 1
        End If
    Loop
    
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Preto"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\misc37.ico")
            I = I + 1
        End If
    Loop
     
    I = 0
    Do While I < 2
        posicao = Int(Rnd * NumPecas)
        If ImgMemo(posicao).Tag = "Vazio" Then
            ImgMemo(posicao).Tag = "Abrir"
            ImgMemo(posicao).Picture = LoadPicture(App.Path & "\Folder03.ico")
            I = I + 1
        End If
    Loop
    
    LblTempo.Caption = "0"
    Timer2.Enabled = True
    
End Sub


Private Sub Timer1_Timer()
    LblTempo.Caption = Val(LblTempo.Caption) + 1
    If LblTempo.Caption = "60" Then
        MsgBox "Você perdeu"
        IniciaJogo
    End If
End Sub

Private Sub Timer2_Timer()
    For I = 0 To 15
        ImgMemo(I).Picture = LoadPicture(App.Path & "\delphi.ico")
    Next I
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub
