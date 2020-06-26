VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "&Limpar"
      Height          =   495
      Left            =   6960
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdVerificar 
      Caption         =   "&Verificar"
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton CmdSortear 
      Caption         =   "&Sortear"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Numeros Sorteados"
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   7335
      Begin VB.Label Lb1Sorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lb1Sorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lb1Sorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lb1Sorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Lb1Sorteado 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtApostado 
      Height          =   405
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numeros Apostados"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.TextBox TxtApostado 
         Height          =   405
         Index           =   4
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtApostado 
         Height          =   435
         Index           =   3
         Left            =   4200
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtApostado 
         Height          =   405
         Index           =   2
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtApostado 
         Height          =   405
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Lb1Acertos 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   3720
      TabIndex        =   2
      Top             =   3720
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade de Acertos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   2385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdLimpar_Click()
    Dim I, J As Byte
        For I = 0 To 4
            TxtApostado(I).Text = ""
        Next I
        
        For J = O To 4
            Lb1Sorteado(J).Caption = ""
        Next J
End Sub

Private Sub CmdSortear_Click()
    Dim I As Byte
  
    Randomize
    For I = 0 To 4
        Lb1Sorteado(I).Caption = Int(Rnd * 50)
    Next I
    
End Sub

Private Sub CmdVerificar_Click()
    Dim Acertos As Byte
    Dim I, J As Byte
    For I = 0 To 4
        For J = 0 To 4
            If TxtApostado(I).Text = Lb1Sorteado(J).Caption Then
                Acertos = Acertos + 1
            End If
        Next J
    
    Next I
    Lb1Acertos.Caption = Acertos
    
    If Acertos = 3 Then
        MsgBox "Parabéns você fez um terno e ganhou um bom dinheiro.", vbInformation, "Aviso"
    ElseIf Acertos = 4 Then
        MsgBox "Parabéns você acertou uma quadra e ficou rico.", vbInformation, "Aviso"
    
    ElseIf Acertos = 5 Then
        MsgBox "Parabéns você Acertou a QUINA e ficou Milionario.", vbInformation, "Aviso"
    Else
        MsgBox "Que pena você não ganhou nada.", vbInformation, "Aviso"
    
    End If
        
    
            
End Sub

Private Sub Lb1Sorteados_Click(Index As Integer)

End Sub
