VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4725
   Icon            =   "Calculadora.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "&Limpar"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox TxtResultado 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton CmdSubtrair 
      Caption         =   "-"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton CmdSomar 
      Caption         =   "+"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton CmdDividir 
      Caption         =   "/"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton CmdMultiplicar 
      Caption         =   "X"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox TxtNum2 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox TxtNum1 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Segundo Número"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Primeiro Número"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Valor1 As Single
Dim Valor2 As Single

Private Sub CmdDividir_Click()
    Valor1 = TxtNum1.Text
    Valor2 = TxtNum2.Text
    TxtResultado.Text = Format(Valor1 / Valor2, "R$ #.#0;$ #.#0")
End Sub

Private Sub CmdLimpar_Click()
    TxtResultado.Text = ""
    TxtNum1.Text = ""
    TxtNum2.Text = ""
    TxtNum1.SetFocus
End Sub

Private Sub CmdMultiplicar_Click()
        Valor1 = TxtNum1.Text
    Valor2 = TxtNum2.Text
    TxtResultado.Text = Format(Valor1 * Valor2, "R$ #.#0;$ #.#0")
End Sub

Private Sub CmdSomar_Click()
    Valor1 = TxtNum1.Text
    Valor2 = TxtNum2.Text
    TxtResultado.Text = Format(Valor1 + Valor2, "R$ #.#0;$ #.#0")
End Sub

Private Sub CmdSubtrair_Click()
        Valor1 = TxtNum1.Text
    Valor2 = TxtNum2.Text
    TxtResultado.Text = Format(Valor1 - Valor2, "R$ #.#0;$ #.#0")
End Sub
