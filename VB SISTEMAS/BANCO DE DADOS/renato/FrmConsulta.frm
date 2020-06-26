VERSION 5.00
Begin VB.Form FrmConsulta 
   Caption         =   "Consulta Vendas"
   ClientHeight    =   4104
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   5436
   LinkTopic       =   "Form1"
   ScaleHeight     =   4104
   ScaleWidth      =   5436
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1560
      TabIndex        =   6
      Top             =   3600
      Width           =   1212
   End
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   372
      Left            =   3000
      TabIndex        =   5
      Top             =   3600
      Width           =   1212
   End
   Begin VB.CommandButton BtnLocalizar 
      Caption         =   "Localizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1212
   End
   Begin VB.ListBox LstCli 
      Height          =   2352
      ItemData        =   "FrmConsulta.frx":0000
      Left            =   120
      List            =   "FrmConsulta.frx":0002
      TabIndex        =   3
      Top             =   1080
      Width           =   5172
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite o Nome do Cliente:"
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5172
      Begin VB.TextBox TxtNome 
         Height          =   408
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   4092
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   612
      End
   End
End
Attribute VB_Name = "FrmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCancelar_Click()
BuscaCompras = ""
Unload Me
End Sub

Private Sub BtnLocalizar_Click()
TBCompras.MoveFirst
LstCli.Clear
Do While TBCompras.EOF = False
If InStr(LCase(TBCompras("Nome")), LCase(TxtNome.Text)) = 1 Then
LstCli.AddItem TBCompras("Nome")
End If
TBCompras.MoveNext
Loop
End Sub

Private Sub BtnOK_Click()
BuscaCompras = LstCli.Text
Unload Me
End Sub

Private Sub Form_Load()
BtnLocalizar.Enabled = False
BtnOk.Enabled = False
BuscaCompras = ""
CarregaClientes
End Sub

Private Sub LstCli_Click()
BtnOk.Enabled = True
BtnOk.Default = True
End Sub

Private Sub LstCli_DblClick()
BtnOK_Click
End Sub

Private Sub TxtNome_Change()
If TxtNome.Text = "" Then
CarregaClientes
BtnLocalizar.Enabled = False
BtnOk.Enabled = False
Else
BtnLocalizar.Enabled = True
BtnOk.Enabled = True
BtnLocalizar.Default = True
BtnLocalizar_Click
End If

End Sub

Private Sub CarregaClientes()

LstCli.Clear
TBCompras.MoveFirst
Do While TBCompras.EOF = False
LstCli.AddItem TBCompras("Nome")
TBCompras.MoveNext
Loop

End Sub

