VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmProConsulta 
   Caption         =   "Localizar Produtos"
   ClientHeight    =   4284
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8508
   LinkTopic       =   "Form1"
   ScaleHeight     =   4284
   ScaleWidth      =   8508
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   372
      Left            =   3240
      TabIndex        =   6
      Top             =   3840
      Width           =   1332
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1680
      TabIndex        =   5
      Top             =   3840
      Width           =   1332
   End
   Begin VB.CommandButton BtnLocalizar 
      Caption         =   "Localizar"
      Enabled         =   0   'False
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
      Top             =   3840
      Width           =   1332
   End
   Begin MSFlexGridLib.MSFlexGrid MsfProdutos 
      Height          =   2652
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   4678
      _Version        =   393216
      Rows            =   1
      Cols            =   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite a descrição do Produto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8292
      Begin VB.TextBox TxtLocDescricao 
         Height          =   372
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   6972
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   192
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   528
      End
   End
End
Attribute VB_Name = "FrmProConsulta"
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
MsfProdutos.Clear
MsfProdutos.Cols = 3
MsfProdutos.Rows = 1
MsfProdutos.ColWidth(1) = 3500
Do While TBCompras.EOF = False
If InStr(LCase(TBCompras("Nome")), LCase(TxtLocDescricao.Text)) = 1 Then
MsfProdutos.AddItem Chr(9) & TBCompras("Nome") & TBCompras("Data_compra")
End If
TBCompras.MoveNext
Loop
End Sub

Private Sub BtnOK_Click()

BuscaCompras = MsfProdutos.Text
Unload Me
End Sub

Private Sub Form_Load()

BtnLocalizar.Enabled = False
BuscaCompras = ""
CarregaProduto
BtnOK.Enabled = False
End Sub

Private Sub CarregaProduto()
MsfProdutos.Clear
MsfProdutos.Cols = 3
MsfProdutos.Rows = 1
TBCompras.MoveFirst
Do While TBCompras.EOF = False
MsfProdutos.AddItem Chr(9) & TBCompras("Nome") & TBCompras("Data_compra")
TBCompras.MoveNext
Loop
End Sub

Private Sub MsfProdutos_Click()
If MsfProdutos.ColSel = 1 Then
BtnOK.Enabled = True
Else
BtnOK.Enabled = False
End If
End Sub

Private Sub TxtLocDescricao_Change()
If TxtLocDescricao.Text = "" Then
BtnLocalizar.Enabled = False
BtnOK.Enabled = False
MsfProdutos.Clear
MsfProdutos.Cols = 3
MsfProdutos.Rows = 1
Else
BtnLocalizar.Enabled = True
End If
End Sub


