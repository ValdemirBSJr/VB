VERSION 5.00
Begin VB.Form FrmCliConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar Clientes"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton BtnLocalizar 
      Caption         =   "Localizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox LstClientes 
      Height          =   3375
      ItemData        =   "FormConsulta.frx":0000
      Left            =   120
      List            =   "FormConsulta.frx":0002
      TabIndex        =   3
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite um nome para localizar:"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox TxtLocNome 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "FrmCliConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCancelar_CLick()

BuscaCliente = ""
Unload Me
End Sub

Private Sub BtnLocalizar_CLick()

TBClientes.MoveFirst
LstClientes.Clear

Do While TBClientes.EOF = False
    If InStr(LCase(TBClientes("nome")), LCase(TxtLocNome.Text)) = 1 Then
        LstClientes.AddItem TBClientes("Nome")
    End If
TBClientes.MoveNext
Loop
    
End Sub

Private Sub BtnOK_CLick()

BuscaCliente = LstClientes.Text
Unload Me

End Sub

Private Sub Form_Load()

BtnLocalizar.Enabled = False
BtnOk.Enabled = False
BuscaCliente = ""
CarregaClientes
End Sub

Private Sub LstClientes_Click()

BtnOk.Enabled = True
BtnOk.Default = True
End Sub

Private Sub LstClientes_DblClick()

BtnOK_CLick
End Sub

Private Sub TxtLocNome_Change()

If TxtLocNome.Text = "" Then
    CarregaClientes
    BtnLocalizar.Enabled = False
    BtnOk.Enabled = False
Else
    BtnLocalizar.Enabled = True
    BtnOk.Enabled = True
    BtnLocalizar.Default = True
    BtnLocalizar_CLick
End If

End Sub

Private Sub CarregaClientes()

LstClientes.Clear
TBClientes.MoveFirst

Do While TBClientes.EOF = False
LstClientes.AddItem TBClientes("Nome")
TBClientes.MoveNext
Loop

End Sub
