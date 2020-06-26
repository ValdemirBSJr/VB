VERSION 5.00
Begin VB.Form FrmCliConsulta 
   Caption         =   "Localizar Clientes"
   ClientHeight    =   4824
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4824
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   372
      Left            =   3000
      TabIndex        =   6
      Top             =   4200
      Width           =   1212
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   372
      Left            =   1560
      TabIndex        =   5
      Top             =   4200
      Width           =   1212
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
      Top             =   4200
      Width           =   1212
   End
   Begin VB.ListBox LstClientes 
      Height          =   2736
      ItemData        =   "FrmCliConsulta.frx":0000
      Left            =   120
      List            =   "FrmCliConsulta.frx":0002
      TabIndex        =   3
      Top             =   1200
      Width           =   6012
   End
   Begin VB.Frame Frame1 
      Caption         =   "Digite o nome do Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6012
      Begin VB.TextBox TxtLocNome 
         Height          =   372
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   4932
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   552
      End
   End
End
Attribute VB_Name = "FrmCliConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnCancelar_Click()
BuscaCliente = ""
Unload Me
End Sub

Private Sub BtnLocalizar_Click()
On Error GoTo erro
TBClientes.MoveFirst
LstClientes.Clear
Do While TBClientes.EOF = False
If InStr(LCase(TBClientes("Nome")), LCase(TxtLocNome.Text)) = 1 Then
LstClientes.AddItem TBClientes("Nome")
End If
TBClientes.MoveNext
Loop
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub BtnOK_Click()
BuscaCliente = LstClientes.Text
Unload Me
End Sub

Private Sub Form_Load()
BtnLocalizar.Enabled = False
BtnOK.Enabled = False
BuscaCliente = ""
CarregaClientes
End Sub

Private Sub LstClientes_Click()
BtnOK.Enabled = True
BtnOK.Default = True
End Sub

Private Sub LstClientes_DblClick()
On Error GoTo erro
BtnOK_Click
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub TxtLocNome_Change()
If TxtLocNome.Text = "" Then
CarregaClientes
BtnLocalizar.Enabled = False
BtnOK.Enabled = False
Else
BtnLocalizar.Enabled = True
BtnOK.Enabled = True
BtnLocalizar.Default = True
BtnLocalizar_Click
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
