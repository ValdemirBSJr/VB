VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmClientes 
   Caption         =   "CLIENTES"
   ClientHeight    =   6900
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8196
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8196
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox MskCliCpf 
      Height          =   372
      Left            =   720
      TabIndex        =   30
      Top             =   1440
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliCep 
      Height          =   372
      Left            =   720
      TabIndex        =   21
      Top             =   4440
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##.###-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliData 
      Height          =   372
      Left            =   1560
      TabIndex        =   19
      Top             =   6120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu de Botões"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4212
      Left            =   6480
      TabIndex        =   17
      Top             =   1080
      Width           =   1572
      Begin VB.CommandButton BtnSair 
         Caption         =   "Sair"
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
         Left            =   240
         TabIndex        =   29
         Top             =   3600
         Width           =   1092
      End
      Begin VB.CommandButton BtnImprimir 
         Caption         =   "Imprimir"
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
         Left            =   240
         TabIndex        =   28
         Top             =   3120
         Width           =   1092
      End
      Begin VB.CommandButton BtnAnterior 
         Caption         =   "Anterior"
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
         Left            =   240
         TabIndex        =   27
         Top             =   2640
         Width           =   1092
      End
      Begin VB.CommandButton BtnProximo 
         Caption         =   "Próximo"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1092
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
         Left            =   240
         TabIndex        =   25
         Top             =   1680
         Width           =   1092
      End
      Begin VB.CommandButton BtnExcluir 
         Caption         =   "Excluir"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1092
      End
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
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
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1092
      End
      Begin VB.CommandButton BtnInserir 
         Caption         =   "Inserir"
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
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1092
      End
   End
   Begin MSMask.MaskEdBox MskCliCnpj 
      Height          =   372
      Left            =   720
      TabIndex        =   16
      Top             =   2040
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   18
      Mask            =   "##.###.###/####-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliEmail 
      Height          =   372
      Left            =   720
      TabIndex        =   13
      Top             =   5520
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   656
      _Version        =   393216
      ForeColor       =   -2147483635
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliFone 
      Height          =   372
      Left            =   840
      TabIndex        =   12
      Top             =   5040
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(##)####-####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliEstado 
      Height          =   372
      Left            =   720
      TabIndex        =   11
      Top             =   3960
      Width           =   612
      _ExtentX        =   1080
      _ExtentY        =   656
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliBairro 
      Height          =   372
      Left            =   840
      TabIndex        =   10
      Top             =   3360
      Width           =   4092
      _ExtentX        =   7218
      _ExtentY        =   656
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliEndereco 
      Height          =   372
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   5052
      _ExtentX        =   8911
      _ExtentY        =   656
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliNome 
      Height          =   372
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   5172
      _ExtentX        =   9123
      _ExtentY        =   656
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliCodigo 
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "CEP:"
      Height          =   192
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Data do Cadastro:"
      Height          =   192
      Left            =   120
      TabIndex        =   18
      Top             =   6240
      Width           =   1308
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   192
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   456
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
      Height          =   192
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   348
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "E-mail:"
      Height          =   192
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   492
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Telefone:"
      Height          =   192
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   684
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   192
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   516
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bairro:"
      Height          =   192
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   468
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   192
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   744
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   480
   End
   Begin VB.Label LbnCodigoCli 
      AutoSize        =   -1  'True
      Caption         =   "Código do Cliente:"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1320
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnLocalizar_Click()

FrmCliConsulta.Show 1
If BuscaCliente <> "" Then
TBClientes.Seek "=", BuscaCliente
Else
TBClientes.Seek "=", TxtCliNome.Text
End If
AtualizaFormulario



End Sub


Private Sub AtualizaFormulario()
On Error GoTo erro
If TBClientes.RecordCount > 0 Then
MskCliCodigo.Text = TBClientes("Codigo")
TxtCliNome.Text = TBClientes("Nome")
TxtCliEndereco.Text = TBClientes("Endereco")
MskCliCpf.Text = TBClientes("Cpf")
MskCliCnpj.Text = TBClientes("Cnpj")
TxtCliBairro.Text = TBClientes("Bairro")
TxtCliEstado.Text = TBClientes("Estado")
MskCliCep.Text = TBClientes("Cep")
MskCliFone.Text = TBClientes("Telefone")
MskCliEmail.Text = TBClientes("Email")
MskCliData.Text = TBClientes("Data")
Else
LimpaFormulario
End If
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a atualização.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub AtualizaCampos()
On Error GoTo erro
 TBClientes("Codigo") = MskCliCodigo.Text
 TBClientes("Nome") = TxtCliNome.Text
 TBClientes("Endereco") = TxtCliEndereco.Text
 TBClientes("Cpf") = MskCliCpf.Text
 TBClientes("Cnpj") = MskCliCnpj.Text
 TBClientes("Bairro") = TxtCliBairro.Text
 TBClientes("Estado") = TxtCliEstado.Text
TBClientes("Cep") = MskCliCep.Text
TBClientes("Telefone") = MskCliFone.Text
 TBClientes("Email") = MskCliEmail.Text
 TBClientes("Data") = MskCliData.Text
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub LimpaFormulario()
On Error GoTo erro
MskCliCodigo.Text = "      "
TxtCliNome.Text = ""
TxtCliEndereco.Text = ""
MskCliCpf.Text = "   .   .   -  "
MskCliCnpj.Text = "  .   .   /    -  "
TxtCliBairro.Text = ""
TxtCliEstado.Text = ""
MskCliCep.Text = "  .   -   "
MskCliFone.Text = "(  )    -    "
MskCliEmail.Text = ""
MskCliData.Text = "  /  /    "
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub HabilitaControles()
MskCliCodigo.Enabled = True
TxtCliNome.Enabled = True
TxtCliEndereco.Enabled = True
MskCliCpf.Enabled = True
MskCliCnpj.Enabled = True
TxtCliBairro.Enabled = True
TxtCliEstado.Enabled = True
MskCliCep.Enabled = True
MskCliFone.Enabled = True
MskCliEmail.Enabled = True
MskCliData.Enabled = True

End Sub

Private Sub DesabilitaControles()

MskCliCodigo.Enabled = False
TxtCliNome.Enabled = False
TxtCliEndereco.Enabled = False
MskCliCpf.Enabled = False
MskCliCnpj.Enabled = False
TxtCliBairro.Enabled = False
TxtCliEstado.Enabled = False
MskCliCep.Enabled = False
MskCliFone.Enabled = False
MskCliEmail.Enabled = False
MskCliData.Enabled = False
End Sub

Private Sub DesativaBotoes()

BtnInserir.Enabled = False
BtnAlterar.Enabled = False
BtnExcluir.Enabled = False
BtnImprimir.Enabled = False
BtnProximo.Enabled = False
BtnAnterior.Enabled = False
BtnLocalizar.Enabled = False
End Sub

Private Sub AtivaBotoes()
On Error GoTo erro
If TBClientes.RecordCount > 0 Then ' verifica se há registros na tabela
BtnInserir.Enabled = True
BtnAlterar.Enabled = True
BtnExcluir.Enabled = True
BtnImprimir.Enabled = True
BtnProximo.Enabled = True
BtnAnterior.Enabled = True
BtnLocalizar.Enabled = True
Else
DesativaBotoes
End If
BtnInserir.Enabled = True
BtnSair.Enabled = True
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub Form_Load()
On Error GoTo erro
TBClientes.Index = "IndNomeCli"
If TBClientes.RecordCount > 0 Then
AtualizaFormulario
End If
DesabilitaControles
AtivaBotoes
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub BtnAlterar_Click()
On Error GoTo erro
If BtnAlterar.Caption = "Alterar" Then
HabilitaControles
DesativaBotoes
BtnAlterar.Caption = "Confirmar"
BtnAlterar.Enabled = True
BtnSair.Caption = "Cancelar"
MskCliCodigo.Enabled = False
TxtCliNome.SetFocus
Else

If MsgBox("Confirma alteração?", vbYesNo, "ALTERAÇÃO") = vbYes Then
TBClientes.Edit
AtualizaCampos
TBClientes.Update
End If
BtnAlterar.Caption = "Alterar"
BtnSair.Caption = "Sair"
AtivaBotoes
AtualizaFormulario
DesabilitaControles
End If
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub BtnAnterior_Click()
If TBClientes.BOF = False Then
TBClientes.MovePrevious
End If
If TBClientes.BOF = True Then
TBClientes.MoveLast
End If
AtualizaFormulario
End Sub

Private Sub BtnExcluir_Click()
On Error GoTo erro
If MsgBox("Deseja excluir este permanentemente este Cliente?", vbYesNo + vbDefaultButton2, "EXCLUSÃO") = vbYes Then
TBClientes.Delete
BtnAnterior_Click
AtivaBotoes
End If
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub BtnImprimir_Click()
On erro GoTo erro
Dim Titulo As String

If MsgBox("Deseja imprimir este Cliente?", vbYesNo, "IMPRESSÃO") = vbNo Then
Exit Sub
End If
Titulo = "Ficha Individual de Clientes"
Cabecalho (Titulo)
Printer.FontSize = 14
Printer.Print
Printer.Print "Código: "; MskCliCodigo.Text
Printer.Print
Printer.Print "Nome: "; TxtCliNome.Text
Printer.Print
Printer.Print "Endereço: "; TxtCliEndereco.Text
Printer.Print
Printer.Print "Bairro: "; TxtCliBairro.Text
Printer.Print
Printer.Print "Cidade: "; TxtCliCidade.Text
Printer.Print
Printer.Print "Estado: "; TxtCliEstado.Text
Printer.Print
Printer.Print "Cep: "; MskCliCep.Text
Printer.Print
Printer.Print "CPF: "; MskCliCpf.Text
Printer.Print
Printer.Print "Telefone: "; MskCliFone.Text
Printer.Print
Printer.Print "Data: "; MskCliData.Text
Printer.Print
Printer.Print "E-mail: "; TxtCliEmail.Text
Printer.EndDoc
erro_Click:
Exit Sub
erro:
MsgBox "Ocorreu um erro ao imprimir!", vbCritical, "IMPRESSÃO"
Exit Sub
End Sub

Private Sub BtnInserir_Click()
On Error GoTo erro
If BtnInserir.Caption = "Inserir" Then
HabilitaControles
DesativaBotoes
LimpaFormulario
MskCliData.Text = Date
BtnInserir.Enabled = True
BtnInserir.Caption = "Confirmar"
BtnSair.Caption = "Cancelar"
'Colocando o campo de preenchimento automatico abaixo, experimental.
If TBClientes.RecordCount = 0 Then
MskCliCodigo.Text = "000001"
TxtCliNome.SetFocus
Else

TBClientes.Index = "IndCodCli"
TBClientes.MoveLast
MskCliCodigo.Text = Format(Val(TBClientes("Codigo")) + 1, "000000")
TxtCliNome.SetFocus
End If
Else
If MsgBox("Deseja gravar este Cliente?", vbYesNo, "CADASTRO DE CLIENTES") = vbYes Then
TBClientes.AddNew
AtualizaCampos
TBClientes.Update
End If

BtnInserir.Caption = "Inserir"
BtnSair.Caption = "Sair"
AtualizaFormulario
DesabilitaControles
AtivaBotoes
End If
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a tarefa.", vbCritical, "ERRO"
Exit Sub

End Sub

Private Sub BtnProximo_Click()

If TBClientes.EOF = False Then
TBClientes.MoveNext
End If
If TBClientes.EOF = True Then
TBClientes.MoveFirst
End If
AtualizaFormulario

End Sub

Private Sub BtnSair_Click()

If BtnSair.Caption = "Sair" Then
Unload Me
Else
AtualizaFormulario
DesabilitaControles
AtivaBotoes
BtnInserir.Caption = "Inserir"
BtnAlterar.Caption = "Alterar"
BtnSair.Caption = "Sair"
End If

End Sub

Private Sub MskCliCodigo_LostFocus()

If MskCliCodigo = "" Then
MsgBox "Insira um código válido!", vbInformation, "AVISO"
MskCliCodigo.SetFocus
End If
End Sub

Private Sub MskCliData_LostFocus()

If IsDate(MskCliData.Text) = False Then
MsgBox "Data Incorreta", vbInformation, "AVISO"
MskCliData.SetFocus
End If

End Sub

































