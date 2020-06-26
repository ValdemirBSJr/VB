VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmClientes 
   Caption         =   "Cadastros de Clientes"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   8100
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Menu de Botões"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6360
      TabIndex        =   24
      Top             =   240
      Width           =   1575
      Begin VB.CommandButton BtnSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton BtnImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton BtnAnterior 
         Caption         =   "Anterior"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton BtnProximo 
         Caption         =   "Próximo"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton BtnLocalizar 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton BtnExcluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton BtnAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton BtnInserir 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox MskCliData 
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliRG 
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   6000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliCgc 
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   5400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "###.###.###-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliFax 
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(##)####-####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliFone 
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "(##)####-####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliCEP 
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##.###-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliEstado 
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliCidade 
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliBairro 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliEndereco 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox TxtCliNome 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskCliCodigo 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "RG:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   285
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "CPF:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5400
      Width           =   345
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fone:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "CEP:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cidade:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código: "
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   585
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AtualizaFormulario()

If TBClientes.RecordCount > 0 Then
    MskCliCodigo.Text = TBClientes("Código")
    TxtCliNome.Text = TBClientes("Nome")
    TxtCliEndereco.Text = TBClientes("Endereço")
    TxtCliBairro.Text = TBClientes("Bairro")
    TxtCliCidade.Text = TBClientes("Cidade")
    TxtCliEstado.Text = TBClientes("Estado")
    MskCliCEP.Text = TBClientes("CEP")
    MskCliCgc.Text = TBClientes("CGC")
    MskCliRG.Text = TBClientes("RG")
    MskCliFone.Text = TBClientes("Fone")
    MskCliFax.Text = TBClientes("Fax")
    MskCliData.Text = TBClientes("Data")
Else
    LimpaFormulario
End If
        
End Sub

Private Sub AtualizaCampos()

TBClientes("Código") = MskCliCodigo.Text
TBClientes("Nome") = TxtCliNome.Text
TBClientes("Endereço") = TxtCliEndereco.Text
TBClientes("Bairro") = TxtCliBairro.Text
TBClientes("Cidade") = TxtCliCidade.Text
TBClientes("Estado") = TxtCliEstado.Text
TBClientes("CEP") = MskCliCEP.Text
TBClientes("CGC") = MskCliCgc.Text
TBClientes("RG") = MskCliRG.Text
TBClientes("Fone") = MskCliFone.Text
TBClientes("Fax") = MskCliFax.Text
TBClientes("Data") = MskCliData.Text

End Sub

Private Sub LimpaFormulario()

MskCliCodigo.Text = "     "
    TxtCliNome.Text = ""
    TxtCliEndereco.Text = ""
    TxtCliBairro.Text = ""
    TxtCliCidade.Text = ""
    TxtCliEstado.Text = ""
    MskCliCEP.Text = "  .   -   "
    MskCliCgc.Text = "   .   .   -  "
    MskCliRG.Text = ""
    MskCliFone.Text = "(  )    -    "
    MskCliFax.Text = "(  )    -    "
    MskCliData.Text = "  /  /    "

End Sub

Private Sub HabilitaControles()

MskCliCodigo.Enabled = True
    TxtCliNome.Enabled = True
    TxtCliEndereco.Enabled = True
    TxtCliBairro.Enabled = True
    TxtCliCidade.Enabled = True
    TxtCliEstado.Enabled = True
    MskCliCEP.Enabled = True
    MskCliCgc.Enabled = True
    MskCliRG.Enabled = True
    MskCliFone.Enabled = True
    MskCliFax.Enabled = True
    MskCliData.Enabled = True

End Sub

Private Sub DesabilitaControles()

MskCliCodigo.Enabled = False
    TxtCliNome.Enabled = False
    TxtCliEndereco.Enabled = False
    TxtCliBairro.Enabled = False
    TxtCliCidade.Enabled = False
    TxtCliEstado.Enabled = False
    MskCliCEP.Enabled = False
    MskCliCgc.Enabled = False
    MskCliRG.Enabled = False
    MskCliFone.Enabled = False
    MskCliFax.Enabled = False
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

If TBClientes.RecordCount > 0 Then
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

End Sub

Private Sub BtnSair_Click()

If BtnSair.Caption = "Sair" Then
    Unload Me
Else
    AtualizaFormulario
    
    AtivaBotoes
    BtnInserir.Caption = "Inserir"
    BtnAlterar.Caption = "Alterar"
    BtnSair.Caption = "Sair"
End If
End Sub

Private Sub Form_Load()

    TBClientes.Index = "IndCliNome"
If TBClientes.RecordCount > 0 Then
    AtualizaFormulario
End If
    DesabilitaControles
    AtivaBotoes

End Sub

Private Sub BtnAlterar_Click()

If BtnAlterar.Caption = "Alterar" Then
    HabilitaControles
    DesativaBotoes
    BtnAlterar.Caption = "Confirmar"
    BtnAlterar.Enabled = True
    BtnSair.Caption = "Cancelar"
    MskCliCodigo.Enabled = False
    TxtCliNome.SetFocus
Else
    If MsgBox("Confirma Alteração?", vbYesNo, "Alteração") = vbYes Then
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
End Sub

Private Sub BtnAnterior_CLick()

If TBClientes.BOF = False Then
    TBClientes.MovePrevious
End If
If TBClientes.BOF = True Then
    TBClientes.MoveLast
End If
    AtualizaFormulario
End Sub

Private Sub BtnExcluir_Click()

If MsgBox("Deseja excluir este cliente?", vbYesNo + vbDefaultButton2, "Exclusão") = vbYes Then
    TBClientes.Delete
    BtnAnterior_CLick
    AtivaBotoes
End If

End Sub

Private Sub BtnImprimir_CLick()

Dim Titulo As String
    If MsgBox("Deseja Imprimir este cliente?", vbYesNo, "Impressão") = vbNo Then
        Exit Sub
    End If
    
    Titulo = "Ficha Individual de Clientes"
    Cabecalho (Titulo)
    Printer.FontSize = 14
    Printer.Print
    Printer.Print "Código: "; MskCliCodigo.Text;
    Printer.Print
    Printer.Print "Nome: "; TxtCliNome.Text;
    Printer.Print
    Printer.Print "Endereço: "; TxtCliEndereco.Text;
    Printer.Print
    Printer.Print "Bairro: "; TxtCliBairro.Text;
    Printer.Print
    Printer.Print "Cidade: "; TxtCliCidade.Text;
    Printer.Print
    Printer.Print "Estado: "; TxtCliEstado.Text;
    Printer.Print
    Printer.Print "CEP: "; MskCliCEP.Text;
    Printer.Print
    Printer.Print "RG: "; MskCliRG.Text;
    Printer.Print
    Printer.Print "Telefone: "; MskCliFone.Text;
    Printer.Print
    Printer.Print "Fax: "; MskCliFax.Text;
    Printer.Print
    Printer.Print "CPF: "; MskCliCdc.Text;
    Printer.Print
    Printer.Print "Data de cadastro: "; MskCliData.Text;
    Printer.Print
    Printer.EndDoc
    
End Sub

Private Sub BtnInserir_Click()

If BtnInserir.Caption = "Inserir" Then
    LimpaFormulario
    HabilitaControles
    DesativaBotoes
    MskCliData.Text = Date
    BtnInserir.Enabled = True
    BtnInserir.Caption = "Confirmar"
    BtnSair.Caption = "Cancelar"
    MskCliCodigo.SetFocus
    
Else
    If MskCliCodigo.Text = "" Then
        MsgBox "Você não digitou um código!", vbCritical, "Cadastro de Clientes"
        MskCliCodigo.SetFocus
        Exit Sub
Else
    If MsgBox("Deseja gravar este cliente?", vbYesNo, "Cadastro de Clientes") = vbYes Then
        TBClientes.AddNew
        AtualizaCampos
        TBClientes.Update
        End If
    End If
    
    BtnInserir.Caption = "Inserir"
    BtnSair.Caption = "Sair"
    AtualizaFormulario
    DesabilitaControles
    AtivaBotoes
    End If
    
End Sub

Private Sub BtnLocalizar_CLick()

    FrmCliConsulta.Show 1
If BuscaCliente <> "" Then
    TBClientes.Seek "=", BuscaCliente
Else
    TBClientes.Seek "=", TxtCliNome.Text
End If
AtualizaFormulario

End Sub

Private Sub BtnProximo_CLick()

If TBClientes.EOF = False Then
    TBClientes.MoveNext
End If

If TBClientes.EOF = True Then
    TBClientes.MoveFirst
End If

    AtualizaFormulario
End Sub

Private Sub MskCliCodigo_LostFocus()

MskCliCodigo.Text = Format(MskCliCodigo.Text, "00000")
TBClientes.Index = "IndCliCodigo"
TBClientes.Seek "=", MskCliCodigo.Text
If TBClientes.NoMatch = False Then
    MsgBox "Já existe um cliente com esse código!", vbInformation, "Inclusão"
    MskCliCodigo.Text = ""
    MskCliCodigo.SetFocus
End If
    TBClientes.Index = "IndCliNome"
End Sub

Private Sub MskCliData_LostFocus()

If IsDate(MskCliData.Text) = False Then
    MsgBox "Data Incorreta!", vbInformation, "Inclusão"
    MskCliData.SetFocus
End If
End Sub


