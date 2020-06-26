VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmVendas 
   Caption         =   "VENDAS E SERVIÇOS"
   ClientHeight    =   6912
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6912
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtSoma 
      Height          =   372
      Left            =   9960
      TabIndex        =   54
      Top             =   6000
      Width           =   972
   End
   Begin VB.TextBox TxtPagMensal 
      Height          =   372
      Left            =   4680
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Frame Frame3 
      Caption         =   "Botões"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4812
      Left            =   10080
      TabIndex        =   42
      Top             =   240
      Width           =   1092
      Begin VB.CommandButton BtnAchar 
         Caption         =   "Localizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   55
         Top             =   3240
         Width           =   852
      End
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
         Left            =   120
         TabIndex        =   50
         Top             =   4200
         Width           =   852
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
         Left            =   120
         TabIndex        =   49
         Top             =   3720
         Width           =   852
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
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Width           =   852
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
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   852
      End
      Begin VB.CommandButton BtnLocalizar 
         Caption         =   "Visualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   852
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
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   852
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
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   852
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
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   852
      End
   End
   Begin VB.TextBox TxtProRestante 
      Height          =   372
      Left            =   1920
      TabIndex        =   40
      Top             =   2520
      Width           =   1692
   End
   Begin VB.TextBox TxtProTotal 
      Height          =   372
      Left            =   1680
      TabIndex        =   39
      Top             =   2040
      Width           =   1692
   End
   Begin MSMask.MaskEdBox MskProVencimento 
      Height          =   372
      Left            =   6720
      TabIndex        =   38
      Top             =   1440
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskProData 
      Height          =   372
      Left            =   4440
      TabIndex        =   37
      Top             =   1440
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MskProCodigo 
      Height          =   372
      Left            =   1800
      TabIndex        =   36
      Top             =   1440
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comprador"
      Height          =   1212
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   9612
      Begin VB.TextBox TxtPega 
         Height          =   288
         Left            =   4200
         TabIndex        =   53
         Top             =   840
         Width           =   4332
      End
      Begin VB.CommandButton BtnClientes 
         Caption         =   "Buscar Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   41
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente com compra efetuada:"
         Height          =   192
         Left            =   1920
         TabIndex        =   56
         Top             =   840
         Width           =   2136
      End
      Begin VB.Label LblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   372
         Left            =   720
         TabIndex        =   52
         Top             =   360
         Width           =   7452
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   192
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   528
      End
   End
   Begin VB.TextBox Txt4x 
      Height          =   372
      Left            =   7800
      TabIndex        =   13
      Text            =   "0.0"
      Top             =   3840
      Width           =   852
   End
   Begin VB.TextBox Txt2x 
      Height          =   372
      Left            =   5880
      TabIndex        =   11
      Text            =   "0.0"
      Top             =   3840
      Width           =   852
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parcelas"
      Height          =   3132
      Left            =   4680
      TabIndex        =   8
      Top             =   3360
      Width           =   5052
      Begin VB.TextBox Txt24x 
         Height          =   372
         Left            =   3120
         TabIndex        =   33
         Text            =   "0.0"
         Top             =   2400
         Width           =   852
      End
      Begin VB.TextBox Txt23x 
         Height          =   372
         Left            =   2160
         TabIndex        =   32
         Text            =   "0.0"
         Top             =   2400
         Width           =   852
      End
      Begin VB.TextBox Txt22x 
         Height          =   372
         Left            =   1200
         TabIndex        =   31
         Text            =   "0.0"
         Top             =   2400
         Width           =   852
      End
      Begin VB.TextBox Txt21x 
         Height          =   372
         Left            =   240
         TabIndex        =   30
         Text            =   "0.0"
         Top             =   2400
         Width           =   852
      End
      Begin VB.TextBox Txt20x 
         Height          =   372
         Left            =   4080
         TabIndex        =   29
         Text            =   "0.0"
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox Txt19x 
         Height          =   372
         Left            =   3120
         TabIndex        =   28
         Text            =   "0.0"
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox Txt18x 
         Height          =   372
         Left            =   2160
         TabIndex        =   27
         Text            =   "0.0"
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox Txt17x 
         Height          =   372
         Left            =   1200
         TabIndex        =   26
         Text            =   "0.0"
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox Txt16x 
         Height          =   372
         Left            =   240
         TabIndex        =   25
         Text            =   "0.0"
         Top             =   1920
         Width           =   852
      End
      Begin VB.TextBox Txt15x 
         Height          =   372
         Left            =   4080
         TabIndex        =   24
         Text            =   "0.0"
         Top             =   1440
         Width           =   852
      End
      Begin VB.TextBox Txt14x 
         Height          =   372
         Left            =   3120
         TabIndex        =   23
         Text            =   "0.0"
         Top             =   1440
         Width           =   852
      End
      Begin VB.TextBox Txt13x 
         Height          =   372
         Left            =   2160
         TabIndex        =   22
         Text            =   "0.0"
         Top             =   1440
         Width           =   852
      End
      Begin VB.TextBox Txt12x 
         Height          =   372
         Left            =   1200
         TabIndex        =   21
         Text            =   "0.0"
         Top             =   1440
         Width           =   852
      End
      Begin VB.TextBox Txt11x 
         Height          =   372
         Left            =   240
         TabIndex        =   20
         Text            =   "0.0"
         Top             =   1440
         Width           =   852
      End
      Begin VB.TextBox Txt10x 
         Height          =   372
         Left            =   4080
         TabIndex        =   19
         Text            =   "0.0"
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox Txt9x 
         Height          =   372
         Left            =   3120
         TabIndex        =   18
         Text            =   "0.0"
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox Txt8x 
         Height          =   372
         Left            =   2160
         TabIndex        =   17
         Text            =   "0.0"
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox Txt7x 
         Height          =   372
         Left            =   1200
         TabIndex        =   16
         Text            =   "0.0"
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox Txt6x 
         Height          =   372
         Left            =   240
         TabIndex        =   15
         Text            =   "0.0"
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox Txt5x 
         Height          =   372
         Left            =   4080
         TabIndex        =   14
         Text            =   "0.0"
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox Txt3x 
         Height          =   372
         Left            =   2160
         TabIndex        =   12
         Text            =   "0.0"
         Top             =   480
         Width           =   852
      End
      Begin VB.TextBox Txt1x 
         Height          =   372
         Left            =   240
         TabIndex        =   10
         Text            =   "0.0"
         Top             =   480
         Width           =   852
      End
   End
   Begin VB.ComboBox CbbOpcao 
      Height          =   288
      Left            =   8040
      TabIndex        =   7
      Top             =   2880
      Width           =   1572
   End
   Begin VB.TextBox TxtDescricao 
      Height          =   3132
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   4092
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o número de pagamentos:"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   2880
      Width           =   3048
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Descrição do Serviço/Venda:"
      Height          =   192
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   2124
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Pagamento Restante:"
      Height          =   192
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Pagamento Total:"
      Height          =   192
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1284
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento:"
      Height          =   192
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   888
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data da Compra:"
      Height          =   192
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1224
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código do Serviço:"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1380
   End
End
Attribute VB_Name = "FrmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAchar_Click()
FrmConsulta.Show 1
If BuscaCompras <> "" Then
TBCompras.Seek "=", BuscaCompras
Else
TBCompras.Seek "=", LblCliente.Caption
End If
AtualizaFormulario
End Sub

Private Sub BtnAlterar_Click()
If BtnAlterar.Caption = "Alterar" Then
HabilitaControles
DesativaBotoes
BtnAlterar.Caption = "Confirmar"
BtnAlterar.Enabled = True

BtnSair.Caption = "Cancelar"
MskProCodigo.Enabled = False
TxtDescricao.SetFocus

Else

If MsgBox("Confirma a alteração desta compra?", vbYesNo, "ALTERAÇÃO") = vbYes Then
TBCompras.Edit
AtualizaCampos
TBCompras.Update
End If
BtnAlterar.Caption = "Alterar"
BtnSair.Caption = "Sair"
AtivaBotoes
AtualizaFormulario
DesabilitaControles
End If
End Sub

Private Sub BtnAnterior_Click()

If TBCompras.BOF = False Then
TBCompras.MovePrevious
End If

If TBCompras.BOF = True Then
TBCompras.MoveLast
End If
AtualizaFormulario
End Sub



Private Sub BtnClientes_Click()
FrmCliConsulta.Show 1
TBClientes.Index = "IndNomeCli"
If BuscaCliente = "" Then
BuscaCliente = LblCliente.Caption
End If
LblCliente.Caption = BuscaCliente
TxtPega.Text = LblCliente.Caption
End Sub

Private Sub BtnExcluir_Click()

If MsgBox("Tem certeza que deseja excluir esta venda/serviço?", vbYesNo, "EXCLUSÃO") = vbYes Then
TBCompras.Delete
BtnAnterior_Click
AtivaBotoes
End If
End Sub



Private Sub BtnImprimir_Click()
On Error GoTo erro
If MsgBox("Deseja Imprimir este produto?", vbYesNo, "IMPRESSÃO") = vbNo Then
Exit Sub
End If
Cabecalho ("Ficha individual de Vendas/Serviço")
Printer.FontSize = 14
Printer.Print
Printer.Print "Cliente: "; MskProCliente.Text & " - " & LblCliente.Caption
Printer.Print "Código do produto: "; MskProCodigo.Text
Printer.Print "Descrição da venda/Serviço: "; TxtDescricao.Text
Printer.Print
Printer.Print "Data da Compra: "; MskProData.Text
Printer.Print
Printer.Print "Data de Vencimento: "; MskProVencimento.Text
Printer.Print
Printer.Print "Valor Total: "; TxtProTotal.Text
Printer.Print
Printer.Print "Valor Pago: "; TxtProRestante.Text
Printer.Print
Printer.Print "Parcelas: "; Txt1x.Text & " / " & Txt2x.Text & " /" & Txt3x.Text & " / " & Txt4x.Text; Txt5x.Text & " / " & Txt6x.Text
Printer.Print
Printer.Print "Parcelas: "; Txt7x.Text & " / " & Txt8x.Text & " /" & Txt9x.Text & " / " & Txt10x.Text; Txt11x.Text & " / " & Txt12x.Text
Printer.Print
Printer.Print "Parcelas: "; Txt13x.Text & " / " & Txt14x.Text & " /" & Txt15x.Text & " / " & Txt16x.Text; Txt17x.Text & " / " & Txt18x.Text
Printer.Print
Printer.Print "Parcelas: "; Txt19x.Text & " / " & Txt20x.Text & " /" & Txt21x.Text & " / " & Txt22x.Text; Txt23x.Text & " / " & Txt24x.Text
Printer.EndDoc
erro_Click:
Exit Sub
erro:
MsgBox "Ocorreu um erro ao imprimir!", vbCritical, "IMPRESSÃO"
Exit Sub
End Sub


Private Sub BtnInserir_Click()
If LblCliente.Caption = "" Then
LblCliente.Caption = "CLIENTE NÃO DEFINIDO"
End If
If BtnInserir.Caption = "Inserir" Then
HabilitaControles
DesativaBotoes
LimpaFormulario
BtnInserir.Enabled = True
BtnInserir.Caption = "Confirmar"
BtnSair.Enabled = True
BtnSair.Caption = "Cancelar"
MskProData.Text = Date
If TBCompras.RecordCount = 0 Then
MskProCodigo.Text = "000001"
TxtDescricao.SetFocus
Else
TBCompras.Index = "IndCodCom"
TBCompras.MoveLast
MskProCodigo.Text = Format(Val(TBCompras("Codigo_compra")) + 1, "000000")
TxtDescricao.SetFocus
End If
Else
If MsgBox("Tem certeza que deseja gravar esta Venda/Serviço?", vbYesNo, "INCLUSÃO") = vbYes Then
TBCompras.AddNew
AtualizaCampos
TBCompras.Update
End If
BtnInserir.Caption = "Inserir"
BtnSair.Caption = "Sair"
AtualizaFormulario
DesabilitaControles
AtivaBotoes
End If
End Sub

Private Sub BtnLocalizar_Click()
FrmProConsulta.Show 1
TBCompras.Index = "IndCodCom"
If BuscaCompras <> "" Then
TBCompras.Seek "=", BuscaCompras
Else
TBCompras.Seek "=", MskProCodigo.Text
End If
AtualizaFormulario
End Sub

Private Sub BtnProximo_Click()

If TBCompras.EOF = False Then
TBCompras.MoveNext
End If
If TBCompras.EOF = True Then
TBCompras.MoveFirst
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
BtnAlterar.Caption = "Alterar"
BtnInserir.Caption = "Inserir"
BtnSair.Caption = "Sair"
End If
End Sub



Private Sub CbbOpcao_LostFocus()
If CbbOpcao.Text = "1X" Then
Txt1x.Enabled = True
Txt2x.Enabled = False
Txt3x.Enabled = False
Txt4x.Enabled = False
Txt5x.Enabled = False
Txt6x.Enabled = False
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "2X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = False
Txt4x.Enabled = False
Txt5x.Enabled = False
Txt6x.Enabled = False
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "3X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = False
Txt5x.Enabled = False
Txt6x.Enabled = False
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "4X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = False
Txt6x.Enabled = False
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "5X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = False
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "6X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "7X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "8X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "9X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "10X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "11X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "12X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "13X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "14X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "15X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "16X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "17X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "18X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "19X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "20X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = True
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "21X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = True
Txt21x.Enabled = True
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "22X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = True
Txt21x.Enabled = True
Txt22x.Enabled = True
Txt23x.Enabled = False
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "23X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = True
Txt21x.Enabled = True
Txt22x.Enabled = True
Txt23x.Enabled = True
Txt24x.Enabled = False
End If
If CbbOpcao.Text = "24X" Then
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = True
Txt21x.Enabled = True
Txt22x.Enabled = True
Txt23x.Enabled = True
Txt24x.Enabled = True
End If
End Sub

Private Sub Form_Load()
CbbOpcao.AddItem "1X"
CbbOpcao.AddItem "2X"
CbbOpcao.AddItem "3X"
CbbOpcao.AddItem "4X"
CbbOpcao.AddItem "5X"
CbbOpcao.AddItem "6X"
CbbOpcao.AddItem "7X"
CbbOpcao.AddItem "8X"
CbbOpcao.AddItem "9X"
CbbOpcao.AddItem "10X"
CbbOpcao.AddItem "11X"
CbbOpcao.AddItem "12X"
CbbOpcao.AddItem "13X"
CbbOpcao.AddItem "14X"
CbbOpcao.AddItem "15X"
CbbOpcao.AddItem "16X"
CbbOpcao.AddItem "17X"
CbbOpcao.AddItem "18X"
CbbOpcao.AddItem "19X"
CbbOpcao.AddItem "20X"
CbbOpcao.AddItem "21X"
CbbOpcao.AddItem "22X"
CbbOpcao.AddItem "23X"
CbbOpcao.AddItem "24X"

TBCompras.Index = "IndNomeCom"
If TBCompras.RecordCount > 0 Then
AtualizaFormulario
End If
DesabilitaControles
AtivaBotoes


End Sub

Private Sub AtualizaFormulario()
If TBCompras.RecordCount > 0 Then
TxtPega.Text = TBCompras("Nome")
TxtDescricao.Text = TBCompras("Descricao")
MskProData.Text = TBCompras("Data_compra")
MskProVencimento.Text = TBCompras("Vencimento")
TxtProTotal.Text = TBCompras("PagamentoTotal")
TxtProRestante.Text = TBCompras("PagamentoRestante")
Txt1x.Text = TBCompras("1x")
Txt2x.Text = TBCompras("2x")
Txt3x.Text = TBCompras("3x")
Txt4x.Text = TBCompras("4x")
Txt5x.Text = TBCompras("5x")
Txt6x.Text = TBCompras("6x")
Txt7x.Text = TBCompras("7x")
Txt8x.Text = TBCompras("8x")
Txt9x.Text = TBCompras("9x")
Txt10x.Text = TBCompras("10x")
Txt11x.Text = TBCompras("11x")
Txt12x.Text = TBCompras("12x")
Txt13x.Text = TBCompras("13x")
Txt14x.Text = TBCompras("14x")
Txt15x.Text = TBCompras("15x")
Txt16x.Text = TBCompras("16x")
Txt17x.Text = TBCompras("17x")
Txt18x.Text = TBCompras("18x")
Txt19x.Text = TBCompras("19x")
Txt20x.Text = TBCompras("20x")
Txt21x.Text = TBCompras("21x")
Txt22x.Text = TBCompras("22x")
Txt23x.Text = TBCompras("23x")
Txt24x.Text = TBCompras("24x")
TxtPagMensal.Text = TBCompras("PagamentoMensal")

Else
LimpaFormulario
End If

End Sub

Private Sub AtualizaCampos()
TBCompras("Nome") = TxtPega.Text
TBCompras("Codigo_compra") = MskProCodigo.Text
TBCompras("Descricao") = TxtDescricao.Text
TBCompras("Data_compra") = MskProData.Text
TBCompras("Vencimento") = MskProVencimento.Text
TBCompras("PagamentoTotal") = TxtProTotal.Text
TBCompras("PagamentoRestante") = TxtProRestante.Text
 TBCompras("1x") = Txt1x.Text
 TBCompras("2x") = Txt2x.Text
 TBCompras("3x") = Txt3x.Text
 TBCompras("4x") = Txt4x.Text
 TBCompras("5x") = Txt5x.Text
 TBCompras("6x") = Txt6x.Text
 TBCompras("7x") = Txt7x.Text
 TBCompras("8x") = Txt8x.Text
 TBCompras("9x") = Txt9x.Text
 TBCompras("10x") = Txt10x.Text
 TBCompras("11x") = Txt11x.Text
 TBCompras("12x") = Txt12x.Text
 TBCompras("13x") = Txt13x.Text
 TBCompras("14x") = Txt14x.Text
 TBCompras("15x") = Txt15x.Text
 TBCompras("16x") = Txt16x.Text
 TBCompras("17x") = Txt17x.Text
 TBCompras("18x") = Txt18x.Text
 TBCompras("19x") = Txt19x.Text
 TBCompras("20x") = Txt20x.Text
 TBCompras("21x") = Txt21x.Text
 TBCompras("22x") = Txt22x.Text
 TBCompras("23x") = Txt23x.Text
TBCompras("24x") = Txt24x.Text
TBCompras("PagamentoMensal") = TxtPagMensal.Text

End Sub

Private Sub LimpaFormulario()
TxtPega.Text = ""
MskProCodigo.Text = "      "
TxtDescricao.Text = ""
MskProData.Text = "  /  /    "
MskProVencimento.Text = "  /  /    "
TxtProTotal.Text = ""
TxtProRestante.Text = ""
Txt1x.Text = ""
Txt2x.Text = ""
Txt3x.Text = ""
Txt4x.Text = ""
Txt5x.Text = ""
Txt6x.Text = ""
Txt7x.Text = ""
Txt8x.Text = ""
Txt9x.Text = ""
Txt10x.Text = ""
Txt11x.Text = ""
Txt12x.Text = ""
Txt13x.Text = ""
Txt14x.Text = ""
Txt15x.Text = ""
Txt16x.Text = ""
Txt17x.Text = ""
Txt18x.Text = ""
Txt19x.Text = ""
Txt20x.Text = ""
Txt21x.Text = ""
Txt22x.Text = ""
Txt23x.Text = ""
Txt24x.Text = ""
TxtPagMensal.Text = ""

End Sub

Private Sub HabilitaControles()
TxtPega.Enabled = True
MskProCodigo.Enabled = True
TxtDescricao.Enabled = True
MskProData.Enabled = True
MskProVencimento.Enabled = True
TxtProTotal.Enabled = True
TxtProRestante.Enabled = True
Txt1x.Enabled = True
Txt2x.Enabled = True
Txt3x.Enabled = True
Txt4x.Enabled = True
Txt5x.Enabled = True
Txt6x.Enabled = True
Txt7x.Enabled = True
Txt8x.Enabled = True
Txt9x.Enabled = True
Txt10x.Enabled = True
Txt11x.Enabled = True
Txt12x.Enabled = True
Txt13x.Enabled = True
Txt14x.Enabled = True
Txt15x.Enabled = True
Txt16x.Enabled = True
Txt17x.Enabled = True
Txt18x.Enabled = True
Txt19x.Enabled = True
Txt20x.Enabled = True
Txt21x.Enabled = True
Txt22x.Enabled = True
Txt23x.Enabled = True
Txt24x.Enabled = True
TxtPagMensal.Enabled = True
BtnClientes.Enabled = True
End Sub

Private Sub DesabilitaControles()
TxtPega.Enabled = False
MskProCodigo.Enabled = False
TxtDescricao.Enabled = False
MskProData.Enabled = False
MskProVencimento.Enabled = False
TxtProTotal.Enabled = False
TxtProRestante.Enabled = False
Txt1x.Enabled = False
Txt2x.Enabled = False
Txt3x.Enabled = False
Txt4x.Enabled = False
Txt5x.Enabled = False
Txt6x.Enabled = False
Txt7x.Enabled = False
Txt8x.Enabled = False
Txt9x.Enabled = False
Txt10x.Enabled = False
Txt11x.Enabled = False
Txt12x.Enabled = False
Txt13x.Enabled = False
Txt14x.Enabled = False
Txt15x.Enabled = False
Txt16x.Enabled = False
Txt17x.Enabled = False
Txt18x.Enabled = False
Txt19x.Enabled = False
Txt20x.Enabled = False
Txt21x.Enabled = False
Txt22x.Enabled = False
Txt23x.Enabled = False
Txt24x.Enabled = False
TxtPagMensal.Enabled = False
BtnClientes.Enabled = False
End Sub

Private Sub DesativaBotoes()

BtnInserir.Enabled = False
BtnAlterar.Enabled = False
BtnExcluir.Enabled = False
BtnLocalizar.Enabled = False
BtnProximo.Enabled = False
BtnAnterior.Enabled = False
BtnImprimir.Enabled = False
End Sub

Private Sub AtivaBotoes()
If TBCompras.RecordCount > 0 Then
BtnInserir.Enabled = True
BtnAlterar.Enabled = True
BtnExcluir.Enabled = True
BtnLocalizar.Enabled = True
BtnProximo.Enabled = True
BtnAnterior.Enabled = True
BtnImprimir.Enabled = True
Else
DesativaBotoes
End If
BtnInserir.Enabled = True
BtnSair.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

BuscaCompras = "      "
End Sub


Private Sub TxtProTotal_LostFocus()
TxtProTotal.Text = Format(TxtProTotal.Text, "R$ #.#0")

End Sub

Private Sub TxtProRestante_LostFocus()
 
TxtProRestante.Text = Format(TxtProRestante.Text, "R$ #.#0")

End Sub

Private Sub Txt1x_LostFocus()
Txt1x.Text = Format(Txt1x.Text, "R$ #.#0")
End Sub
Private Sub Txt2x_LostFocus()
Txt2x.Text = Format(Txt2x.Text, "R$ #.#0")

End Sub
Private Sub Txt3x_LostFocus()
Txt3x.Text = Format(Txt3x.Text, "R$ #.#0")

End Sub

Private Sub Txt4x_LostFocus()
Txt4x.Text = Format(Txt4x.Text, "R$ #.#0")

End Sub
Private Sub Txt5x_LostFocus()
Txt5x.Text = Format(Txt5x.Text, "R$ #.#0")

End Sub
Private Sub Txt6x_LostFocus()
Txt6x.Text = Format(Txt6x.Text, "R$ #.#0")

End Sub
Private Sub Txt7x_LostFocus()
Txt7x.Text = Format(Txt7x.Text, "R$ #.#0")

End Sub
Private Sub Txt8x_LostFocus()
Txt8x.Text = Format(Txt8x.Text, "R$ #.#0")

End Sub
Private Sub Txt9x_LostFocus()
Txt9x.Text = Format(Txt9x.Text, "R$ #.#0")

End Sub
Private Sub Txt10x_LostFocus()
Txt10x.Text = Format(Txt10x.Text, "R$ #.#0")

End Sub
Private Sub Txt11x_LostFocus()
Txt11x.Text = Format(Txt11x.Text, "R$ #.#0")

End Sub
Private Sub Txt12x_LostFocus()
Txt12x.Text = Format(Txt12x.Text, "R$ #.#0")

End Sub
Private Sub Txt13x_LostFocus()
Txt13x.Text = Format(Txt13x.Text, "R$ #.#0")

End Sub
Private Sub Txt14x_LostFocus()
Txt14x.Text = Format(Txt14x.Text, "R$ #.#0")

End Sub
Private Sub Txt15x_LostFocus()
Txt15x.Text = Format(Txt15x.Text, "R$ #.#0")

End Sub
Private Sub Txt16x_LostFocus()
Txt16x.Text = Format(Txt16x.Text, "R$ #.#0")

End Sub
Private Sub Txt17x_LostFocus()
Txt17x.Text = Format(Txt17x.Text, "R$ #.#0")

End Sub
Private Sub Txt18x_LostFocus()
Txt18x.Text = Format(Txt18x.Text, "R$ #.#0")

End Sub
Private Sub Txt19x_LostFocus()
Txt19x.Text = Format(Txt19x.Text, "R$ #.#0")

End Sub
Private Sub Txt20x_LostFocus()
Txt20x.Text = Format(Txt20x.Text, "R$ #.#0")

End Sub
Private Sub Txt21x_LostFocus()
Txt21x.Text = Format(Txt21x.Text, "R$ #.#0")

End Sub
Private Sub Txt22x_LostFocus()
Txt22x.Text = Format(Txt22x.Text, "R$ #.#0")

End Sub
Private Sub Txt23x_LostFocus()
Txt23x.Text = Format(Txt23x.Text, "R$ #.#0")

End Sub

Private Sub Txt24x_LostFocus()
Txt24x.Text = Format(Txt24x.Text, "R$ #.#0")
End Sub

