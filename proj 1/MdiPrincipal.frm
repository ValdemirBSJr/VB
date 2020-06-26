VERSION 5.00
Begin VB.MDIForm MdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Controle de Estoque"
   ClientHeight    =   4590
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7605
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu MnuCadastrosClientes 
         Caption         =   "Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuCadastrosFornecedores 
         Caption         =   "Fornecedores"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuCadastrosProdutos 
         Caption         =   "Produtos"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuCadastrosFuncionários 
         Caption         =   "Funcionários"
         Shortcut        =   ^U
      End
      Begin VB.Menu MnuCadastrosSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCadastrosSair 
         Caption         =   "Sair"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu MnuMovimento 
      Caption         =   "Movimento"
      Begin VB.Menu MnuMovimentoEntrada 
         Caption         =   "Entrada de Pedidos"
         Index           =   1
      End
      Begin VB.Menu MnuMovimentoEntrada 
         Caption         =   "Entrada de Produtos"
         Index           =   2
      End
      Begin VB.Menu MnuMovimentoSaída 
         Caption         =   "Saída de Produtos"
      End
   End
   Begin VB.Menu MnuRelatórios 
      Caption         =   "Relatórios"
      Begin VB.Menu MnuRelatóriosFornecedores 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu MnuRelatóriosClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu MnuRelatóriosProdutos 
         Caption         =   "Produtos"
      End
      Begin VB.Menu MnuRelatóriosProFornecedor 
         Caption         =   "Produtos por Fornecedor"
      End
   End
End
Attribute VB_Name = "MdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
AbreArquivo
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
FechaArquivo
End Sub

Private Sub MnuCadastrosClientes_Click()
FrmClientes.Show
End Sub

Private Sub MnuCadastrosSair_Click()
End
End Sub
