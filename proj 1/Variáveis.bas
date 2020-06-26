Attribute VB_Name = "Variáveis"
Public BancoDeDados As Database
Public TBClientes As Recordset
Public TBFornecedores As Recordset
Public TBProdutos As Recordset
Public TBFuncionários As Recordset
Public TBBuscaFornecedor As String
Public BuscaProduto As String
Public BuscaCliente As String

Public Sub AbreArquivo()

Set BancoDeDados = OpenDatabase(App.Path & "\Dados.mdb")
Set TBClientes = BancoDeDados.OpenRecordset("Clientes", dbOpenTable)
Set TBFornecedores = BancoDeDados.OpenRecordset("Fornecedores", dbOpenTable)
Set TBProdutos = BancoDeDados.OpenRecordset("Produtos", dbOpenTable)
Set TBFuncionários = BancoDeDados.OpenRecordset("Funcionários", dbOpenTable)

End Sub

Public Sub FechaArquivo()

TBClientes.Close
TBFuncionários.Close
TBFornecedores.Close
TBProdutos.Close
BancoDeDados.Close

End Sub

Function Cabecalho(Titulo As String)

Printer.Print
Printer.Print
Printer.FontName = "arial"
Printer.FontSize = 24
Printer.Print "Sistema Integrado de Controle de Estoque"
Printer.FontSize = 14
Printer.Print Titulo; Tab(70); Date & "-" & Time
Printer.Print String(80, "-")


End Function
