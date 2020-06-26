Attribute VB_Name = "Variaveis"
Public BancoDeDados As Database
Public TBClientes As Recordset
Public TBCompras As Recordset
Public BuscaCliente As String
Public BuscaCompras As String
Public Soma, Pag As Currency
Public conndyn As Connection
Public rsdyn As Recordset
Public Restante As Currency
Public v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19, v20, v21, v22, v23, v24 As Currency

Public Sub AbreArquivo()

Set BancoDeDados = OpenDatabase(App.Path & "\dados.mdb")
Set TBClientes = BancoDeDados.OpenRecordset("Clientes", dbOpenTable)
Set TBCompras = BancoDeDados.OpenRecordset("Compras", dbOpenTable)
End Sub

Public Sub FechaArquivo()

TBClientes.Close
TBCompras.Close
BancoDeDados.Close

End Sub

Function Cabecalho(Titulo As String)

Printer.Print
Printer.Print
Printer.FontName = "arial"
Printer.FontSize = 24
Printer.Print "SISTEMA INTEGRADO DE CONTROLE EMPRESARIAL"
Printer.FontSize = 14
Printer.Print Titulo; Tab(70); Date & "-" & Time
Printer.Print String(80, "-")


End Function






