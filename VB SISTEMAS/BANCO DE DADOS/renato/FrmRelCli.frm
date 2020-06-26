VERSION 5.00
Begin VB.Form FrmRelCli 
   Caption         =   "RELATÓRIO DOS CLIENTES"
   ClientHeight    =   2340
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Encerrar Relatório"
      Height          =   492
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   2532
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Criar"
      Height          =   492
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2412
   End
   Begin VB.Label LblMensagem 
      AutoSize        =   -1  'True
      Caption         =   "Relatório de Clientes"
      Height          =   552
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   1860
   End
End
Attribute VB_Name = "FrmRelCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oExcel As Object
Dim objExlSht As Object
Dim db As Database
Dim Sn As Recordset   ' Recordset do tipo Snapshot

Private Type ExlCell
   row As Long
   col As Long
End Type

Sub CmdExcel_Click()
On Error GoTo Erro
Dim stCell As ExlCell

MousePointer = vbHourglass ' Muda o ponteiro do mouse

Set oExcel = CreateObject("Excel.Application")
oExcel.Workbooks.Add   'inclui o workbook
Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)

Set db = OpenDatabase("C:\BANCO DE DADOS\renato\dados.MDB")
Set Sn = db.OpenRecordset("Clientes", dbOpenSnapshot)

' Inclui os dados a partir da celula A1
stCell.row = 1
stCell.col = 1
CopiarTabelaExcel Sn, objExlSht, stCell

' Salva a planilha
objExlSht.SaveAs "C:\BANCO DE DADOS\relatorios\Relatório_Clientes_" & Format(Date, "mmddyyyy") & "_" & Format(Time, "hhmmss") & ".xlsx"

oExcel.Visible = True
FrmRelCli.Show
Erro_Click:
Exit Sub
Erro:
MsgBox "Não foi possível carregar relatório", vbCritical, "RELATÓRIO"
Exit Sub
End Sub

Private Sub CopiarTabelaExcel(rs As Recordset, ws As Worksheet, StartingCell As ExlCell)
Dim Vetor() As Variant
Dim row As Long, col As Long
Dim fd As Field

rs.MoveLast
ReDim Vetor(rs.RecordCount + 1, rs.Fields.Count)

' Copia as colunas do cabecalho para um vetor
col = 0
For Each fd In rs.Fields
  Vetor(0, col) = fd.Name
  col = col + 1
Next
' copia o rs par um vetor
rs.MoveFirst
For row = 1 To rs.RecordCount - 1
   For col = 0 To rs.Fields.Count - 1
       Vetor(row, col) = rs.Fields(col).Value
       'O Excel não suporta valores NULL em uma célula.
       If IsNull(Vetor(row, col)) Then Vetor(row, col) = ""
   Next
   rs.MoveNext
Next
ws.Range(ws.Cells(StartingCell.row, StartingCell.col), ws.Cells(StartingCell.row + rs.RecordCount + 1, _
StartingCell.col + rs.Fields.Count)).Value = Vetor
End Sub

Private Sub cmdsair_Click()
On Error GoTo Erro
LblMensagem.Caption = "Encerrando o Excel"
LblMensagem.Refresh
objExlSht.Application.Quit

Set objExlSht = Nothing   ' remove a variavel objeto
Set oExcel = Nothing       ' remove a variavel objeto
Set Sn = Nothing             ' reomove a variavel objeto
Set db = Nothing             ' reomove a variavel objeto

MousePointer = vbDefault     ' Restaura o ponteiro do mouse.
LblMensagem.Caption = "Relatório criado com sucesso. Clique em Encerrar Relatório e vá na pasta relatórios "
LblMensagem.Refresh
Erro_Click:
Exit Sub
Erro:
MsgBox "O relatório não foi encerrado corretamente. Feche pelo botão ''Encerrar Relatório'' da próxima vez e verifique se o relatório está atualizado!", vbCritical, "RELATÓRIO"
Exit Sub
End Sub

