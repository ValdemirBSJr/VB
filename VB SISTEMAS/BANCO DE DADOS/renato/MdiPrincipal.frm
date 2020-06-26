VERSION 5.00
Begin VB.MDIForm MdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "CADASTRO 1.0"
   ClientHeight    =   3396
   ClientLeft      =   192
   ClientTop       =   816
   ClientWidth     =   8076
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuIniciar 
      Caption         =   "Iniciar"
      Begin VB.Menu MnuCadastros 
         Caption         =   "Cadastros"
      End
      Begin VB.Menu MnuVendas 
         Caption         =   "Vendas"
      End
      Begin VB.Menu MnuSair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu MnuRelatorio 
      Caption         =   "Relatório"
      Begin VB.Menu MnuRelCli 
         Caption         =   "Relatório de Clientes"
      End
      Begin VB.Menu MnuRelVendas 
         Caption         =   "Relatório de vendas"
      End
   End
   Begin VB.Menu MnuBackup 
      Caption         =   "BACK UP"
      Begin VB.Menu MnuCB 
         Caption         =   "Criar Back up"
      End
   End
   Begin VB.Menu MnuFerramentas 
      Caption         =   "Ferramentas"
      Begin VB.Menu MnuCalculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu MnuExplorer 
         Caption         =   "Explorer"
      End
   End
   Begin VB.Menu MnuSobre 
      Caption         =   "Sobre"
   End
End
Attribute VB_Name = "MdiPrincipal"
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

Private Sub MDIForm_Load()
Dim DataMax As Date
    DataMax = "01/01/2013"
    
    If Now() > DataMax Then
        MsgBox "Expirou o prazo de validade! Ligue (83)-9171-2024.VBezerra (badmoon25@gmail.com)", vbCritical, "Valdemir Bezerra"
        End
    End If

AbreArquivo

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
FechaArquivo
End Sub



Private Sub MnuCadastros_Click()
FrmClientes.Show
End Sub

Private Sub MnuCalculadora_Click()
On Error GoTo erro
Shell "calc", vbNormalFocus
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível inicializar a calculadora!", vbCritical, "ERRO CALCULADORA"
Exit Sub
End Sub



Private Sub MnuCB_Click()
On Error GoTo erro
FechaArquivo
FileCopy "C:\BANCO DE DADOS\renato\dados.mdb", "C:\BANCO DE DADOS\BACKUP\BC_dados.mdb"
AbreArquivo

MsgBox "Back Up realizado com sucesso! Favor verificar na pasta BACKUP e recorte o banco de dados e o salve em uma pasta de sua preferência!", vbInformation, "BACKUP EFETUADO"
Shell "explorer.exe C:\BANCO DE DADOS\BACKUP", vbNormalFocus
erro_Click:
Exit Sub
erro:
MsgBox "Ocorreu o erro : " & Err.Number & " - " & Err.Description & " Contate o Administrador"
Exit Sub
End Sub

Private Sub MnuExplorer_Click()
On Error GoTo erro
Shell "explorer.exe", vbNormalFocus
erro_Click:
Exit Sub
erro:
MsgBox "Não foi possível iniciar o explorer.", vbCritical, "FALHA EXPLORER"
Exit Sub
End Sub

Private Sub MnuRelCli_Click()

FrmRelCli.Show
End Sub

Private Sub MnuRelVendas_Click()
FrmRelCom.Show
End Sub

Private Sub mnuSair_Click()
End
End Sub

Private Sub MnuSobre_Click()
FrmInfo.Show
End Sub

Private Sub MnuVendas_Click()
FrmVendas.Show
End Sub
