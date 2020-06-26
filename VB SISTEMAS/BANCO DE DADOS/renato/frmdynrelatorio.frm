VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdynrelatorio 
   Caption         =   "Relatório dos Clientes"
   ClientHeight    =   6312
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10584
   LinkTopic       =   "Form1"
   ScaleHeight     =   6312
   ScaleWidth      =   10584
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   372
      Left            =   8280
      TabIndex        =   6
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton cmdTodos 
      Caption         =   "Exibir Todos"
      Height          =   372
      Left            =   6960
      TabIndex        =   5
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   372
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton cmdProcurar 
      Caption         =   "Procurar"
      Height          =   372
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtprocurar 
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2772
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registros"
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10332
      Begin MSDataGridLib.DataGrid dgDados 
         Height          =   5052
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10092
         _ExtentX        =   17801
         _ExtentY        =   8911
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "lblmsg"
      Height          =   192
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   480
   End
End
Attribute VB_Name = "frmdynrelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
      Set conndyn = New Connection
      conndyn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source=" & App.Path & "\dados.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet       OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"
End Sub
Private Sub Form_Activate()
      consulta ("select * from Clientes")
End Sub


Private Sub consulta(sql1 As String)

Set rsdyn = New Recordset
rsdyn.CursorLocation = adUseClient
rsdyn.Open sql1, conndyn, adOpenForwardOnly, adLockReadOnly

Set dgDados.DataSource = rsdyn
dgDados.Refresh

lblmsg.Caption = "(" & rsdyn.RecordCount & ")" & " Item (s) " & " encontrados!"

End Sub


Private Sub cmdProcurar_Click()

On Error GoTo trataerro

If txtprocurar.Text <> "" Then
     consulta ("select * from Clientes where Nome='" & Trim(txtprocurar) & "'")
Else
    MsgBox "Informe um nome válido.", vbInformation
    txtprocurar.SetFocus
End If
Exit Sub

trataerro:
    txtprocurar = "Não foi possível efetuar a ação."
End Sub

Private Sub cmdImprimir_Click()

With rptdinamico
     Set .DataSource = Nothing
     .DataMember = ""
     Set .DataSource = rsdyn.DataSource
     With .Sections("Section1").Controls
          For i = 1 To .Count
              If TypeOf .Item(i) Is RptTextBox Then
               'O datamember deverá sempre ser enquanto estiver criando relatorios dinamicos
               .Item(i).DataMember = ""
               .Item(i).DataField = rsdyn.Fields(i - 1).Name
            End If
       Next i
   End With
.Show
End With
End Sub

Private Sub cmdTodos_Click()
    Form_Activate
End Sub

Private Sub Form_Unload(Cancel As Integer)
   rsdyn.Close
  conndyn.Close
End Sub









