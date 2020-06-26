VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   Picture         =   "AULA9.frx":0000
   ScaleHeight     =   7005
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnLimpar 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   3600
      TabIndex        =   20
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton BtnProcessar 
      Caption         =   "Processar"
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox TxtData 
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Lb1Segundos 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   18
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Lb1Minutos 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Lb1Horas 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   16
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Lb1Dias 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Lb1Semanas 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Lb1Meses 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Lb1Anos 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   12
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Lb1DataSistema 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "Tempo em segundos"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Tempo em minutos"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Tempo em horas"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Tempo em dias"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Tempo em Semanas"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Tempo em Meses"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Tempo anos"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "A data atual é"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Digite a data de seu nascimento"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Há quanto tempo estou na terra"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnLimpar_Click()
    TxtData.Text = ""
    Lb1Anos.Caption = ""
    Lb1Meses.Caption = ""
    Lb1Semana.Caption = ""
    Lb1Dias.Caption = ""
    Lb1Horas.Caption = ""
    Lb1Minutos.Caption = ""
    Lb1Segundos.Caption´ = ""
    TxtData.SetFocus
    BtnProcessar.Enabled = False
    
End Sub

Private Sub BtnProcessar_Click()
    Dim Data As Date
    If IsDate(TxtData.Text) = False Then
        MsgBox "Você digitou a data errada!", vbCritical, "Data Inválida"
        TxtData.SetFocus
        Exit Sub
    End If
        Data = TxtData.Text
        Lb1Anos.Caption = DateDiff("yyyy", Data, Date)
        Lb1Meses.Caption = DateDiff("M", Data, Date)
        Lb1Semanas.Caption = DateDiff("W", Data, Date)
        Lb1Horas.Caption = DateDiff("H", Date, Date)
        Lb1Minutos.Caption = DateDiff("N", Data, Date)
        Lb1Segundos.Caption = DateDiff("S", Data, Date)
End Sub

Private Sub Form_Load()
     Lb1DataSistema.Caption = Date

End Sub


Private Sub TxtData_LostFocus()
        
    BtnProcessar.Enabled = True
End Sub

