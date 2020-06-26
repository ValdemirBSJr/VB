VERSION 5.00
Begin VB.Form FrmInfo 
   Caption         =   "SOBRE"
   ClientHeight    =   2100
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8628
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   8628
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnInfo 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "DÚVIDAS ENTRAR EM CONTATO PELO (83) 9171-2024/ badmoon25@gmail.com"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   5292
   End
   Begin VB.Label LbnInfo 
      AutoSize        =   -1  'True
      Caption         =   "SISTEMA DESENVOLVIDO POR VALDEMIR BEZERRA PARA RENATO."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   7284
   End
End
Attribute VB_Name = "FrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnInfo_Click()
FrmInfo.Hide
End Sub
