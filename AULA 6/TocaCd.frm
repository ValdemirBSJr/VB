VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toca CD"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl MMControl1 
      Height          =   570
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   1005
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      EjectEnabled    =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Lb1Numero1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Número de Música"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Lb1Numero 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Numero da Música"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4680
      Picture         =   "TocaCd.frx":0000
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MUSIC PREPARA EPITACIO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arquivo As String
Dim Volta As Boolean
Private Sub Form_Load()
    Volta = False
    Arquivo = Dir("I:\*.*")
    MMControl1.FileName = "I:" & Arquivo
    MMControl1.Command = "open"
    Lb1Numero1.Caption = MMControl1.Track
    MMControl1.Command = "play"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MMControl1.Command = "stop"
    MMControl1.Command = "close"

End Sub


Private Sub MMControl1_NextClick(Cancel As Integer)
    If Lb1Numero.Caption = Lb1Numero1.Caption Then
        MMControl1.Command = "stop"
        MMControl1.Command = "close"
        Volta = False
        MMControl1.FileName = "I:" & Arquivo
        MMControl1.Command = "open"
        MMControl1.Command = "play"
        Lb1Numero.Caption = 2
        Exit Sub
    End If
    
    Lb1Numero.Caption = Val(Lb1Numero.Caption) + 1
End Sub

Private Sub MMCOntrol1_PrevClick(Cancel As Integer)
    If Lb1Numero.Caption = "1" Then
        Exit Sub
    End If
    If Volta = True Then
        Lb1Numero.Caption = Val(Lb1Numero.Caption) - 1
        Volta = False
    Else
        Volta = True
    End If
End Sub

    
    
        
    
        
        
        
        
    



End Sub
