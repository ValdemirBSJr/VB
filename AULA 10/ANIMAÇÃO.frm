VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1320
      Top             =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "PRODATA INFORMÁTICA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "ANIMAÇÃO.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Image1.Move Image1.Left + 100, Image1.Top + 50, Image1.Width + 200, Image1.Height + 200
        If Image.Height = 6000 And Image1.Width = 6000 Then
            Timer1.Enabled = False
            Label1.Visible = True
            Label1.Left = 2000
            Label1.Top = 850
            Label1.Height = 600
            Label1.Width = 8000
            Label1.FontSize = 30
            Label1.ForeColor = &HFF&
            Label1.BackColor = &HC0FFFF
        End If
End Sub
