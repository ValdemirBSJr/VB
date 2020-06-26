VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invasores do Espaço"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   6000
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5280
      Top             =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000018&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   4560
      Y2              =   5160
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5760
      Picture         =   "Invasores.frx":0000
      Top             =   4560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3600
      Picture         =   "Invasores.frx":030A
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   5
      Left            =   4440
      Picture         =   "Invasores.frx":074C
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   4
      Left            =   4800
      Picture         =   "Invasores.frx":0A56
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   3
      Left            =   3240
      Picture         =   "Invasores.frx":0D60
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   2
      Left            =   2400
      Picture         =   "Invasores.frx":106A
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "Invasores.frx":1374
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   0
      Left            =   720
      Picture         =   "Invasores.frx":167E
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
