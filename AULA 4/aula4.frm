VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invasores"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   5520
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000018&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   4200
      Y2              =   4800
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4200
      Picture         =   "aula4.frx":0000
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2040
      Picture         =   "aula4.frx":0442
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   5
      Left            =   4080
      Picture         =   "aula4.frx":0884
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   4
      Left            =   3000
      Picture         =   "aula4.frx":0B8E
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   3
      Left            =   2160
      Picture         =   "aula4.frx":0E98
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "aula4.frx":11A2
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "aula4.frx":14AC
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "aula4.frx":17B6
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lugar As Boolean
Dim Posicao As Integer
Dim Matou As Byte
'essas 3 linhas abaixo são de som
Dim rc As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
    If KeyCode = vbKeyF2 Then
        IniciaJogo
    ElseIf KeyCode = vbKeyF3 Then
        Timer1.Enabled = False
        Timer1.Enabled = False
        MsgBox "Pausa", vbExclamation, "Jogo"
        Timer1.Enabled = True
        Timer1.Enabled = True
    ElseIf KeyCode = vbKeySpace And Line2.Visible = False Then
        rc = sndPlaySound(App.Path & "\Tiro.WAV", SND_ASYNC) 'esse é de som
        'rc = sndPlaySound("c:\windows\tiro.wav", SND_ASYNC) 'ou usa esse
        Line2.Visible = True
        Line2.X1 = Image1.Left + 230
        Line2.X2 = Image1.Left + 230
        Line2.Y1 = Image1.Top - 360
        Line2.Y2 = Image1.Top
        Timer2.Enabled = True
    ElseIf KeyCode = vbKeyRight Then
        If Image1.Left < 6600 Then
            Image1.Left = Image1.Left + 200
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If Image1.Left > 0 Then
            Image1.Left = Image1.Left - 200
        End If
    End If
End Sub

Private Sub Form_Load()
    IniciaJogo
End Sub

Private Sub Timer1_Timer()
    
    For I = 0 To 5
        
        Posicao = Int(Rnd * 100)
        If Alien(I).Top < 6360 Then
            Alien(I).Top = Alien(I).Top + 100
            If Lugar = False Then
                If Alien(I).Left < 4200 Then
                    Alien(I).Left = Alien(I).Left + Posicao
                    Lugar = True
                Else
                    Alien(I).Left = Alien(I).Left - 2000
                    Lugar = True
                End If
            Else
                If Alien(I).Left > 120 Then
                    Alien(I).Left = Alien(I).Left - Posicao
                    Lugar = False
                Else
                    Alien(I).Left = Alien(I).Left + 2000
                    Lugar = False
                End If
            End If
        Else
            Alien(I).Top = -30
        End If
        If Alien(I).Left >= Image1.Left And Alien(I).Left <= Image1.Left + 480 Then
            If Alien(I).Top + 480 >= 5280 And Alien(I).Top + 480 <= 5760 Then
                rc = sndPlaySound(App.Path & "\Perdeu.WAV", SND_ASYNC)
                Image1.Picture = Image2.Picture
                Timer1.Enabled = False
                Timer2.Enabled = False
                If MsgBox("Você perdeu", vbYesNo, "Jogo") = vbYes Then
                    IniciaJogo
                Else
                    Exit Sub
                End If
            End If
        End If
    Next I
End Sub

Private Sub Timer2_Timer()
    Line2.Y1 = Line2.Y1 - 250
    Line2.Y2 = Line2.Y2 - 250
    conta = conta + 1
    If Line2.Y1 < 0 Then
       Line2.Visible = False
       Timer2.Enabled = False
    End If
    For x = 0 To 5
      pega = Alien(x).Left
        If Line2.X1 >= Alien(x).Left And Line2.X2 <= Alien(x).Left + 480 Then
            If Line2.Y1 >= Alien(x).Top And Line2.Y2 <= Alien(x).Top + 680 Then
                If Alien(x).Visible = True Then
                   Line2.Visible = False
                   Matou = Matou + 1
                End If
                Alien(x).Visible = False
                End If
        End If
    Next x
    If Matou = 6 Then
        If MsgBox("você venceu", vbYesNo, "Jogo") = vbYes Then
            Timer1.Enabled = False
            Timer2.Enabled = False
            IniciaJogo
        Else
            Timer1.Enabled = False
            Timer2.Enabled = False
        End If
    End If
End Sub

Private Sub IniciaJogo()
    For I = 0 To 5
        Alien(I).Visible = True
    Next I
    Alien(0).Left = 3120
    Alien(0).Top = 120
    Alien(1).Left = 1800
    Alien(1).Top = 120
    Alien(2).Left = 840
    Alien(2).Top = 1440
    Alien(3).Left = 2160
    Alien(3).Top = 1560
    Alien(4).Left = 3480
    Alien(4).Top = 480
    Alien(5).Left = 600
    Alien(5).Top = 720
    Image1.Picture = Image1.Picture
    Timer1.Enabled = True
    Timer2.Enabled = True
    Matou = 0
End Sub

