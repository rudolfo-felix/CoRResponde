VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CoRResponde"
   ClientHeight    =   10290
   ClientLeft      =   5925
   ClientTop       =   2220
   ClientWidth     =   15645
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000009&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CoRResponde.frx":0000
   ScaleHeight     =   10290
   ScaleWidth      =   15645
   Begin VB.PictureBox pic2 
      Height          =   4575
      Left            =   5520
      ScaleHeight     =   4515
      ScaleWidth      =   4395
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ajuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   9000
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apagar o desenho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   19
      Top             =   9000
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   9360
      Width           =   255
   End
   Begin VB.PictureBox papa 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   13080
      Picture         =   "CoRResponde.frx":DD54
      ScaleHeight     =   1275
      ScaleWidth      =   2280
      TabIndex        =   15
      Top             =   7080
      Width           =   2340
   End
   Begin VB.PictureBox comboio 
      Height          =   1335
      Left            =   240
      Picture         =   "CoRResponde.frx":11AE1
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   14
      Top             =   7080
      Width           =   2295
   End
   Begin VB.PictureBox ferrari 
      Height          =   1335
      Left            =   10560
      Picture         =   "CoRResponde.frx":1302B
      ScaleHeight     =   1275
      ScaleWidth      =   2355
      TabIndex        =   13
      Top             =   7080
      Width           =   2415
   End
   Begin VB.PictureBox ruca 
      Height          =   1335
      Left            =   8880
      Picture         =   "CoRResponde.frx":1502D
      ScaleHeight     =   1275
      ScaleWidth      =   1485
      TabIndex        =   12
      Top             =   7080
      Width           =   1545
   End
   Begin VB.PictureBox lamb 
      Height          =   1335
      Left            =   6360
      Picture         =   "CoRResponde.frx":15E28
      ScaleHeight     =   1275
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   7080
      Width           =   2415
   End
   Begin VB.PictureBox rudy 
      Height          =   1335
      Left            =   2640
      Picture         =   "CoRResponde.frx":17179
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6720
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   8160
      Top             =   1920
   End
   Begin VB.PictureBox monster 
      Height          =   1335
      Left            =   3960
      Picture         =   "CoRResponde.frx":18C19
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   7080
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   5160
      Picture         =   "CoRResponde.frx":1A7AD
      ScaleHeight     =   1305
      ScaleWidth      =   5385
      TabIndex        =   17
      Top             =   8640
      Width           =   5415
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Desenha a granada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Line Line14 
      X1              =   15360
      X2              =   15360
      Y1              =   2760
      Y2              =   6360
   End
   Begin VB.Line Line13 
      X1              =   240
      X2              =   240
      Y1              =   2760
      Y2              =   6360
   End
   Begin VB.Line Line11 
      X1              =   10200
      X2              =   15360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line10 
      X1              =   3000
      X2              =   240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line8 
      X1              =   3000
      X2              =   10200
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line7 
      X1              =   7680
      X2              =   7680
      Y1              =   4560
      Y2              =   2760
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   15360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      X1              =   11280
      X2              =   11280
      Y1              =   6360
      Y2              =   2760
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   15360
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   11280
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   4080
      Y1              =   2760
      Y2              =   6360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   6360
      TabIndex        =   9
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label rudyname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rudy Félix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   8
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pernas do Ruca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   12120
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label papaname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Papa Francisco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   7680
      TabIndex        =   6
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lamborghini"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comboio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ferrari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tens 20 segundos para corresponder as imagens aos respetivos nomes e desenhar a granada."
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   15615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim drawing
Option Explicit

Private lastX As Single
Private lastY As Single

Private Sub Form_Load()
    Picture1.DrawWidth = 3
    Label2.Caption = "20"
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Picture1.Line (lastX, lastY)-(X, Y), RGB(51, 51, 51)
    End If
    lastX = X
    lastY = Y
End Sub
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------

Private Sub comboio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    comboio.Left = comboio.Left + (X - lx)
    comboio.Top = comboio.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub

Private Sub rudy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    rudy.Left = rudy.Left + (X - lx)
    rudy.Top = rudy.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub

Private Sub lamb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    lamb.Left = lamb.Left + (X - lx)
    lamb.Top = lamb.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub

Private Sub ruca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    ruca.Left = ruca.Left + (X - lx)
    ruca.Top = ruca.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub

Private Sub ferrari_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    ferrari.Left = ferrari.Left + (X - lx)
    ferrari.Top = ferrari.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub

Private Sub papa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    papa.Left = papa.Left + (X - lx)
    papa.Top = papa.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub


Private Sub monster_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static lx As Single, ly As Single
  If Button = vbLeftButton Then
    monster.Left = monster.Left + (X - lx)
    monster.Top = monster.Top + (Y - ly)
  Else
    lx = X: ly = Y
  End If
  
End Sub
Private Sub Command2_Click()
 pic2.Visible = True
 pic2.Picture = LoadPicture("F:\PC\Downloads\CoRRespondeFinal\CoRResponde\lol.jpg")
 
 End Sub
 
Private Sub pic2_Click()
 pic2.Visible = False
 
End Sub
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Picture3.DrawWidth = 3
        Picture3.Line (X, Y)-(X, Y), RGB(51, 51, 51)
    End If
 Form2.Show
 Timer2.Enabled = False

End Sub

Private Sub Command1_Click()
 Picture1.Appearance = False
 Picture3.Appearance = False
End Sub
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Private Sub Timer1_Timer()
  If Timer2.Enabled = True Then
   Form3.Show
  End If
 Timer2.Enabled = False
 
End Sub

Private Sub Timer2_Timer()
 Label2.Caption = Val(Label2.Caption) - 1
End Sub

