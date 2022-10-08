VERSION 5.00
Begin VB.Form anamenu 
   BackColor       =   &H80000013&
   Caption         =   "Akþehir Anadolu Kýz Meslek ve Kýz Meslek Lisesi Bilgisayar Bölümü-2008"
   ClientHeight    =   3900
   ClientLeft      =   5265
   ClientTop       =   2955
   ClientWidth     =   8325
   Icon            =   "anamenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   8325
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Tercihler"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   240
         Top             =   3000
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Öðrenci Ýþlemleri"
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Sýnýf Ýþlemleri"
         Height          =   735
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
         Caption         =   "Sertifika Basýmý"
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Çýkýþ"
         Height          =   735
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.Label saat 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Saat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label yil 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label gun 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label ay 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "anamenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
'*                              *
'* Programcý:Muh@mmed Zengin    *
'* Yapým Tar: 2008-2009         *
'* Prog. Ýsm: Akml Öðr Otomsyn. *
'*                              *
'*http://www.muhammedzengin.com *
'*   muhammed201@gmail.com      *
'*                              *
'*                              *
'********************************
Private Sub Command1_Click()
ogrenci.Show
End Sub

Private Sub Command2_Click()
Siniflar.Show
End Sub

Private Sub Command3_Click()
Devamsiz.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
ay.Caption = MonthName(Month(Now))
gun.Caption = Day(Now)
yil.Caption = Year(Now)
End Sub

Private Sub Timer1_Timer()
saat.Caption = Time
End Sub
