VERSION 5.00
Begin VB.Form varmiyokmu 
   Caption         =   "��renci Kay�t Sorgusu"
   ClientHeight    =   3540
   ClientLeft      =   2010
   ClientTop       =   1740
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7335
   Begin VB.Frame Frame1 
      Caption         =   "Tc Kimlik Numaras� �le ��renci Sorgula"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "SORGULA"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1055
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaxLength       =   11
         TabIndex        =   1
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Kay�tl� ��renci olup olmad���n� ��renmek i�in Tc kimlik numaras�n� yaz�p SORGULA butonuna bas�n�z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Image mesaj 
      Height          =   3345
      Left            =   3840
      Picture         =   "varmiyokmu.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "varmiyokmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call veri_ac(False, False)
Call tablo_ac("Select * from  ogrenci where nuftckimlik='" & Text1.Text & "'")
If tablo.RecordCount >= 1 Then
Label1.Caption = "BU TC K�ML�K NUMARASI �LE ��RENC� KAYIT YAPILMI�"
mesaj.Picture = LoadPicture(App.Path & "\images\info.jpg")
Else
Label1.Caption = "S�STEMDE BU TC K�ML�K NUMARASINA SAH�P ��RENC� BULUNAMAMAMI�TIR"
mesaj.Picture = LoadPicture(App.Path & "\images\hata.jpg")
End If
End Sub

