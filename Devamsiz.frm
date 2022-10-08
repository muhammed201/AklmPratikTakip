VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Devamsiz 
   Caption         =   "Kurs Biterme Belgesi Basýmý"
   ClientHeight    =   10290
   ClientLeft      =   270
   ClientTop       =   420
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Sertifika Basýmý"
      Height          =   7215
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   13335
      Begin VB.Frame Frame5 
         Caption         =   "Yetkili Kiþiler"
         Height          =   2415
         Left            =   5280
         TabIndex        =   51
         Top             =   240
         Width           =   3735
         Begin VB.TextBox muduryardimcisi 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Text            =   "Rahim YAÞA"
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox mudur 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Text            =   "Meryem BÝTER"
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label muduryardimcisilab 
            Alignment       =   2  'Center
            Caption         =   "MÜDÜR YARDIMCISI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "OKUL MÜDÜRÜ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "Yazdýr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "Ön Ýzleme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox ogrnotu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   47
         Top             =   6600
         Width           =   2775
      End
      Begin VB.TextBox ogrsure 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   45
         Top             =   6600
         Width           =   2415
      End
      Begin VB.TextBox ogrbit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         Top             =   6600
         Width           =   2415
      End
      Begin VB.TextBox ogrbas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   6600
         Width           =   2415
      End
      Begin VB.TextBox ogrbelveren 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   40
         Text            =   "Konya-Akþehir Kýz Meslek Lisesi Pratik Kýz Sanat Okulu"
         Top             =   5400
         Width           =   5895
      End
      Begin VB.TextBox ogrbelgeno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   38
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox ogrvertarih 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   35
         Top             =   4200
         Width           =   2415
      End
      Begin VB.TextBox ogrdersinden 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   40
         TabIndex        =   32
         Top             =   3600
         Width           =   5055
      End
      Begin VB.TextBox ogrbolumu 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   30
         Text            =   "Çocuk Geliþimi"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox ogradisoyadi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox ogrdtarih 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox ogrdyeri 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox ogrbabaadi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox ogrsoyadi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox ogradi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Kurs Bitirme Notu-Derecesi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8670
         TabIndex        =   48
         Top             =   6240
         Width           =   2835
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Toplam Kurs Süre (Saat)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5805
         TabIndex        =   46
         Top             =   6240
         Width           =   2565
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Kursa Bitirdiði Tarih"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3180
         TabIndex        =   44
         Top             =   6240
         Width           =   2055
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Kursa Baþladýðý Tarih"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   195
         TabIndex        =   42
         Top             =   6240
         Width           =   2265
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Belgeyi Veren okul :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   5520
         Width           =   2100
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Belge No                 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   37
         Top             =   4920
         Width           =   2070
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "tarihinde verilmiþtir."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4680
         TabIndex        =   36
         Top             =   4320
         Width           =   2025
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "belirten bu belge "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   4320
         Width           =   1830
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Meslek/Eðitimi Kursunu baþarý ile tamamlamýþ olduðunu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5640
         TabIndex        =   33
         Top             =   3720
         Width           =   5745
      End
      Begin VB.Label Label9 
         Caption         =   "Bölümü"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   31
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "'nýn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6240
         TabIndex        =   29
         Top             =   3000
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Yukarýda kimliði yazýlý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   27
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Doðum Tarihi :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Doðum Yeri    :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Babý Adý         :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Soyadý            :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Adý                  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Öðrenci Detaylarý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7320
      TabIndex        =   2
      Top             =   240
      Width           =   6135
      Begin VB.Frame Frame3 
         Caption         =   "Sýralama Ýþlemleri"
         Height          =   2295
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton Option5 
            Caption         =   "Seçili Sýnýfa Göre"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1320
            Width           =   1935
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Sýnýfa göre sýrala"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Soyisime göre sýrala"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Ýsime göre sýrala"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Numaraya göre sýrala"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.TextBox soyadi 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox adi 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox sinifi 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox nosu 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Soyadý"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Adý"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Sýnýfý"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Numarasý"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Öðrenciler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin MSFlexGridLib.MSFlexGrid ogrliste 
         Bindings        =   "Devamsiz.frx":0000
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   4227327
         ForeColorFixed  =   65535
         GridLines       =   2
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Devamsiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sart, turu As String


Private Sub Combo1_Click()
sart = "where sinifi='" & Combo1.Text & "'"
Call ogrlisteyukle("order by adi")
End Sub




Private Sub Command1_Click()

Dim Ex As Excel.Application
Set Ex = New Excel.Application
   Ex.Visible = True   'Görünmesini istemiyorsak False
   'Ex.Workbooks.Add 'Yeni bir çalýþma sayfasý oluþturmak için kullanýrýz.
   'Ex.Workbooks.Open ("D:\Muhammed\ReSearch\nakit1.xls") 'Bunu da eðer hazýr bir excel sayfamýz var da onun üzerinde çalýþmak istiyorsak Workbooks.Add yerine kullanabiliriz.
   Ex.Workbooks.Open (App.Path & "\kursbitirme.xls")
   
    Ex.Sheets("Sayfa1").Range("D7").Value = ogradi
    Ex.Sheets("Sayfa1").Range("D8").Value = UCase(ogrsoyadi)
    Ex.Sheets("Sayfa1").Range("D9").Value = ogrbabaadi
    Ex.Sheets("Sayfa1").Range("D10").Value = ogrdyeri
    Ex.Sheets("Sayfa1").Range("D11").Value = ogrdtarih
    Ex.Sheets("Sayfa1").Range("D14").Value = ogradisoyadi
    Ex.Sheets("Sayfa1").Range("H14").Value = UCase(ogrbolumu)
    Ex.Sheets("Sayfa1").Range("J14").Value = UCase(ogrdersinden)
    Ex.Sheets("Sayfa1").Range("H16").Value = ogrvertarih
    Ex.Sheets("Sayfa1").Range("C20").Value = ogrbelgeno
    Ex.Sheets("Sayfa1").Range("C20").Value = ogrbelgeno
    Ex.Sheets("Sayfa1").Range("D22").Value = UCase(ogrbelveren)
    Ex.Sheets("Sayfa1").Range("B27").Value = ogrbas
    Ex.Sheets("Sayfa1").Range("D27").Value = ogrbit
    Ex.Sheets("Sayfa1").Range("F27").Value = ogrsure
    Ex.Sheets("Sayfa1").Range("H27").Value = UCase(ogrnotu)
    Ex.Sheets("Sayfa1").Range("B31").Value = mudur
    Ex.Sheets("Sayfa1").Range("H31").Value = muduryardimcisi
    Ex.Sheets("Sayfa1").PrintPreview
    
    
End Sub

Private Sub Command2_Click()
Dim Ex As Excel.Application
Set Ex = New Excel.Application
   Ex.Visible = False   'Görünmesini istemiyorsak False
   'Ex.Workbooks.Add 'Yeni bir çalýþma sayfasý oluþturmak için kullanýrýz.
   'Ex.Workbooks.Open ("D:\Muhammed\ReSearch\nakit1.xls") 'Bunu da eðer hazýr bir excel sayfamýz var da onun üzerinde çalýþmak istiyorsak Workbooks.Add yerine kullanabiliriz.
   Ex.Workbooks.Open (App.Path & "\kursbitirme.xls")
   
    Ex.Sheets("Sayfa1").Range("D7").Value = ogradi
    Ex.Sheets("Sayfa1").Range("D8").Value = UCase(ogrsoyadi)
    Ex.Sheets("Sayfa1").Range("D9").Value = ogrbabaadi
    Ex.Sheets("Sayfa1").Range("D10").Value = ogrdyeri
    Ex.Sheets("Sayfa1").Range("D11").Value = ogrdtarih
    Ex.Sheets("Sayfa1").Range("D14").Value = ogradisoyadi
    Ex.Sheets("Sayfa1").Range("H14").Value = UCase(ogrbolumu)
    Ex.Sheets("Sayfa1").Range("J14").Value = UCase(ogrdersinden)
    Ex.Sheets("Sayfa1").Range("H16").Value = ogrvertarih
    Ex.Sheets("Sayfa1").Range("C20").Value = ogrbelgeno
    Ex.Sheets("Sayfa1").Range("C20").Value = ogrbelgeno
    Ex.Sheets("Sayfa1").Range("D22").Value = UCase(ogrbelveren)
    Ex.Sheets("Sayfa1").Range("B27").Value = ogrbas
    Ex.Sheets("Sayfa1").Range("D27").Value = ogrbit
    Ex.Sheets("Sayfa1").Range("F27").Value = ogrsure
    Ex.Sheets("Sayfa1").Range("H27").Value = UCase(ogrnotu)
    Ex.Sheets("Sayfa1").Range("B31").Value = mudur
    Ex.Sheets("Sayfa1").Range("H31").Value = muduryardimcisi
    Ex.Sheets("Sayfa1").PrintOut
   
End Sub

Private Sub Form_Load()
Call ogrlisteyukle("order by ogrno")
Call comboyukle
End Sub

Sub ogrlisteyukle(sirala As String)
Dim X As Integer
Call veri_ac(False, False)
Call tablo_ac("Select * from ogrenci " & sart & sirala)
ogrliste.Cols = 4
ogrliste.Rows = 1

ogrliste.TextMatrix(0, 0) = "NO"
ogrliste.TextMatrix(0, 1) = "SINIF"
ogrliste.TextMatrix(0, 2) = "ADI"
ogrliste.TextMatrix(0, 3) = "SOYADI"

ogrliste.ColWidth(0) = 500
ogrliste.ColWidth(1) = 1500
ogrliste.ColWidth(2) = 2000
ogrliste.ColWidth(3) = 2000
X = 0
Do While Not tablo.EOF
X = X + 1
ogrliste.AddItem ""
ogrliste.TextMatrix(X, 0) = tablo("ogrno")
ogrliste.TextMatrix(X, 1) = tablo("sinifi")
ogrliste.TextMatrix(X, 2) = tablo("adi")
ogrliste.TextMatrix(X, 3) = tablo("soyadi")
tablo.MoveNext
Loop
tablo.Close
veri.Close
End Sub

Private Sub ogradi_Change()
ogradisoyadi = ogradi & " " & UCase(ogrsoyadi)
End Sub

Private Sub ogrliste_Click()
On Local Error Resume Next
Dim X, Y As Integer
Y = ogrliste.Row
For X = 1 To ogrliste.Rows - 1
    ogrliste.Row = X
    ogrliste.Col = 0
    ogrliste.CellBackColor = ogrliste.BackColor
    ogrliste.Col = 1
    ogrliste.CellBackColor = ogrliste.BackColor
    ogrliste.Col = 2
    ogrliste.CellBackColor = ogrliste.BackColor
    ogrliste.Col = 3
    ogrliste.CellBackColor = ogrliste.BackColor
Next

If ogrliste.Row = 0 Then
    ogrliste.Col = 0
    ogrliste.CellBackColor = 12632256
    ogrliste.Col = 1
    ogrliste.CellBackColor = 12632256
    ogrliste.Col = 2
    ogrliste.CellBackColor = 12632256
    ogrliste.Col = 3
    ogrliste.CellBackColor = 12632256
    Exit Sub
End If
    
    ogrliste.Row = Y
    ogrliste.Col = 0
    ogrliste.CellBackColor = 4326608
    ogrliste.Col = 1
    ogrliste.CellBackColor = 4326608
    ogrliste.Col = 2
    ogrliste.CellBackColor = 4326608
    ogrliste.Col = 3
    ogrliste.CellBackColor = 4326608
    secilisatir = Y
    nosu.Text = ogrliste.TextMatrix(ogrliste.Row, 0)
    sinifi.Text = ogrliste.TextMatrix(ogrliste.Row, 1)
    adi.Text = ogrliste.TextMatrix(ogrliste.Row, 2)
    soyadi.Text = ogrliste.TextMatrix(ogrliste.Row, 3)
    Call sertifika
End Sub




Private Sub ogrliste_EnterCell()
On Local Error Resume Next
    nosu.Text = ogrliste.TextMatrix(ogrliste.Row, 0)
    sinifi.Text = ogrliste.TextMatrix(ogrliste.Row, 1)
    adi.Text = ogrliste.TextMatrix(ogrliste.Row, 2)
    soyadi.Text = ogrliste.TextMatrix(ogrliste.Row, 3)
End Sub

Private Sub ogrsoyadi_Change()
ogradisoyadi = ogradi & " " & UCase(ogrsoyadi)
End Sub

Private Sub Option1_Click()
Combo1.Enabled = False
Call ogrlisteyukle("order by ogrno")
End Sub



Private Sub Option2_Click()
Combo1.Enabled = False
Call ogrlisteyukle("order by adi")
End Sub

Private Sub Option3_Click()
Combo1.Enabled = False
Call ogrlisteyukle("order by soyadi")
End Sub

Private Sub Option4_Click()
Combo1.Enabled = False
Call ogrlisteyukle("order by sinifi")
End Sub

Sub kayit()
Call veri_ac(False, False)
Call tablo_ac("Select * from devamsizlik")
MsgBox takvim.Month & " " & takvim.Day & turu
End Sub

Sub sertifika()
Call veri_ac(False, False)
Call tablo_ac("Select * from ogrenci where ogrno='" & ogrliste.TextMatrix(ogrliste.Row, 0) & "'")
If tablo.RecordCount <= 0 Then
    
Else
    ogradi = tablo("adi")
    ogrsoyadi = UCase(tablo("soyadi"))
    ogrbabaadi = tablo("nufbaba")
    ogrdyeri = tablo("nufdogumyeri")
    ogrdtarih = tablo("nufdogumtarihi")
    ogradisoyadi = ogradi & " " & UCase(ogrsoyadi)
    ogrvertarih = Date
    ogrvertarih = Replace(ogrvertarih, ".", "/")
    

End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
    Combo1.Enabled = True
Else
    Combo1.Enabled = False
End If
End Sub


Sub comboyukle()
Call veri_ac(False, False)
Call tablo_ac("select * from siniflar order by sinif")

Do While Not tablo.EOF
Combo1.AddItem tablo("sinif")
tablo.MoveNext
Loop
veri.Close
End Sub
