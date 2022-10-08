VERSION 5.00
Begin VB.Form yeniogrenci 
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   10335
      Begin VB.ComboBox sinifi 
         DataField       =   "sinifi"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cinsiyeti"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Top             =   1320
         Width           =   975
         Begin VB.ComboBox cinsiyet 
            DataField       =   "cinsiyet"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "yeniogrenci.frx":0000
            Left            =   0
            List            =   "yeniogrenci.frx":000A
            TabIndex        =   12
            Text            =   "KIZ"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox bolumu 
         DataField       =   "bolumu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox soyadi 
         DataField       =   "soyadi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox adi 
         DataField       =   "adi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox numarasi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrno"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox osuresi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogretimsure"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   6840
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox dali 
         BackColor       =   &H00FFFFFF&
         DataField       =   "dali"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Öðrenci Ýçin Resim Yükle"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Resmi Kameradan Al"
         Height          =   375
         Left            =   5520
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sýnýfý"
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
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Soyadý"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Adý"
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
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numarasý"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Öðretim Süresi(Yýl)"
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
         Left            =   4680
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dalý"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image vesikalik 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ogrresim"
         DataSource      =   "Data1"
         Height          =   1815
         Left            =   8160
         Picture         =   "yeniogrenci.frx":001A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   7935
      Left            =   10440
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton Command2 
         Caption         =   "Kaydet"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox SSTab1 
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   10275
      TabIndex        =   21
      Top             =   2760
      Width           =   10335
      Begin VB.TextBox velifaks 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velifaks"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   164
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox velicep 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliceptel"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   163
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox veliistel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliistel"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   162
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox velievpk 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velievadresipk"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   161
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox velimeslegi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velimeslegi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   160
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox velisoyadi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velisoyadi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   159
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox velitcno 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velitckimlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   158
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox veliemail 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliemail"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   157
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox velievtel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliistel"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   156
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox veliispk 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliisadresipk"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   155
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox velituru 
         DataField       =   "velituru"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "yeniogrenci.frx":24005C
         Left            =   -72360
         List            =   "yeniogrenci.frx":240069
         TabIndex        =   154
         Text            =   "velituru"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox veliadi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliadi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   153
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "ogralandegistirdi"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "yeniogrenci.frx":240080
         Left            =   -71880
         List            =   "yeniogrenci.frx":24008A
         TabIndex        =   152
         Text            =   "HAYIR"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "ogrnotmatik"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "yeniogrenci.frx":24009B
         Left            =   -71880
         List            =   "yeniogrenci.frx":2400A5
         TabIndex        =   151
         Text            =   "HAYIR"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "ogryetiskurs"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "yeniogrenci.frx":2400B6
         Left            =   -71880
         List            =   "yeniogrenci.frx":2400C0
         TabIndex        =   150
         Text            =   "HAYIR"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogruygulamayapilan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66120
         TabIndex        =   149
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogruygulamakalan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -65520
         TabIndex        =   148
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrkalanetkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -65520
         TabIndex        =   147
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogruygulamaetkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66120
         TabIndex        =   146
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrokudugukitap"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66120
         TabIndex        =   145
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogryapilanetkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66120
         TabIndex        =   144
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrtoplametkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66120
         TabIndex        =   143
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrceptelefon"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66960
         TabIndex        =   142
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrtelefon"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66960
         TabIndex        =   141
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrsehir"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66960
         TabIndex        =   140
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrsemt"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66960
         TabIndex        =   139
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrpostakodu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66960
         TabIndex        =   138
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox kalan 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrstajkalan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -69480
         TabIndex        =   137
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox yapilan 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrstajyapilan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -71040
         TabIndex        =   136
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox staj 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrstajsuresi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72600
         TabIndex        =   135
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox email 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogremail"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   134
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox yab2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogryabanci2"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -70560
         TabIndex        =   133
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox davranis 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrdavranispuani"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   132
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox kulub 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrkulub"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   131
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox yab1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogryabanci1"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   130
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox kartno 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrbankano"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   129
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox veliegitimdurumu 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliegitimdurumu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   128
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox veligelirduzeyi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veligelirduzeyi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   127
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ComboBox Combo5 
         DataField       =   "kaymezunokul"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "yeniogrenci.frx":2400D1
         Left            =   -72240
         List            =   "yeniogrenci.frx":2400D3
         Sorted          =   -1  'True
         TabIndex        =   126
         Text            =   "Combo5"
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaykaytarih"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   125
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaysinavdurumu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   124
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaygirispuani"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   123
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaybursorani"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   122
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Þehir Okulu"
         DataField       =   "kaysehirokul"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72120
         TabIndex        =   121
         Top             =   840
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gündüzlü"
         DataField       =   "kaygunduzlu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -68520
         TabIndex        =   120
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Burslu"
         DataField       =   "kayburslu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -69960
         TabIndex        =   119
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Servis"
         DataField       =   "kayservis"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -69360
         TabIndex        =   118
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Yatýlý"
         DataField       =   "kayyatili"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66840
         TabIndex        =   117
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Yemek"
         DataField       =   "kayyemek"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67320
         TabIndex        =   116
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Geldiði Okul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   100
         Top             =   360
         Width           =   9975
         Begin VB.ComboBox dipgelokul 
            DataField       =   "dipokuladi"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "yeniogrenci.frx":2400D5
            Left            =   2640
            List            =   "yeniogrenci.frx":2400D7
            Sorted          =   -1  'True
            TabIndex        =   108
            Text            =   "Combo5"
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox dipgelbelge 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipbelgecinsi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   107
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox dipgeltarih 
            BackColor       =   &H00FFFFFF&
            DataField       =   "diptarih"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   106
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Özlük Dosyasý Geldi"
            DataField       =   "dipozlukdosyasi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   240
            TabIndex        =   105
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox dipgelsayi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipsayisi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   104
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox dipgeltur 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgelokultur"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   103
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox dipgelokulno 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgelokulno"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   102
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox dipgeldip 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipdipnotu"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   101
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label49 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mezun Olduðu Okul"
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
            TabIndex        =   115
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Belge Cinsi"
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
            TabIndex        =   114
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tarih"
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
            Index           =   0
            Left            =   120
            TabIndex        =   113
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label52 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sayýsý"
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
            TabIndex        =   112
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label53 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Geldiði Okul Türü"
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
            Left            =   5400
            TabIndex        =   111
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label54 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Okul No"
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
            Left            =   5400
            TabIndex        =   110
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label55 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diploma Notu"
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
            Left            =   5400
            TabIndex        =   109
            Top             =   1320
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Gittiði Okul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2895
         Left            =   -74880
         TabIndex        =   68
         Top             =   2520
         Width           =   9975
         Begin VB.TextBox dipgitdipnot 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitdipnotu"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   92
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox dipgitokulno 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitokulno"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   91
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox dipgitdippuan 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitdippuani"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   90
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox dipgitsayi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitsayi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   3840
            TabIndex        =   89
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox dipgittarih 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgittarih"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   88
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox dipgitbelge 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitbelgecinsi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   87
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox dipgitokul 
            DataField       =   "dipgitokuladi"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "yeniogrenci.frx":2400D9
            Left            =   2640
            List            =   "yeniogrenci.frx":2400DB
            Sorted          =   -1  'True
            TabIndex        =   86
            Text            =   "Combo5"
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox dipgitaciklama 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitaciklama"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   85
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Frame Frame6 
            Caption         =   "**Künye Defteri Gönderilen Belge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1095
            Left            =   120
            TabIndex        =   76
            Top             =   1680
            Width           =   6975
            Begin VB.TextBox dipkungonbelge 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyegonbelge"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   120
               TabIndex        =   80
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox dipkuntarihno 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyeistarihno"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1680
               TabIndex        =   79
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox dipkungontarih 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyegontarih"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   3360
               TabIndex        =   78
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox dipkunaltarihno 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyealtarihno"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   5040
               TabIndex        =   77
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Gön. Belge"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label63 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Ýst. Tarihi/No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1680
               TabIndex        =   83
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label64 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Gön. Tarihi"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3360
               TabIndex        =   82
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Al. Tarih/No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5040
               TabIndex        =   81
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Ýþ Yeri Açma Bilgileri"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1095
            Left            =   7200
            TabIndex        =   69
            Top             =   1680
            Width           =   2655
            Begin VB.TextBox Text34 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipisbelgeserino"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1440
               TabIndex        =   72
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox Text35 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipisbelgetarih"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1440
               TabIndex        =   71
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox Text36 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipisbelgeno"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1440
               TabIndex        =   70
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Bel. Seri No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label67 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Belge Tarihi"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label68 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Belge No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   720
               Width           =   1335
            End
         End
         Begin VB.Label Label56 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diploma Notu"
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
            Left            =   5400
            TabIndex        =   99
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Okul No"
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
            Left            =   5400
            TabIndex        =   98
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label58 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diploma Puaný"
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
            Left            =   5400
            TabIndex        =   97
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tarih/Sayý"
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
            Index           =   1
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label60 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Belge Cinsi"
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
            TabIndex        =   95
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label61 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Okul Adý"
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
            TabIndex        =   94
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Açýklama"
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
            TabIndex        =   93
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   10095
         Begin VB.TextBox nufveryer 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufveryer"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   46
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox nufsira 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufsira"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   45
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox nufaile 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufaile"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   44
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox nufcilt 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufcilt"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   43
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox nufmah 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufmah"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   42
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox nufuyruk 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufuyruk"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   41
            Top             =   3480
            Width           =   2415
         End
         Begin VB.TextBox nufilce 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufilce"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   40
            Top             =   3120
            Width           =   2415
         End
         Begin VB.TextBox nufil 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufil"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   39
            Top             =   2760
            Width           =   2415
         End
         Begin VB.ComboBox nufkan 
            DataField       =   "nufkan"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "yeniogrenci.frx":2400DD
            Left            =   3840
            List            =   "yeniogrenci.frx":2400F9
            Sorted          =   -1  'True
            TabIndex        =   38
            Text            =   "Combo5"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.ComboBox nufdini 
            DataField       =   "nufdini"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "yeniogrenci.frx":24013F
            Left            =   1320
            List            =   "yeniogrenci.frx":240149
            Sorted          =   -1  'True
            TabIndex        =   37
            Text            =   "Combo5"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.ComboBox nufmedeni 
            DataField       =   "nufmedeni"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "yeniogrenci.frx":24015B
            Left            =   2640
            List            =   "yeniogrenci.frx":240168
            Sorted          =   -1  'True
            TabIndex        =   36
            Text            =   "Combo5"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox nufdogumtarihi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufdogumtarihi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   35
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox nufdogumyeri 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufdogumyeri"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   34
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox nufana 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufana"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   33
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox nufbaba 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufbaba"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   32
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox nufcuzdanserino 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufcuzdanserino"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   31
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox nufverneden 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufverneden"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   30
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox nufkayitno 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufkayitno"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   29
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox nufvertarih 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufvertarih"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   28
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox nufaskerlik 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufaskerlik"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   27
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox nuftckimlik 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nuftckimlik"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   26
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label Label84 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Verildiði Yer"
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
            Left            =   5280
            TabIndex        =   67
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label83 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sýra No"
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
            Left            =   5280
            TabIndex        =   66
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label82 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aile Sýra No"
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
            Left            =   5280
            TabIndex        =   65
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cilt No"
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
            Left            =   5280
            TabIndex        =   64
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label80 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mahalle/Köy"
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
            Left            =   5280
            TabIndex        =   63
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label79 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Uyruðu"
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
            TabIndex        =   62
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label78 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ýlçe"
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
            TabIndex        =   61
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label77 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ýl"
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
            TabIndex        =   60
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label76 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Kan Grubu"
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
            Left            =   2640
            TabIndex        =   59
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label75 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dini"
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
            TabIndex        =   58
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label74 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Medeni Hali"
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
            TabIndex        =   57
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label73 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Doðum Tarihi"
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
            TabIndex        =   56
            Top             =   1680
            Width           =   2535
         End
         Begin VB.Label Label72 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Doðum Yeri"
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
            Width           =   2535
         End
         Begin VB.Label Label71 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ana Adý"
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
            TabIndex        =   54
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label70 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Baba Adý"
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
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label69 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cüzdan Seri No"
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
            TabIndex        =   52
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label85 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Veriliþ Nedeni"
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
            Left            =   5280
            TabIndex        =   51
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Label86 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Kayýt No"
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
            Left            =   5280
            TabIndex        =   50
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label87 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Veriliþ Tarihi"
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
            Left            =   5280
            TabIndex        =   49
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label Label88 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Askerlik Þubesi"
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
            Left            =   5280
            TabIndex        =   48
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label89 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "T.C. Kimlik No"
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
            Left            =   5280
            TabIndex        =   47
            Top             =   3480
            Width           =   2415
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4575
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   10095
         Begin VB.PictureBox dusunce 
            DataField       =   "dusunce"
            DataSource      =   "Data1"
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3675
            ScaleWidth      =   9675
            TabIndex        =   23
            Top             =   720
            Width           =   9735
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cüzdan Seri No"
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
            TabIndex        =   24
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.PictureBox velievadresi 
         DataField       =   "velievadresi"
         DataSource      =   "Data1"
         Height          =   975
         Left            =   -67200
         ScaleHeight     =   915
         ScaleWidth      =   2355
         TabIndex        =   165
         Top             =   1200
         Width           =   2415
      End
      Begin VB.PictureBox veliisadresi 
         DataField       =   "veliisadresi"
         DataSource      =   "Data1"
         Height          =   975
         Left            =   -72360
         ScaleHeight     =   915
         ScaleWidth      =   2355
         TabIndex        =   166
         Top             =   1200
         Width           =   2415
      End
      Begin VB.PictureBox ogrevadresi 
         DataField       =   "ogrevadresi"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   -66840
         ScaleHeight     =   675
         ScaleWidth      =   1875
         TabIndex        =   167
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Faks"
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
         Left            =   -69720
         TabIndex        =   211
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cep Telefon"
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
         Left            =   -69720
         TabIndex        =   210
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefon"
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
         Left            =   -69720
         TabIndex        =   209
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Posta Kodu"
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
         Left            =   -69720
         TabIndex        =   208
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ev Adresi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -69720
         TabIndex        =   207
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mesleði"
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
         Left            =   -69720
         TabIndex        =   206
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Soyadý"
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
         Left            =   -69720
         TabIndex        =   205
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T.C. Kimlik No"
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
         Left            =   -74880
         TabIndex        =   204
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email"
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
         Left            =   -74880
         TabIndex        =   203
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefon"
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
         Left            =   -74880
         TabIndex        =   202
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Posta Kodu"
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
         Left            =   -74880
         TabIndex        =   201
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ýþ Adresi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         TabIndex        =   200
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Veli Türü"
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
         Left            =   -74880
         TabIndex        =   199
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Veli Adý"
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
         Left            =   -74880
         TabIndex        =   198
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Uygulamalarý Ýzleme Etkinlik Süresi"
         Height          =   375
         Left            =   -68760
         TabIndex        =   197
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Uygulamalarý Ýzleme Etkinlik Süresi"
         Height          =   375
         Left            =   -68760
         TabIndex        =   196
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Okuduðu Kitap Sayýsý"
         Height          =   375
         Left            =   -68760
         TabIndex        =   195
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yapýlan/Kalan Etk."
         Height          =   375
         Left            =   -68760
         TabIndex        =   194
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Top. Sos. Etk. Saati"
         Height          =   375
         Left            =   -68760
         TabIndex        =   193
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cep Telefonu"
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
         Left            =   -68760
         TabIndex        =   192
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefon"
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
         Left            =   -68760
         TabIndex        =   191
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Þehir"
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
         Left            =   -68760
         TabIndex        =   190
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Semt"
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
         Left            =   -68760
         TabIndex        =   189
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Posta Kodu"
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
         Left            =   -68760
         TabIndex        =   188
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ev Adresi"
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
         Left            =   -68760
         TabIndex        =   187
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alan Deðiþtirdi"
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
         Left            =   -74880
         TabIndex        =   186
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Notmatik Kimliði Geçerli"
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
         Left            =   -74880
         TabIndex        =   185
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yetiþtir.Kurs.Katýlacak"
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
         Left            =   -74880
         TabIndex        =   184
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kalan"
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
         Left            =   -70560
         TabIndex        =   183
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yapýlan"
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
         Left            =   -72000
         TabIndex        =   182
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Staj Süresi"
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
         Left            =   -74880
         TabIndex        =   181
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email"
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
         Left            =   -74880
         TabIndex        =   180
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Davranýþ Puaný"
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
         Left            =   -74880
         TabIndex        =   179
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kulüb"
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
         Left            =   -74880
         TabIndex        =   178
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yabancý Dil/2. Yabancý Dil"
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
         Left            =   -74880
         TabIndex        =   177
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banka Kart No"
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
         Left            =   -74880
         TabIndex        =   176
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Eðitim Durumu"
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
         Left            =   -74880
         TabIndex        =   175
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gelir Düzeyi"
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
         Left            =   -69720
         TabIndex        =   174
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mezun Olduðu Okul"
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
         Left            =   -74760
         TabIndex        =   173
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Okuduðu Okul"
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
         Left            =   -74760
         TabIndex        =   172
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kayýt Tarihi"
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
         Left            =   -74760
         TabIndex        =   171
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sýnav Durumu"
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
         Left            =   -74760
         TabIndex        =   170
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Okula Giriþ Puaný"
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
         Left            =   -74760
         TabIndex        =   169
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Burs Oraný"
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
         Left            =   -74760
         TabIndex        =   168
         Top             =   2280
         Width           =   2535
      End
   End
End
Attribute VB_Name = "yeniogrenci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Local Error GoTo hata
If numarasi.Text = Empty Or adi.Text = Empty Or soyadi.Text = Empty Then
MsgBox "Numarasý, Adý yada Soyadý kýsmýný boþ býrakamazsýnýz", vbCritical, "Hata"
Exit Sub
End If
Call veri_ac(False, False)
Call tablo_ac("Select * from ogrenci")
tablo.AddNew
With tablo
!ogrno = numarasi
!adi = adi.Text
!soyadi = soyadi.Text
!sinifi = sinifi.Text
!bolumu = bolumu.Text
!dali = dali.Text
!cinsiyet = cinsiyet.Text
!ogretimsure = Val(osuresi)
!ogrbankano = kartno.Text
!ogryabanci1 = yab1.Text
!ogryabanci2 = yab2.Text
!ogrkulub = kulub.Text
If davranis.Text = "" Or Empty Then
!ogrdavranispuani = Null
Else
!ogrdavranispuani = davranis
End If
!ogremail = email.Text
!ogrstajsuresi = Val(staj.Text)
!ogrstajyapilan = Val(yapilan.Text)
!ogrstajkalan = Val(kalan.Text)
!ogryetiskurs = Combo1.Text
!ogrnotmatik = Combo2.Text
!ogralandegistirdi = Combo3.Text
!ogrevadresi = ogrevadresi.Text
!ogrpostakodu = Text3.Text
!ogrsemt = Text4.Text
!ogrsehir = Text5.Text
!ogrtelefon = Text6.Text
!ogrceptelefon = Text7.Text
!ogrtoplametkinlik = Val(Text8.Text)
!ogryapilanetkinlik = Val(Text9.Text)
!ogrkalanetkinlik = Val(Text12.Text)
!ogrokudugukitap = Val(Text10.Text)
!ogruygulamaetkinlik = Val(Text11.Text)
!ogruygulamayapilan = Val(Text14.Text)
!ogruygulamakalan = Val(Text13.Text)
!veliadi = veliadi.Text
!velituru = velituru.Text
!veliisadresi = veliisadresi.Text
!veliisadresipk = veliispk.Text
!veliistel = veliistel.Text
!veliemail = veliemail.Text
!velitckimlik = velitcno.Text
!veliegitimdurumu = veliegitimdurumu.Text
!velisoyadi = velisoyadi.Text
!velimeslegi = velimeslegi.Text
!velievadresi = velievadresi.Text
!velievadresipk = velievpk.Text
!velievtel = velievtel.Text
!veliceptel = velicep.Text
!velifaks = velifaks.Text
!veligelirduzeyi = veligelirduzeyi.Text
!kaymezunokul = Combo5.Text
!kaysehirokul = Check7.Value
!kaykaytarih = Text1.Text
!kaysinavdurumu = Text2.Text
    If Text15 = "" Or Text15 = Empty Then
        !kaygirispuani = Null
    Else
        !kaygirispuani = Text15
    End If
    If Text16 = Empty Or Text16 = "" Then
        !kaybursorani = Null
    Else
        !kaybursorani = Text16
    End If
!kayburslu = Check9.Value
!kaygunduzlu = Check1.Value
!kayyatili = Check3.Value
!kayservis = Check2.Value
!kayyemek = Check4.Value
!dipokuladi = dipgelokul.Text
!dipbelgecinsi = dipgelbelge.Text
!diptarih = dipgeltarih
!dipsayisi = dipgelsayi
!dipgelokultur = dipgeltur.Text
!dipgelokulno = Val(dipgelokulno)

    If dipgeldip = Empty Or dipgeldip = "" Then
        !dipdipnotu = Null
    Else
        !dipdipnotu = dipgeldip
    End If
!dipgitokuladi = dipgitokul
!dipgitbelgecinsi = dipgitbelge
!dipgittarih = dipgittarih
!dipgitsayi = dipgitsayi
!dipgitaciklama = dipgitaciklama.Text
!dipkunyegonbelge = dipkungonbelge.Text
!dipkunyeistarihno = dipkuntarihno
!dipkunyealtarihno = dipkunaltarihno
!dipgitokulno = Val(dipgitokulno)
    If dipgitdipnot = Empty Or dipgitdipnot = "" Then
        !dipgitdipnotu = Null
    Else
        !dipgitdipnotu = dipgitdipnot
    End If
 
    If dipgitdippuan = Empty Or dipgitdippuan = "" Then
        !dipgitdippuani = Null
    Else
        !dipgitdippuani = dipgitdippuan
    End If

!dipisbelgeserino = Text34
!dipisbelgetarih = Text35
!dipisbelgeno = Text36
!nufcuzdanserino = nufcuzdanserino.Text
!nufbaba = nufbaba.Text
!nufana = nufana.Text
!nufdogumyeri = nufdogumyeri.Text
!nufdogumtarihi = nufdogumtarihi.Text
!nufmedeni = nufmedeni.Text
!nufdini = nufdini.Text
!nufkan = nufkan.Text
!nufil = nufil.Text
!nufilce = nufilce.Text
!nufuyruk = nufuyruk.Text
!nufmah = nufmah.Text
!nufcilt = nufcilt.Text
!nufaile = nufaile.Text
!nufsira = nufsira.Text
!nufveryer = nufveryer.Text
!nufverneden = nufverneden.Text
!nufkayitno = nufkayitno.Text
!nufvertarih = nufvertarih
!nufaskerlik = nufaskerlik.Text
!nuftckimlik = nuftckimlik.Text
!dusunce = dusunce.Text

End With

tablo.Update
MsgBox "Kayýt gerçekleþtirilmiþtir"
tablo.Close
veri.Close
ogrenci.Data1.Refresh
Unload Me
hata:
If Err = 3022 Then
MsgBox "Bu numara ile kayýt bulunmaktadýr. Lütfen kullanýlmayan bir numara seçiniz", vbCritical, "Hata"
End If
End Sub

Private Sub Form_Load()
Call comboyukle
End Sub

Sub comboyukle()
Call veri_ac(False, False)
Call tablo_ac("Select * from siniflar")
Do While Not tablo.EOF
sinifi.AddItem tablo!sinif
tablo.MoveNext
Loop
tablo.Close
veri.Close
End Sub
