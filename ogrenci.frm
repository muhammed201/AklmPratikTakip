VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ogrenci 
   Caption         =   "Öðrenci Bilgileri"
   ClientHeight    =   9300
   ClientLeft      =   2700
   ClientTop       =   1110
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11895
   Begin VB.CommandButton Command8 
      Caption         =   "BUL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   254
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox bul 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   253
      Text            =   "Öðrenci No"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SORGULA"
      Height          =   375
      Left            =   10680
      TabIndex        =   61
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame10 
      Height          =   855
      Left            =   240
      TabIndex        =   54
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton son 
         Caption         =   ">>"
         Height          =   495
         Left            =   1920
         TabIndex        =   58
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton sonraki 
         Caption         =   ">"
         Height          =   495
         Left            =   1320
         TabIndex        =   57
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton onceki 
         Caption         =   "<"
         Height          =   495
         Left            =   720
         TabIndex        =   56
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton ilk 
         Caption         =   "<<"
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Data Data1 
      BackColor       =   &H000080FF&
      Caption         =   "<<Geri-Ýleri>>"
      Connect         =   "Access 2000;"
      DatabaseName    =   "veri_Backup.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ogrenci"
      Top             =   8400
      Width           =   2745
   End
   Begin VB.Frame Frame3 
      Height          =   7935
      Left            =   10680
      TabIndex        =   51
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton kaydet 
         Caption         =   "Kaydet"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Sil"
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Güncelle"
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton yenikayit 
         Caption         =   "Yeni Kayýt"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton sinif 
      Caption         =   "Sýnýf"
      Height          =   375
      Left            =   240
      TabIndex        =   50
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ö"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   49
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A"
      Height          =   375
      Index           =   30
      Left            =   1320
      TabIndex        =   48
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B"
      Height          =   375
      Index           =   29
      Left            =   1560
      TabIndex        =   47
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   375
      Index           =   28
      Left            =   1800
      TabIndex        =   46
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ç"
      Height          =   375
      Index           =   27
      Left            =   2040
      TabIndex        =   45
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      Height          =   375
      Index           =   26
      Left            =   2280
      TabIndex        =   44
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E"
      Height          =   375
      Index           =   25
      Left            =   2520
      TabIndex        =   43
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F"
      Height          =   375
      Index           =   24
      Left            =   2760
      TabIndex        =   42
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G"
      Height          =   375
      Index           =   23
      Left            =   3000
      TabIndex        =   41
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H"
      Height          =   375
      Index           =   22
      Left            =   3240
      TabIndex        =   40
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I"
      Height          =   375
      Index           =   21
      Left            =   3480
      TabIndex        =   39
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ý"
      Height          =   375
      Index           =   20
      Left            =   3720
      TabIndex        =   38
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "J"
      Height          =   375
      Index           =   19
      Left            =   3960
      TabIndex        =   37
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "K"
      Height          =   375
      Index           =   18
      Left            =   4200
      TabIndex        =   36
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "L"
      Height          =   375
      Index           =   17
      Left            =   4440
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M"
      Height          =   375
      Index           =   16
      Left            =   4680
      TabIndex        =   34
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      Height          =   375
      Index           =   15
      Left            =   4920
      TabIndex        =   33
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O"
      Height          =   375
      Index           =   14
      Left            =   5160
      TabIndex        =   32
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P"
      Height          =   375
      Index           =   13
      Left            =   5640
      TabIndex        =   31
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      Height          =   375
      Index           =   12
      Left            =   5880
      TabIndex        =   30
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S"
      Height          =   375
      Index           =   11
      Left            =   6120
      TabIndex        =   29
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Þ"
      Height          =   375
      Index           =   10
      Left            =   6360
      TabIndex        =   28
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "T"
      Height          =   375
      Index           =   9
      Left            =   6600
      TabIndex        =   27
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "U"
      Height          =   375
      Index           =   8
      Left            =   6840
      TabIndex        =   26
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ü"
      Height          =   375
      Index           =   7
      Left            =   7080
      TabIndex        =   25
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "V"
      Height          =   375
      Index           =   6
      Left            =   7320
      TabIndex        =   24
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Y"
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   23
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Z"
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   22
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "*"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   10335
      Begin VB.CommandButton Command2 
         Caption         =   "Numarayý Deðiþtirin"
         Height          =   375
         Left            =   2640
         TabIndex        =   59
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Resmi Kameradan Al"
         Height          =   375
         Left            =   5520
         TabIndex        =   53
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Öðrenci Ýçin Resim Yükle"
         Height          =   375
         Left            =   5520
         TabIndex        =   52
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox dali 
         BackColor       =   &H00FFFFFF&
         DataField       =   "dali"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Text            =   "Erken Çocukluk Eðitimi"
         Top             =   1800
         Width           =   2295
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
      Begin VB.TextBox numarasi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrno"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox adi 
         DataField       =   "adi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox soyadi 
         DataField       =   "soyadi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox bolumu 
         DataField       =   "bolumu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Text            =   "Çocuk Geliþimi ve Eðitimi"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cins."
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
         Left            =   3960
         TabIndex        =   13
         Top             =   1320
         Width           =   735
         Begin VB.ComboBox cinsiyet 
            DataField       =   "cinsiyet"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "ogrenci.frx":0000
            Left            =   0
            List            =   "ogrenci.frx":000A
            TabIndex        =   4
            Text            =   "KIZ"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.ComboBox sinifi 
         DataField       =   "sinifi"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Text            =   "sinifi"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label91 
         Caption         =   "Label91"
         DataField       =   "ogrresim"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   8280
         TabIndex        =   60
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image vesikalik 
         BorderStyle     =   1  'Fixed Single
         Height          =   1815
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1695
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
         TabIndex        =   20
         Top             =   1800
         Width           =   855
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
         TabIndex        =   19
         Top             =   1440
         Width           =   2175
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
         TabIndex        =   18
         Top             =   240
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
         TabIndex        =   16
         Top             =   960
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   62
      Top             =   2880
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Öðrenci"
      TabPicture(0)   =   "ogrenci.frx":001A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label17"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label20"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label21"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label22"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label23"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label24"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label25"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label26"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ogrevadresi"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "kartno"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "yab1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "kulub"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "davranis"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "yab2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "email"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "staj"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "yapilan"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "kalan"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text3"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text4"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text5"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text6"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text7"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text8"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text9"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text10"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text11"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text12"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text13"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text14"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Combo1"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Combo2"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Combo3"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).ControlCount=   47
      TabCaption(1)   =   "Veli"
      TabPicture(1)   =   "ogrenci.frx":0036
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label42"
      Tab(1).Control(1)=   "Label41"
      Tab(1).Control(2)=   "Label27"
      Tab(1).Control(3)=   "Label28"
      Tab(1).Control(4)=   "Label29"
      Tab(1).Control(5)=   "Label30"
      Tab(1).Control(6)=   "Label31"
      Tab(1).Control(7)=   "Label32"
      Tab(1).Control(8)=   "Label33"
      Tab(1).Control(9)=   "Label34"
      Tab(1).Control(10)=   "Label35"
      Tab(1).Control(11)=   "Label36"
      Tab(1).Control(12)=   "Label37"
      Tab(1).Control(13)=   "Label38"
      Tab(1).Control(14)=   "Label39"
      Tab(1).Control(15)=   "Label40"
      Tab(1).Control(16)=   "velievadresi"
      Tab(1).Control(17)=   "veliisadresi"
      Tab(1).Control(18)=   "veligelirduzeyi"
      Tab(1).Control(19)=   "veliegitimdurumu"
      Tab(1).Control(20)=   "veliadi"
      Tab(1).Control(21)=   "velituru"
      Tab(1).Control(22)=   "veliispk"
      Tab(1).Control(23)=   "veliistel"
      Tab(1).Control(24)=   "veliemail"
      Tab(1).Control(25)=   "velitcno"
      Tab(1).Control(26)=   "velisoyadi"
      Tab(1).Control(27)=   "velimeslegi"
      Tab(1).Control(28)=   "velievpk"
      Tab(1).Control(29)=   "velievtel"
      Tab(1).Control(30)=   "velicep"
      Tab(1).Control(31)=   "velifaks"
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "Kayýt"
      TabPicture(2)   =   "ogrenci.frx":0052
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "kaymezunokul"
      Tab(2).Control(1)=   "kaykaytarihi"
      Tab(2).Control(2)=   "kaysinavdrumu"
      Tab(2).Control(3)=   "kaygirispuani"
      Tab(2).Control(4)=   "kaybursorani"
      Tab(2).Control(5)=   "kaysehirokul"
      Tab(2).Control(6)=   "kaygunduzlu"
      Tab(2).Control(7)=   "kayburslu"
      Tab(2).Control(8)=   "kayservis"
      Tab(2).Control(9)=   "kayyatili"
      Tab(2).Control(10)=   "kayyemek"
      Tab(2).Control(11)=   "Label43"
      Tab(2).Control(12)=   "Label44"
      Tab(2).Control(13)=   "Label45"
      Tab(2).Control(14)=   "Label46"
      Tab(2).Control(15)=   "Label47"
      Tab(2).Control(16)=   "Label48"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Diploma"
      TabPicture(3)   =   "ogrenci.frx":006E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Nüfus"
      TabPicture(4)   =   "ogrenci.frx":008A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Düþünce"
      TabPicture(5)   =   "ogrenci.frx":00A6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame9"
      Tab(5).ControlCount=   1
      Begin VB.TextBox velifaks 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velifaks"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   208
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox velicep 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliceptel"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   207
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox velievtel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velievtel"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   206
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox velievpk 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velievadresipk"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   205
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox velimeslegi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velimeslegi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   204
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox velisoyadi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velisoyadi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   203
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox velitcno 
         BackColor       =   &H00FFFFFF&
         DataField       =   "velitckimlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   202
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox veliemail 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliemail"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   201
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox veliistel 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliistel"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   200
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox veliispk 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliisadresipk"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   199
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox velituru 
         DataField       =   "velituru"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ogrenci.frx":00C2
         Left            =   -72360
         List            =   "ogrenci.frx":00CF
         TabIndex        =   198
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
         TabIndex        =   197
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "ogralandegistirdi"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ogrenci.frx":00E6
         Left            =   3120
         List            =   "ogrenci.frx":00F0
         TabIndex        =   196
         Text            =   "HAYIR"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "ogrnotmatik"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ogrenci.frx":0101
         Left            =   3120
         List            =   "ogrenci.frx":010B
         TabIndex        =   195
         Text            =   "HAYIR"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "ogryetiskurs"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ogrenci.frx":011C
         Left            =   3120
         List            =   "ogrenci.frx":0126
         TabIndex        =   194
         Text            =   "HAYIR"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogruygulamayapilan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8880
         TabIndex        =   193
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogruygulamakalan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   9480
         TabIndex        =   192
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrkalanetkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   191
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogruygulamaetkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8880
         TabIndex        =   190
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrokudugukitap"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8880
         TabIndex        =   189
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogryapilanetkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8880
         TabIndex        =   188
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrtoplametkinlik"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8880
         TabIndex        =   187
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrceptelefon"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8040
         TabIndex        =   186
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrtelefon"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8040
         TabIndex        =   185
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrsehir"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8040
         TabIndex        =   184
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrsemt"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8040
         TabIndex        =   183
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrpostakodu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   8040
         TabIndex        =   182
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox kalan 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrstajkalan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   181
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox yapilan 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrstajyapilan"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   3960
         TabIndex        =   180
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox staj 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrstajsuresi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   2400
         TabIndex        =   179
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox email 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogremail"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   3000
         TabIndex        =   178
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox yab2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogryabanci2"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   4440
         TabIndex        =   177
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox davranis 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrdavranispuani"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   3000
         TabIndex        =   176
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox kulub 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrkulub"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   3000
         TabIndex        =   175
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox yab1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogryabanci1"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   3000
         TabIndex        =   174
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox kartno 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ogrbankano"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   3000
         TabIndex        =   173
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox veliegitimdurumu 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veliegitimdurumu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72360
         TabIndex        =   172
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox veligelirduzeyi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "veligelirduzeyi"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67200
         TabIndex        =   171
         Top             =   3600
         Width           =   2415
      End
      Begin VB.ComboBox kaymezunokul 
         DataField       =   "kaymezunokul"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ogrenci.frx":0137
         Left            =   -72240
         List            =   "ogrenci.frx":0139
         Sorted          =   -1  'True
         TabIndex        =   170
         Text            =   "Combo5"
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox kaykaytarihi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaykaytarih"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   169
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox kaysinavdrumu 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaysinavdurumu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   168
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox kaygirispuani 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaygirispuani"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   167
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox kaybursorani 
         BackColor       =   &H00FFFFFF&
         DataField       =   "kaybursorani"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72240
         TabIndex        =   166
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CheckBox kaysehirokul 
         Caption         =   "Þehir Okulu"
         DataField       =   "kaysehirokul"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -72120
         TabIndex        =   165
         Top             =   840
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox kaygunduzlu 
         Caption         =   "Gündüzlü"
         DataField       =   "kaygunduzlu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -68520
         TabIndex        =   164
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox kayburslu 
         Caption         =   "Burslu"
         DataField       =   "kayburslu"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -69960
         TabIndex        =   163
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox kayservis 
         Caption         =   "Servis"
         DataField       =   "kayservis"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -69360
         TabIndex        =   162
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox kayyatili 
         Caption         =   "Yatýlý"
         DataField       =   "kayyatili"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -66840
         TabIndex        =   161
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox kayyemek 
         Caption         =   "Yemek"
         DataField       =   "kayyemek"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   -67320
         TabIndex        =   160
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
         TabIndex        =   144
         Top             =   360
         Width           =   9975
         Begin VB.ComboBox dipokuladi 
            DataField       =   "dipokuladi"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "ogrenci.frx":013B
            Left            =   2640
            List            =   "ogrenci.frx":013D
            Sorted          =   -1  'True
            TabIndex        =   152
            Text            =   "Combo5"
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox dipbelgecinsi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipbelgecinsi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   151
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox diptarih 
            BackColor       =   &H00FFFFFF&
            DataField       =   "diptarih"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   150
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Özlük Dosyasý Geldi"
            DataField       =   "dipozlukdosyasi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   240
            TabIndex        =   149
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox dipsayisi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipsayisi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   148
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox dipgelokultur 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgelokultur"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   147
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox dipgelokulno 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgelokulno"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   146
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox dipdipnotu 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipdipnotu"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   145
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
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
            TabIndex        =   156
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            TabIndex        =   153
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
         TabIndex        =   112
         Top             =   2520
         Width           =   9975
         Begin VB.TextBox dipgitdipnotu 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitdipnotu"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   136
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox dipgitokulno 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitokulno"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   135
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox dipgitdippuani 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitdippuani"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   134
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox dipgitsayi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitsayi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   3840
            TabIndex        =   133
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox dipgittarih 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgittarih"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   132
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox dipgitbelgecinsi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "dipgitbelgecinsi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   131
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox dipgitokuladi 
            DataField       =   "dipgitokuladi"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "ogrenci.frx":013F
            Left            =   2640
            List            =   "ogrenci.frx":0141
            Sorted          =   -1  'True
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   120
            Top             =   1680
            Width           =   6975
            Begin VB.TextBox dipkunyegonbelge 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyegonbelge"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   120
               TabIndex        =   124
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox dipkunyeistarihno 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyeistarihno"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1680
               TabIndex        =   123
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox dipkunyegontarih 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyegontarih"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   3360
               TabIndex        =   122
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox dipkunyealtarihno 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipkunyealtarihno"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   5040
               TabIndex        =   121
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
               TabIndex        =   128
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
               TabIndex        =   127
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
               TabIndex        =   126
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
               TabIndex        =   125
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
            TabIndex        =   113
            Top             =   1680
            Width           =   2655
            Begin VB.TextBox dipisbelgeserino 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipisbelgeserino"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1440
               TabIndex        =   116
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox dipisbelgetarih 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipisbelgetarih"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1440
               TabIndex        =   115
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox dipisbelgeno 
               BackColor       =   &H00FFFFFF&
               DataField       =   "dipisbelgeno"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   1440
               TabIndex        =   114
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
               TabIndex        =   119
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
               TabIndex        =   118
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
               TabIndex        =   117
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
            TabIndex        =   143
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
            TabIndex        =   142
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   139
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
            TabIndex        =   138
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
            TabIndex        =   137
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   69
         Top             =   480
         Width           =   10095
         Begin VB.TextBox nufveryer 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufveryer"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   90
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox nufsira 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufsira"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   89
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox nufaile 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufaile"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   88
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox nufcilt 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufcilt"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   87
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox nufmah 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufmah"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   86
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox nufuyruk 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufuyruk"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   85
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox nufilce 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufilce"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   84
            Top             =   3480
            Width           =   2415
         End
         Begin VB.TextBox nufil 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufil"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   83
            Top             =   3120
            Width           =   2415
         End
         Begin VB.ComboBox nufkan 
            DataField       =   "nufkan"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "ogrenci.frx":0143
            Left            =   3840
            List            =   "ogrenci.frx":015F
            Sorted          =   -1  'True
            TabIndex        =   82
            Text            =   "Combo5"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.ComboBox nufdini 
            DataField       =   "nufdini"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "ogrenci.frx":01A5
            Left            =   1320
            List            =   "ogrenci.frx":01AF
            Sorted          =   -1  'True
            TabIndex        =   81
            Text            =   "Combo5"
            Top             =   2760
            Width           =   1335
         End
         Begin VB.ComboBox nufmedeni 
            DataField       =   "nufmedeni"
            DataSource      =   "Data1"
            Height          =   315
            ItemData        =   "ogrenci.frx":01C1
            Left            =   2640
            List            =   "ogrenci.frx":01CE
            Sorted          =   -1  'True
            TabIndex        =   80
            Text            =   "Combo5"
            Top             =   2400
            Width           =   2415
         End
         Begin VB.TextBox nufdogumtarihi 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufdogumtarihi"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   79
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox nufdogumyeri 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufdogumyeri"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   78
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox nufana 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufana"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   77
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox nufbaba 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufbaba"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   76
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox nufcuzdanserino 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufcuzdanserino"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   75
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox nufverneden 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufverneden"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   74
            Top             =   2400
            Width           =   2295
         End
         Begin VB.TextBox nufkayitno 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufkayitno"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   73
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox nufvertarih 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufvertarih"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   72
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox nufaskerlik 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nufaskerlik"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   7680
            TabIndex        =   71
            Top             =   3480
            Width           =   2295
         End
         Begin VB.TextBox nuftckimlik 
            BackColor       =   &H00FFFFFF&
            DataField       =   "nuftckimlik"
            DataSource      =   "Data1"
            Height          =   375
            Left            =   2640
            TabIndex        =   70
            Top             =   240
            Width           =   2415
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
            TabIndex        =   111
            Top             =   2040
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
            TabIndex        =   110
            Top             =   1680
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
            TabIndex        =   109
            Top             =   1320
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
            TabIndex        =   108
            Top             =   960
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
            TabIndex        =   107
            Top             =   600
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
            Left            =   5280
            TabIndex        =   106
            Top             =   240
            Width           =   2415
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
            TabIndex        =   105
            Top             =   3480
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
            TabIndex        =   104
            Top             =   3120
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
            TabIndex        =   103
            Top             =   2760
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
            TabIndex        =   102
            Top             =   2760
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
            TabIndex        =   101
            Top             =   2400
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
            TabIndex        =   100
            Top             =   2040
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
            TabIndex        =   99
            Top             =   1680
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
            TabIndex        =   98
            Top             =   1320
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
            TabIndex        =   97
            Top             =   960
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
            TabIndex        =   96
            Top             =   600
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
            TabIndex        =   95
            Top             =   2400
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
            TabIndex        =   94
            Top             =   2760
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
            TabIndex        =   93
            Top             =   3120
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
            TabIndex        =   92
            Top             =   3480
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
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   66
         Top             =   480
         Width           =   10095
         Begin VB.TextBox dusunce 
            DataField       =   "dusunce"
            DataSource      =   "Data1"
            Height          =   3735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   67
            Top             =   720
            Width           =   9735
         End
         Begin VB.Label Label90 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hakkýnda Düþünceler"
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
            TabIndex        =   68
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.TextBox veliisadresi 
         DataField       =   "veliisadresi"
         DataSource      =   "Data1"
         Height          =   975
         Left            =   -72360
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox velievadresi 
         DataField       =   "velievadresi"
         DataSource      =   "Data1"
         Height          =   975
         Left            =   -67200
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox ogrevadresi 
         DataField       =   "ogrevadresi"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   8040
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   480
         Width           =   2055
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
         TabIndex        =   252
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
         TabIndex        =   251
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
         TabIndex        =   250
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
         TabIndex        =   249
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
         TabIndex        =   248
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
         TabIndex        =   247
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
         TabIndex        =   246
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
         TabIndex        =   245
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
         TabIndex        =   244
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
         TabIndex        =   243
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
         TabIndex        =   242
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
         TabIndex        =   241
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
         TabIndex        =   240
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
         TabIndex        =   239
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Uygulamalarý Ýzleme Etkinlik Yap/Kal"
         Height          =   375
         Left            =   6240
         TabIndex        =   238
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Uygulamalarý Ýzleme Etkinlik Süresi"
         Height          =   375
         Left            =   6240
         TabIndex        =   237
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Okuduðu Kitap Sayýsý"
         Height          =   375
         Left            =   6240
         TabIndex        =   236
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yapýlan/Kalan Etk."
         Height          =   375
         Left            =   6240
         TabIndex        =   235
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Top. Sos. Etk. Saati"
         Height          =   375
         Left            =   6240
         TabIndex        =   234
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
         Left            =   6240
         TabIndex        =   233
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
         Left            =   6240
         TabIndex        =   232
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
         Left            =   6240
         TabIndex        =   231
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
         Left            =   6240
         TabIndex        =   230
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
         Left            =   6240
         TabIndex        =   229
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
         Height          =   735
         Left            =   6240
         TabIndex        =   228
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
         Left            =   120
         TabIndex        =   227
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
         Left            =   120
         TabIndex        =   226
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
         Left            =   120
         TabIndex        =   225
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
         Left            =   4440
         TabIndex        =   224
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
         Left            =   3000
         TabIndex        =   223
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
         Left            =   120
         TabIndex        =   222
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
         Left            =   120
         TabIndex        =   221
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
         Left            =   120
         TabIndex        =   220
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
         Left            =   120
         TabIndex        =   219
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
         Left            =   120
         TabIndex        =   218
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
         Left            =   120
         TabIndex        =   217
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
         TabIndex        =   216
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
         TabIndex        =   215
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
         TabIndex        =   214
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
         TabIndex        =   213
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
         TabIndex        =   212
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
         TabIndex        =   211
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
         TabIndex        =   210
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
         TabIndex        =   209
         Top             =   2280
         Width           =   2535
      End
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   3960
      Top             =   8280
      Width           =   4095
   End
End
Attribute VB_Name = "ogrenci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bul_Click()
bul.Text = Empty
End Sub

Private Sub Command1_Click(Index As Integer)
harf = Command1(Index).Caption
indis.Show
End Sub

Private Sub Command2_Click()
sorgula.Show
End Sub

Private Sub Command6_Click()
If numarasi = Empty Or numarasi = " " Then
    MsgBox "Lütfen resim yüklemek istediðiniz öðrenciyi seçiniz", vbExclamation, "Uyarý"
    Exit Sub
Else
    Similasyon.Show
End If
End Sub



Private Sub Command7_Click()
varmiyokmu.Show
End Sub

Private Sub Command8_Click()
Data1.Recordset.FindFirst "ogrno='" & bul.Text & "'"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Çýkmak üzerisiniz. Tüm deðiþiklikler kaydedilecek", vbExclamation, "Uyarý"
End Sub

Private Sub Label91_Change()
On Local Error GoTo hata
' KIZ MESLEK LOGOSU GELECEK UNUTMAAAAAAAA
vesikalik.Picture = LoadPicture(App.Path & "\Ogrenci_Resimleri\" & numarasi & ".jpg")
hata:
If Err = 53 Then
    vesikalik.Picture = LoadPicture(App.Path & "\Ogrenci_Resimleri\resimyok.jpg")
End If
End Sub


Private Sub staj_Change()
kalan = Val(staj) - Val(yapilan)
End Sub

Private Sub Text11_Change()
Text13 = Val(Text11) - Val(Text14)
End Sub

Private Sub Text14_Change()
Text13 = Val(Text11) - Val(Text14)
End Sub

Private Sub Text8_Change()
Text12 = Val(Text8) - Val(Text9)
End Sub

Private Sub Text9_Change()
Text12 = Val(Text8) - Val(Text9)
End Sub

Private Sub yapilan_Change()
kalan = Val(staj) - Val(yapilan)
End Sub

Private Sub yenikayit_Click()
On Local Error GoTo hata
'yeniogrenci.Show
Call comboyukle
Data1.Recordset.AddNew
yenikayit.Enabled = False
kaydet.Enabled = True
Command3.Enabled = False
hata:
If Err = 3426 Then
MsgBox "Kayýtlý bilginin üstüne bilgi eklediniz. Kayýtýn geçerli olmasý için güncelle demeniz lazým", vbExclamation, "Uyarý"
End If
End Sub


Private Sub kaydet_Click()
On Local Error GoTo hata
Call veri_ac(False, False)
Call tablo_ac("select * from ogrenci where ogrno='" & numarasi.Text & "' or  nuftckimlik='" & nuftckimlik.Text & "'")
If tablo.RecordCount >= 1 Then
MsgBox "Bu numarayla yada Tc Kimlik numarasý ile önceden kayýt yapýlmýþtýr. Lütfen okul numarasýný yada TC kimlik numarasýný kontrol edip tekrar deneyiniz", vbCritical, "Bu Numara yada TC Numarasý var"
Exit Sub
End If
veri.Close

If numarasi.Text = Empty Or adi.Text = Empty Or soyadi.Text = Empty Or nuftckimlik.Text = Empty Then
        MsgBox "Tc Kimlik Numarasý, Okul Numarasý, Adý yada Soyadý kýsmýný boþ býrakamazsýnýz!", vbCritical, "Hata"
        Exit Sub
    End If
Data1.Recordset.AddNew
yenikayit.Enabled = True
kaydet.Enabled = False
MsgBox "Öðrenci Kayýt Edilmiþtir", vbInformation, "KAYIT"
Data1.Recordset.MoveLast
hata:
If Err = 3426 Then
    MsgBox "Bu numaraya ait öðrenci bulunmaktadýr. Lütfen yeni bir numara belirleyiniz", vbCritical, "HATA"
    numarasi.Text = Empty
    Exit Sub
    Else
    If Err <> 0 Then MsgBox Err & " " & Error
End If

End Sub

Private Sub Command3_Click()
    If numarasi.Text = Empty Or adi.Text = Empty Or soyadi.Text = Empty Then
        MsgBox "Güncellenecek bilgi yok. Lütfen önce bir öðrenci seçiniz!", vbCritical, "Hata"
        Exit Sub
    End If
Data1.UpdateRecord
Data1.Recordset.Bookmark = Data1.Recordset.LastModified
yenikayit.Enabled = True
kaydet.Enabled = False
MsgBox "Güncelleþtirme gerçekleþmiþtir", vbInformation, "Bilgi"
End Sub

Private Sub Command4_Click()
Dim silme
On Local Error Resume Next
    If numarasi.Text = Empty Then
        MsgBox "Önce silinecek kaydý seçmeniz lazým!", vbCritical, "Hata"
        Exit Sub
    End If
    
silme = MsgBox("Kaydý gerçekten silmek istediðinize eminmisini?", vbYesNo, "Kayýt Silme")

    If silme = 6 Then
        With Data1.Recordset
            .Delete
            .MoveNext
                If .EOF Then .MoveLast
        End With
        MsgBox "Kayýt Silinmiþtir", vbInformation, "Kayýt Silme"
    Else
        MsgBox "Silme iþlemini iptal ettiniz", vbInformation, "Kayýt Silme"
       
    End If
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
Private Sub ilk_Click()
On Local Error Resume Next
Data1.Recordset.MoveFirst
End Sub

Private Sub onceki_Click()
On Local Error Resume Next
    If Data1.Recordset.BOF = True Then
        MsgBox "Ýlk kayýttasýnýz", vbInformation, "Bilgi"
    Else
        Data1.Recordset.MovePrevious
    End If

End Sub

Private Sub son_Click()
On Local Error Resume Next
Data1.Recordset.MoveLast
End Sub

Private Sub sonraki_Click()
On Local Error Resume Next

    If Data1.Recordset.EOF = True Then
        MsgBox "Son kayýttasýnýz", vbInformation, "Bilgi"
    Else
        Data1.Recordset.MoveNext
    End If
End Sub

