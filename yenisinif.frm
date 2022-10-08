VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form yenisinif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yeni Sýnýf Kaydý"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin MSACAL.Calendar bittarih 
      Height          =   2295
      Left            =   3120
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   4695
      _Version        =   524288
      _ExtentX        =   8281
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2008
      Month           =   10
      Day             =   20
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Tur"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Tur"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Tur"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar bastarih 
      Height          =   2415
      Left            =   1320
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
      _Version        =   524288
      _ExtentX        =   8916
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2008
      Month           =   10
      Day             =   20
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Tur"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Tur"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Tur"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ýptal Et(Ana Menüye Geri dön)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yeni Sýnýfý Aç"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   1560
      Width           =   7095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Yeni Sýnýf Kaydý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11055
      Begin VB.TextBox toplamkursaati 
         Height          =   405
         Left            =   9120
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox kurbittarihi 
         Height          =   405
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox kurbastarihi 
         Height          =   405
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox sogretmeni 
         Height          =   405
         Left            =   3720
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox skontenjan 
         Height          =   405
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox sismi 
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Toplam Kurs Süresi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Kurs Bit. Tar."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Kurs Baþ. Tar."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Sýnýf Öðretmenin Adý-Soyadý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Kontenjan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Sýnýf Ýsmi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "yenisinif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bastarih_Click()
kurbastarihi.Text = bastarih.Value
End Sub

Private Sub bittarih_Click()
kurbittarihi.Text = bittarih.Value
End Sub

Private Sub Command1_Click()
Call veri_ac(False, False)
Call tablo_ac("Select * from siniflar")
tablo.AddNew
tablo("sinif") = sismi.Text
tablo("kapasite") = Val(skontenjan.Text)
tablo("sinifogretmeni") = sogretmeni.Text
tablo("kursbastarihi") = kurbastarihi.Text
tablo("kursbittarihi") = kurbittarihi.Text
tablo("toplamkurssaat") = toplamkursaati.Text
tablo.Update
MsgBox "Kayýt Eklenmiþtir", vbInformation, "Kayýt Ýþlemleri"
Unload Me
Call Siniflar.yukle
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
bastarih.Value = Now
bittarih.Value = Now
End Sub

Private Sub kurbastarihi_GotFocus()
bastarih.Visible = True
End Sub

Private Sub kurbastarihi_LostFocus()
bastarih.Visible = False
End Sub

Private Sub kurbittarihi_GotFocus()
bittarih.Visible = True
End Sub

Private Sub kurbittarihi_LostFocus()
bittarih.Visible = False
End Sub
