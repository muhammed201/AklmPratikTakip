VERSION 5.00
Begin VB.Form sinifduzenle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S�n�f D�zenleme"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "S�n�f D�zenleme "
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6495
      Begin VB.TextBox sismi 
         Height          =   405
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox skontenjan 
         Height          =   405
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox sogretmeni 
         Height          =   405
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "S�n�f �smi"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1455
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
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "S�n�f ��retmenin Ad�-Soyad�"
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
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kay�tlar� G�ncelle"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ptal Et(Ana Men�ye Geri d�n)"
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
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
   End
End
Attribute VB_Name = "sinifduzenle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If sismi.Text = Empty Or skontenjan.Text = Empty Then
MsgBox "S�n�f ismini yada Kontenjan�n� bo� b�rakamazs�n�z!", vbCritical, "HATA"
Exit Sub
End If
Call veri_ac(False, False)
Call tablo_ac("Select * from siniflar where id=" & siniflar.sinifliste.TextMatrix(secilisatir, 0))
tablo.Edit
tablo("sinif") = sismi.Text
tablo("kapasite") = Val(skontenjan.Text)
tablo("sinifogretmeni") = sogretmeni.Text
tablo.Update
MsgBox "Bilgiler g�ncellenmi�tir ", vbInformation, "G�ncelleme ��lemi"
Call siniflar.yukle
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'MsgBox Siniflar.sinifliste.TextMatrix(secilisatir, 0)
If siniflar.sinifliste.TextMatrix(secilisatir, 0) = "ID" Then
MsgBox "L�tfen bir kay�t se�iniz ve daha sonra deneyiniz. Kay�t se�mek i�in d�zeltme yapmak istedi�iniz kayd� t�klay�n�z."
Unload Me
Exit Sub
End If
Call veri_ac(False, False)
Call tablo_ac("Select * from siniflar where id=" & siniflar.sinifliste.TextMatrix(secilisatir, 0))
sismi.Text = tablo("sinif")
skontenjan.Text = tablo("kapasite")
sogretmeni.Text = tablo("sinifogretmeni")
tablo.Close
veri.Close
End Sub
