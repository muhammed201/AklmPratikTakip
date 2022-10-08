VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form sorgula 
   Caption         =   "Numara Sorgula"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid liste 
      Height          =   3855
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6800
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sorgulanacak Aralýkdaki Sayýlar"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "Sorgula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox bitis 
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox baslangic 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label yukleniyor 
         Caption         =   "Taranýyor lütfen bekleyiniz....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "arasý"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bitiþ"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Baþlangýç"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "sorgula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X, Y As Integer
yukleniyor.Visible = True
MsgBox baslangic & " ile " & bitis & " arasýnda tarama yapýlacaktýr", vbOKOnly, "Onay"
liste.Cols = 2
liste.Rows = 1
Call veri_ac(False, False)
Y = 0


liste.TextMatrix(0, 0) = "NUMARA"
liste.TextMatrix(0, 1) = "DURUMU"
For X = baslangic To bitis
    Y = Y + 1
    Call tablo_ac("Select * from ogrenci where ogrno='" & X & "'")
        If tablo.RecordCount = 0 Then
            liste.AddItem ""
            liste.TextMatrix(Y, 0) = X
            liste.TextMatrix(Y, 1) = "(BOÞ)"
        Else
            liste.AddItem ""
            liste.TextMatrix(Y, 0) = X
            liste.TextMatrix(Y, 1) = "KULLANIMDA"

        End If
    tablo.Close
 Next
 yukleniyor.Visible = False
Call boyutlandir
End Sub



Private Sub liste_DblClick()
Dim aktar As Integer
If liste.TextMatrix(liste.Row, 1) = "KULLANIMDA" Then
    MsgBox "Bu numara baþka öðrenci için kullanýlýyor. Lütfen baþka bir numara seçiniz!", vbExclamation, "Uyarý"
    Exit Sub
ElseIf liste.TextMatrix(liste.Row, 1) = Empty Then
    MsgBox "Lütfen arama yaptýrýn", vbExclamation, "Uyarý"
    Exit Sub
Else
    aktar = liste.TextMatrix(liste.Row, 0)
    Unload Me
    ogrenci.Show
    ogrenci.numarasi.Text = aktar
End If
End Sub

Sub boyutlandir()
Me.Height = 7320
End Sub
