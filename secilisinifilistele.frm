VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form secilisinifilistele 
   Caption         =   "Sýnýf Listesi"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11145
   FillColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "EXCEL'E AKTAR"
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Öðren Soyadýna Göre Sýrala"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Öðrencinin Adýna Göre Sýrala"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid sinifliste 
      Bindings        =   "secilisinifilistele.frx":0000
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   8404992
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
   Begin VB.Label bilgi 
      AutoSize        =   -1  'True
      Caption         =   "Genel Bilgi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   1155
   End
End
Attribute VB_Name = "secilisinifilistele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public listele As String

Private Sub Command1_Click()
Dim don As Integer
Dim Ex As Excel.Application
Set Ex = New Excel.Application
   Ex.Visible = True   'Görünmesini istemiyorsak False
   'Ex.Workbooks.Add 'Yeni bir çalýþma sayfasý oluþturmak için kullanýrýz.
   'Ex.Workbooks.Open ("D:\Muhammed\ReSearch\nakit1.xls") 'Bunu da eðer hazýr bir excel sayfamýz var da onun üzerinde çalýþmak istiyorsak Workbooks.Add yerine kullanabiliriz.
   Ex.Workbooks.Open (App.Path & "\aksamci.xls")
'HANGÝ SINIF OLDUÐUNUN BAÞLIÐI
Ex.Sheets("LISTE").Range("A1").Value = listele & " GRUBU"
    For don = 0 To 47
        Ex.Sheets("LISTE").Range("A" & 3 + don).Value = Empty
        Ex.Sheets("LISTE").Range("B" & 3 + don).Value = Empty
    Next
  For don = 0 To sinifliste.Rows - 1
        Ex.Sheets("LISTE").Range("A" & 2 + don).Value = sinifliste.TextMatrix(don, 0)
        Ex.Sheets("LISTE").Range("B" & 2 + don).Value = sinifliste.TextMatrix(don, 1) & " " & sinifliste.TextMatrix(don, 2)
  Next


    'Ex.Sheets("Poliklinik").PrintOut 'Yazdýr

End Sub

Private Sub Command2_Click()
Call snf_yukle("Order by adi")
End Sub

Private Sub Command3_Click()
Call snf_yukle("Order by soyadi")
End Sub


Private Sub Form_Load()
listele = Siniflar.sinifliste.TextMatrix(secilisatir, 1)
Me.Caption = listele & " Þubesi Sýnýf Listesi"
Call snf_yukle("Order by adi")
End Sub


Sub snf_yukle(listeleme_bicimi As String)
Dim X As Integer
Call veri_ac(False, False)
Call tablo_ac("Select * from ogrenci where sinifi='" & listele & "' " & listeleme_bicimi)
sinifliste.Clear

sinifliste.Cols = 3
sinifliste.Rows = 1

sinifliste.TextMatrix(0, 0) = "S.N"
sinifliste.TextMatrix(0, 1) = "ADI"
sinifliste.TextMatrix(0, 2) = "SOYADI"

sinifliste.ColWidth(0) = 500
sinifliste.ColWidth(1) = 3000
sinifliste.ColWidth(2) = 3000

X = 0
Do While Not tablo.EOF
X = X + 1
sinifliste.AddItem ""
sinifliste.TextMatrix(X, 0) = X
sinifliste.TextMatrix(X, 1) = tablo("adi")
sinifliste.TextMatrix(X, 2) = tablo("soyadi")
tablo.MoveNext
Loop
tablo.Close
veri.Close
bilgi.Caption = listele & " Þubesinden toplam  " & X & " adet öðrenci listelenmiþtir"
End Sub
