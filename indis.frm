VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form indis 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   2385
   ClientTop       =   810
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7740
   Begin MSFlexGridLib.MSFlexGrid liste 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "(*) Aktarmak istediðiniz kaydý çift týklayýnýz."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   7575
   End
End
Attribute VB_Name = "indis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim X As Integer
If harf = "*" Then
    Me.Caption = "Bütün öðrenci listesi"
Else
    Me.Caption = harf & " ile baþlayanlarýn listesi"
End If
X = 0
Call veri_ac(False, False)
Call tablo_ac("Select * from ogrenci where adi like '" & harf & "*' order by adi")
liste.TextMatrix(0, 0) = "NO"
liste.TextMatrix(0, 1) = "ADI"
liste.TextMatrix(0, 2) = "SOYADI"
liste.TextMatrix(0, 3) = "SINIFI"

Do While Not tablo.EOF
X = X + 1
liste.AddItem ""
liste.TextMatrix(X, 0) = tablo("ogrno")
liste.TextMatrix(X, 1) = tablo("adi")
liste.TextMatrix(X, 2) = tablo("soyadi")
liste.TextMatrix(X, 3) = tablo("sinifi")

tablo.MoveNext
Loop
End Sub



Private Sub liste_DblClick()
Dim aktar As Integer
If liste.TextMatrix(liste.Row, 0) = Empty Or liste.TextMatrix(liste.Row, 0) = "NO" Then
    MsgBox "Seçtiðiniz hücrede hiç bir öðrenci bilgisi bulunmamaktadir!", vbExclamation, "Uyarý"
    Exit Sub
Else
    ogrenci.Data1.Recordset.FindFirst "ogrno='" & liste.TextMatrix(liste.Row, 0) & "'"
    ogrenci.Show
    ogrenci.yenikayit.Enabled = True
    ogrenci.kaydet.Enabled = False
    ogrenci.Command3.Enabled = True
    Unload Me
    End If
End Sub
