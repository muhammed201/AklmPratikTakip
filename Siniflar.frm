VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Siniflar 
   Caption         =   "S�n�flar"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   13695
      Begin VB.CommandButton Command2 
         Caption         =   "Ana Men�ye D�n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   13455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kay�tl� S�n�f Listeleri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin MSFlexGridLib.MSFlexGrid sinifliste 
         Bindings        =   "Siniflar.frx":0000
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   4
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
      Begin VB.Label Label1 
         Caption         =   "��lem yapmak istedi�iniz S�n�f� t�klayarak se�iniz ve sa� t�klayarak men�den i�lem se�iniz l�tfen..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   12135
      End
   End
   Begin VB.Menu jmenu 
      Caption         =   "Men�"
      Visible         =   0   'False
      Begin VB.Menu Ekle 
         Caption         =   "Yeni S�n�f Ekle"
      End
      Begin VB.Menu Duzenle 
         Caption         =   "Se�ili S�n�f� D�zenle"
      End
      Begin VB.Menu Sil 
         Caption         =   "Se�ili S�n�f� Sil"
      End
      Begin VB.Menu sinifyazdir 
         Caption         =   "S�n�f Listesini Yaz�ya D�k"
      End
   End
End
Attribute VB_Name = "Siniflar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Unload Me
anamenu.Show
End Sub

Private Sub Duzenle_Click()
On Local Error Resume Next
sinifduzenle.Show
End Sub

Private Sub Ekle_Click()
yenisinif.Show
End Sub

Private Sub Form_Load()
Call yukle
End Sub

Sub yukle()
Dim X As Integer
Call veri_ac(False, False)
Call tablo_ac("Select * from siniflar")
sinifliste.Cols = 7
sinifliste.Rows = 1

sinifliste.TextMatrix(0, 0) = "ID"
sinifliste.TextMatrix(0, 1) = "SINIF"
sinifliste.TextMatrix(0, 2) = "KAPAS�TE"
sinifliste.TextMatrix(0, 3) = "SINIF ��R."
sinifliste.TextMatrix(0, 4) = " KURS BA�. TAR."
sinifliste.TextMatrix(0, 5) = "KURS BUT. TAR."
sinifliste.TextMatrix(0, 6) = "TOP. KUR. S�R."
sinifliste.ColWidth(0) = 500
sinifliste.ColWidth(1) = 4000
sinifliste.ColWidth(3) = 3000
sinifliste.ColWidth(4) = 2000
sinifliste.ColWidth(5) = 2000
sinifliste.ColWidth(6) = 2000

X = 0
Do While Not tablo.EOF
X = X + 1
sinifliste.AddItem ""
sinifliste.TextMatrix(X, 0) = tablo("id")
sinifliste.TextMatrix(X, 1) = tablo("sinif")
sinifliste.TextMatrix(X, 2) = tablo("kapasite")
sinifliste.TextMatrix(X, 3) = tablo("sinifogretmeni")
sinifliste.TextMatrix(X, 4) = tablo("kursbastarihi")
sinifliste.TextMatrix(X, 5) = tablo("kursbittarihi")
sinifliste.TextMatrix(X, 6) = tablo("toplamkurssaat")
tablo.MoveNext
Loop
tablo.Close
veri.Close
End Sub

Private Sub Sil_Click()
Dim secim, silinecek
silinecek = Siniflar.sinifliste.TextMatrix(secilisatir, 1)
If silinecek = "ID" Or silinecek = "SINIF" Or silinecek = "KAPAS�TE" Or silinecek = "SINIF ��R." Then
MsgBox "L�tfen silinecek s�n�f� se�iniz ", vbCritical, "Silme i�lemleri"
Exit Sub
End If
secim = MsgBox(silinecek & " isimli s�n�f� silmek istedi�inize eminmisiniz", vbYesNo + vbExclamation)
If secim = vbYes Then
    Call veri_ac(False, False)
    Call tablo_ac("Select * from siniflar where id=" & silinecek)
    tablo.Delete
        MsgBox "S�n�f ba�ar�l� bir �ekilde silinmi�tir", vbInformation
    Call yukle
Else
    MsgBox "Silme i�lemini iptal ettiniz"
End If

End Sub

Private Sub sinifliste_Click()
On Local Error Resume Next
Dim X, Y As Integer
Y = sinifliste.Row
For X = 1 To sinifliste.Rows - 1
    sinifliste.Row = X
    sinifliste.Col = 0
    sinifliste.CellBackColor = sinifliste.BackColor
    sinifliste.Col = 1
    sinifliste.CellBackColor = sinifliste.BackColor
    sinifliste.Col = 2
    sinifliste.CellBackColor = sinifliste.BackColor
    sinifliste.Col = 3
    sinifliste.CellBackColor = sinifliste.BackColor
     sinifliste.Col = 4
    sinifliste.CellBackColor = sinifliste.BackColor
     sinifliste.Col = 5
    sinifliste.CellBackColor = sinifliste.BackColor
     sinifliste.Col = 6
    sinifliste.CellBackColor = sinifliste.BackColor
Next

If sinifliste.Row = 0 Then
    sinifliste.Col = 0
    sinifliste.CellBackColor = 12632256
    sinifliste.Col = 1
    sinifliste.CellBackColor = 12632256
    sinifliste.Col = 2
    sinifliste.CellBackColor = 12632256
    sinifliste.Col = 3
    sinifliste.CellBackColor = 12632256
    sinifliste.Col = 4
    sinifliste.CellBackColor = 12632256
    sinifliste.Col = 5
    sinifliste.CellBackColor = 12632256
    sinifliste.Col = 6
    sinifliste.CellBackColor = 12632256
    Exit Sub
End If
    
    sinifliste.Row = Y
    sinifliste.Col = 0
    sinifliste.CellBackColor = 4326608
    sinifliste.Col = 1
    sinifliste.CellBackColor = 4326608
    sinifliste.Col = 2
    sinifliste.CellBackColor = 4326608
    sinifliste.Col = 3
    sinifliste.CellBackColor = 4326608
    sinifliste.Col = 4
    sinifliste.CellBackColor = 4326608
    sinifliste.Col = 5
    sinifliste.CellBackColor = 4326608
    sinifliste.Col = 6
    sinifliste.CellBackColor = 4326608
    
    secilisatir = Y


End Sub

Private Sub sinifliste_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If jmenu.Visible = True Then GoTo son
If Button = 2 Then
PopupMenu jmenu
End If
son:
End Sub

Private Sub sinifyazdir_Click()
secilisinifilistele.Show
Unload Me
End Sub
