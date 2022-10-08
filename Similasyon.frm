VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Similasyon 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Resim Ýþlemleri"
   ClientHeight    =   8835
   ClientLeft      =   3165
   ClientTop       =   2265
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   11490
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "ÝPTAL"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "ONAYLA"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ön Ýzleme"
      Height          =   3375
      Left            =   8280
      TabIndex        =   5
      Top             =   720
      Width           =   2775
      Begin VB.Image vesikalik 
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox picCapture 
      BackColor       =   &H80000006&
      Height          =   4335
      Left            =   1320
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   4
      Top             =   1680
      Width           =   5295
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3840
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Webcam'a baðlan"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Resmi Çek"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.ListBox lstDevices 
      Height          =   840
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   135
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   7530
      Left            =   120
      Picture         =   "Similasyon.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   8115
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Aygýt Seçiniz"
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
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Similasyon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const WM_CAP As Integer = &H400

Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Const WM_CAP_EDIT_COPY As Long = WM_CAP + 30

Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Const WM_CAP_SET_SCALE As Long = WM_CAP + 53
Const WS_CHILD As Long = &H40000000
Const WS_VISIBLE As Long = &H10000000
Const SWP_NOMOVE As Long = &H2
Const SWP_NOSIZE As Integer = 1
Const SWP_NOZORDER As Integer = &H4
Const HWND_BOTTOM As Integer = 1

Dim iDevice As Long  ' Current device ID
Dim hHwnd As Long ' Handle to preview window

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean

Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
    (ByVal lpszWindowName As String, ByVal dwStyle As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Integer, ByVal hWndParent As Long, _
    ByVal nID As Long) As Long

Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, _
    ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, _
    ByVal cbVer As Long) As Boolean
    Private X As Byte

Private Sub cmdSave_Click()
    Dim bm As Image
    Dim ism As String
   
   

    '
    ' Copy image to clipboard
    '
    
    SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
    ClosePreviewWindow

    'picCapture.Picture = Clipboard.GetData
    vesikalik.Picture = Clipboard.GetData
    
    'CommonDialog1.CancelError = True
    'CommonDialog1.FileName = "Webcam1"
    'CommonDialog1.Filter = "Bitmap |*.bmp"
     OpenPreviewWindow '
    
    On Error GoTo NoSave
    'CommonDialog1.ShowSave
    ism = "vesikalik.jpg"
    SavePicture vesikalik.Picture, App.Path & "\temp\" & ism  'CommonDialog1.FileName

    vesikalik.Picture = LoadPicture(App.Path & "\temp\" & ism)
    
    
NoSave:
    
End Sub

Private Sub cmdStart_Click()
   
    iDevice = lstDevices.ListIndex
    OpenPreviewWindow
    Shape1.BackColor = &HFF00&

End Sub





Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo hata
Shape1.BackColor = &HFF&


    LoadDeviceList
    
    If lstDevices.ListCount > 0 Then
        lstDevices.Selected(0) = True
    Else
    cmdStart.Enabled = False
        lstDevices.AddItem ("No Device Available")
    End If
    cmdSave.Enabled = False
hata:
    Resume Next

End Sub

Private Sub LoadDeviceList()
    Dim strName As String
    Dim strVer As String
    Dim iReturn As Boolean
    Dim X As Long
    
    X = 0
    strName = Space(100)
    strVer = Space(100)
    '
    ' Load name of all avialable devices into the lstDevices
    '

    Do
        '
        '   Get Driver name and version
        '
        iReturn = capGetDriverDescriptionA(X, strName, 100, strVer, 100)

        '
        ' If there was a device add device name to the list
        '
        If iReturn Then lstDevices.AddItem Trim$(strName)
        X = X + 1
    Loop Until iReturn = False
End Sub

Private Sub OpenPreviewWindow()
    '
    ' Open Preview window in picturebox
    '
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, _
        480, picCapture.hwnd, 0)

    '
    ' Connect to device
    '
    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
        '
        'Set the preview scale
        '
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0

        '
        'Set the preview rate in milliseconds
        '
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0

        '
        'Start previewing the image from the camera
        '
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0

        '
        ' Resize window to fit in picturebox
        '
        SetWindowPos hHwnd, HWND_BOTTOM, 0, 0, picCapture.ScaleWidth, picCapture.ScaleHeight, _
                SWP_NOMOVE Or SWP_NOZORDER

        cmdSave.Enabled = True
        
        cmdStart.Enabled = False
    Else
        '
        ' Error connecting to device close window
        '
        DestroyWindow hHwnd

        cmdSave.Enabled = False
    End If
 End Sub

Private Sub ClosePreviewWindow()
    '
    ' Disconnect from device
    '
    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0
        '
    ' close window
    '

    DestroyWindow hHwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
        X = 0
        ClosePreviewWindow
          
    
End Sub







