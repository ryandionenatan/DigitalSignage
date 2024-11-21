VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15990
   ControlBox      =   0   'False
   FillColor       =   &H80000001&
   ForeColor       =   &H80000001&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerStop 
      Enabled         =   0   'False
      Interval        =   25000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer TimerTeksBerjalan 
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer TimerArtikel 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin WMPLibCtl.WindowsMediaPlayer wm 
      Height          =   7320
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   12300
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   21696
      _cy             =   12912
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Iklan"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "IKLANKAN BARANG ANDA DISINI HUB: 082387882222"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   12720
      TabIndex        =   5
      Top             =   9240
      Width           =   6855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Waktu"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   44.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   8760
      Width           =   12255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   7920
      Width           =   12255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Marquee"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   885
      Left            =   120
      TabIndex        =   2
      Top             =   10320
      Width           =   2685
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000017&
      Caption         =   "--MAINTENANCE--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   7095
      Left            =   12720
      TabIndex        =   1
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "--UNDER CONSTRUCTION--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   12720
      TabIndex        =   0
      Top             =   600
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable bantu
Dim contrig As Integer

Private Sub Form_Activate()
'Ubah form warna hitam
Me.BackColor = vbBlack
'Hubung ke database digital signage
Call koneksi
'Video
a = "1.m3u"
wm.URL = App.Path & "\video\" & a & ""
wm.Controls.play
'Posisi label awal
Label1.Left = Me.Width
Label2.Top = -8620
Label3.Top = -7080
'Ambil info di tabel
getnews = "select * from berita"
Set rs = con.Execute(getnews)
getad = "select * from informasi"
Set rs2 = con.Execute(getad)
contrig = 2
Label1 = rs2!informasi
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Tombol Escape untuk keluar
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub TimerArtikel_Timer()
'Memasukkan Berita Saat label diluar screen
If Label2.Top = -8620 And Label3.Top = -7080 Then
contrig = 2
End If
If contrig = 2 Then
If Not rs.EOF Then
Label2 = rs!judul
Label3 = rs!artikel
Else
rs.MoveFirst
Label2 = rs!judul
Label3 = rs!artikel
End If
'Menurunkan label
Label2.Top = Label2.Top + 200
Label3.Top = Label3.Top + 200
End If
'Switch ke TimerStop
If Label2.Top > 600 And Label3.Top > 2040 Then
contrig = 1
End If
If contrig = 1 Then
TimerStop.Enabled = True
TimerArtikel.Enabled = False
End If
End Sub

Private Sub TimerStop_Timer()
'Mengembalikan posisi label ke luar screen
Label2.Top = -8620
Label3.Top = -7080
'Mengambil berita berikutnya
rs.MoveNext
TimerArtikel.Enabled = True
TimerStop.Enabled = False
End Sub

Private Sub TimerTeksBerjalan_Timer()
'Kembalikan label teks berjalan ke luar kanan layar
    If (Label1.Left + Label1.Width) <= 0 Then
        Label1.Left = Me.Width
'Panggil informasi berikutnya
        rs2.MoveNext
    End If
    If Not rs2.EOF Then
        Label1 = rs2!informasi
        Label1.Left = Label1.Left - 30
        Else
        rs2.MoveFirst
        Label1 = rs2!informasi
        Label1.Left = Label1.Left - 30
    End If
'Jam dan Tanggal
Label4.Caption = Format(Now, "dddd, d mmmm yyyy")
Label5.Caption = Format(Now, "h:mm:ss")
End Sub
