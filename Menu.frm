VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar STBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2595
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnmotor 
         Caption         =   "Motor"
      End
      Begin VB.Menu mncustomer 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnoperator 
         Caption         =   "Operator"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mncash 
         Caption         =   "Pembelian Cash"
      End
      Begin VB.Menu mnkredit 
         Caption         =   "Pembelian Kredit"
      End
      Begin VB.Menu mnbayarcicilan 
         Caption         =   "Pembayaran Cicilan"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapmotor 
         Caption         =   "Laporan Data Motor"
      End
      Begin VB.Menu mnlapcustomer 
         Caption         =   "Laporan Data Customer"
      End
      Begin VB.Menu mnlapbeli 
         Caption         =   "Laporan Pembelian"
      End
      Begin VB.Menu mnlapbayar 
         Caption         =   "Laporan Pembayaran"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub



Private Sub mnbayarcicilan_Click()
BayarCicilan.Show vbModal
End Sub

Private Sub mncash_Click()
BeliCash.Show vbModal
End Sub

Private Sub mncustomer_Click()
Customer.Show vbModal
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnkredit_Click()
BeliKredit.Show vbModal
End Sub

Private Sub mnlapbayarperid_Click()
    CR.ReportFileName = App.Path & "\lap bayar cicilan per id.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub mnlapcash_Click()
    CR.ReportFileName = App.Path & "\lap beli cash.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub mnlapcicilan_Click()
    CR.ReportFileName = App.Path & "\lap bayar cicilan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub mnlapbayar_Click()
LapPembayaran.Show vbModal
End Sub

Private Sub mnlapbeli_Click()
LapPembelian.Show vbModal
End Sub

Private Sub mnlapcustomer_Click()
    CR.ReportFileName = App.Path & "\lap customer.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub mnlapkredit_Click()
    CR.ReportFileName = App.Path & "\lap beli kredit.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub mnlapmotor_Click()
    CR.ReportFileName = App.Path & "\lap motor.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1

End Sub

Private Sub mnmotor_Click()
Motor.Show vbModal
End Sub

Private Sub mnoperator_Click()
Operator.Show vbModal
End Sub

Private Sub mnsql_Click()
UjiSQL.Show
End Sub

