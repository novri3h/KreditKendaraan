VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapPembelian 
   Caption         =   "Laporan Pembelian Cash Dan Kredit"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
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
   ScaleHeight     =   3885
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Laporan Pembelian Kredit"
      Height          =   3615
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Cetak Semua Data"
         Height          =   735
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   2775
      End
      Begin VB.ComboBox Combo6 
         Height          =   345
         Left            =   1320
         TabIndex        =   12
         Top             =   1920
         Width           =   1750
      End
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1320
         TabIndex        =   11
         Top             =   1560
         Width           =   1750
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Bulanan"
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bulan"
         Height          =   345
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahun"
         Height          =   345
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Harian"
         Height          =   225
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Pembelian Cash"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak Semua Data"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Width           =   1750
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1750
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Harian"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Laporan Bulanan"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1350
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   3360
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "LapPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
'cari data tanggal di tabel belicash
RSBeliCash.Open "Select Distinct Tanggal From BeliCash order By 1", CONN
RSBeliCash.Requery
Do Until RSBeliCash.EOF
    'tampilkan dalam combo1
    Combo1.AddItem Format(RSBeliCash!Tanggal, "DD-MMM-YYYY")
    RSBeliCash.MoveNext
Loop

Dim RSBulan As New ADODB.Recordset
'cari bulan dalam tabel belicash
RSBulan.Open "select distinct month(Tanggal) as Bulan from BeliCash", CONN
Do While Not RSBulan.EOF
    'tampilkan dalam combo2
    Combo2.AddItem RSBulan!Bulan & Space(5) & MonthName(RSBulan!Bulan)
    RSBulan.MoveNext
Loop

Dim RSTahun As New ADODB.Recordset
'cari tahun di tabel belicash
RSTahun.Open "select distinct year(Tanggal)  as Tahun from BeliCash", CONN
Do While Not RSTahun.EOF
    'tampilkan dalam combo3
    Combo3.AddItem RSTahun!Tahun
    RSTahun.MoveNext
Loop


RSBeliKredit.Open "Select Distinct Tanggal From BeliKredit order By 1", CONN
RSBeliKredit.Requery
Do Until RSBeliKredit.EOF
    Combo4.AddItem Format(RSBeliKredit!Tanggal, "DD-MMM-YYYY")
    RSBeliKredit.MoveNext
Loop

Dim RSBulanKredit As New ADODB.Recordset
RSBulanKredit.Open "select distinct month(Tanggal) as Bulan from BeliKredit", CONN
Do While Not RSBulanKredit.EOF
    Combo5.AddItem RSBulanKredit!Bulan & Space(5) & MonthName(RSBulanKredit!Bulan)
    RSBulanKredit.MoveNext
Loop

Dim RSTahunKredit As New ADODB.Recordset
RSTahunKredit.Open "select distinct year(Tanggal)  as Tahun from BeliKredit", CONN
Do While Not RSTahunKredit.EOF
    Combo6.AddItem RSTahunKredit!Tahun
    RSTahunKredit.MoveNext
Loop

CONN.Close
End Sub

Private Sub COMBO1_Click()
    CR.SelectionFormula = "Totext({BeliCash.Tanggal})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\lap beli cash harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo3_Click()
    Call BukaDB
    RSBeliCash.Open "select * from BeliCash where month(Tanggal)='" & Val(Left(Combo2, 2)) & "' and year(Tanggal)='" & (Combo3) & "'", CONN
    If RSBeliCash.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({BeliCash.Tanggal})=" & Val(Left(Combo2, 2)) & " and Year({BeliCash.Tanggal})=" & Val(Combo3.Text)
    CR.ReportFileName = App.Path & "\LAP beli cash bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo4_Click()
    CR.SelectionFormula = "Totext({BeliKredit.Tanggal})='" & CDate(Combo4) & "'"
    CR.ReportFileName = App.Path & "\LAP BELI KREDIT HARIAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo6_Click()
Call BukaDB
RSBeliKredit.Open "select * from BeliKredit where month(Tanggal)='" & Val(Left(Combo5, 2)) & "' and year(Tanggal)='" & (Combo6) & "'", CONN
If RSBeliKredit.EOF Then
    MsgBox "Data tidak ditemukan"
    Exit Sub
    Combo4.SetFocus
End If
CR.SelectionFormula = "Month({BeliKredit.Tanggal})=" & Val(Left(Combo5, 2)) & " and Year({BeliKredit.Tanggal})=" & Val(Combo6.Text)
CR.ReportFileName = App.Path & "\LAP BELI KREDIT BULANAN.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 1

End Sub


Private Sub Command1_Click()
    CR.ReportFileName = App.Path & "\lap beli cash.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Command2_Click()
    CR.ReportFileName = App.Path & "\lap BELI KREDIT.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

