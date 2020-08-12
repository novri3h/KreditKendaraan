VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BayarCicilan 
   Caption         =   "Pembayaran Cicilan"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
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
   ScaleHeight     =   7860
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKeterangan 
      Height          =   350
      Left            =   1800
      TabIndex        =   41
      Text            =   "-"
      Top             =   3720
      Width           =   5205
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   1600
   End
   Begin VB.TextBox TxtAngsuran 
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   2520
      Width           =   1600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   350
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   4200
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   2040
      TabIndex        =   3
      Top             =   4200
      Width           =   1000
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   345
      Left            =   5040
      Top             =   4200
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   345
      Left            =   3120
      Top             =   4200
      Visible         =   0   'False
      Width           =   1900
      _ExtentX        =   3360
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BayarCicilan.frx":0000
      Height          =   1500
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   2646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tabel Kredit"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "BayarCicilan.frx":0015
      Height          =   1500
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   2646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tabel Bayar Cicilan"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Keterangan"
      Height          =   345
      Left            =   120
      TabIndex        =   42
      Top             =   3720
      Width           =   1605
   End
   Begin VB.Label LblTerlambat 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   40
      Top             =   1080
      Width           =   1605
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Terlambat (hari)"
      Height          =   345
      Left            =   3600
      TabIndex        =   39
      Top             =   1080
      Width           =   1605
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Warna Motor"
      Height          =   345
      Left            =   120
      TabIndex        =   38
      Top             =   2880
      Width           =   1605
   End
   Begin VB.Label LblWarna 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   37
      Top             =   2880
      Width           =   1605
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Denda"
      Height          =   345
      Left            =   3600
      TabIndex        =   36
      Top             =   1440
      Width           =   1605
   End
   Begin VB.Label LblDenda 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   35
      Top             =   1440
      Width           =   1605
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jatuh Tempo Tgl"
      Height          =   345
      Left            =   3600
      TabIndex        =   34
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label LblTanggalTempo 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   33
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   32
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Label LblTelepon 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   31
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HP"
      Height          =   345
      Left            =   120
      TabIndex        =   30
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label LblHP 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   29
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label LblTelahBayar 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   28
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telah Dibayar"
      Height          =   345
      Left            =   3600
      TabIndex        =   27
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa Bulan Lalu"
      Height          =   345
      Left            =   3600
      TabIndex        =   26
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label LblSisaLalu 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   25
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label LblSisaSekarang 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   24
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa Pembayaran"
      Height          =   345
      Left            =   3600
      TabIndex        =   23
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label LblAlamat 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   22
      Top             =   1440
      Width           =   1600
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   1600
   End
   Begin VB.Label LblNama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   20
      Top             =   1080
      Width           =   1600
   End
   Begin VB.Label LblMerk 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   19
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label LblHargaKredit 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   18
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Merk Motor"
      Height          =   345
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Harga Kredit"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1605
   End
   Begin VB.Label LblTanggalbyr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   15
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label LblCicilanKe 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5400
      TabIndex        =   14
      Top             =   2880
      Width           =   1605
   End
   Begin VB.Label NomorByr 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   13
      Top             =   120
      Width           =   1600
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor Pembayaran"
      Height          =   350
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1600
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ID Kredit"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1605
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   350
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1600
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Bayar"
      Height          =   345
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Angsuran + Denda"
      Height          =   345
      Left            =   3600
      TabIndex        =   8
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cicilan Ke "
      Height          =   345
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Width           =   1605
   End
End
Attribute VB_Name = "BayarCicilan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBKredit.mdb"
Adodc1.RecordSource = "BeliKredit"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBKredit.mdb"
Adodc2.RecordSource = "bayarcicilan"
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2
DataGrid2.Refresh

RSBeliKredit.Open "select * from belikredit where keterangan <>'LUNAS'", CONN
Combo1.Clear
Do While Not RSBeliKredit.EOF
    Combo1.AddItem RSBeliKredit!IdKredit
    RSBeliKredit.MoveNext
Loop

Call Auto
LblTanggalbyr = Date
End Sub

Private Sub Auto()
Call BukaDB
RSBayarCicilan.Open "select * from Bayarcicilan Where NomorByr In(Select Max(NomorByr)From Bayarcicilan)Order By NomorByr Desc", CONN
RSBayarCicilan.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSBayarCicilan
        If .EOF Then
            Urutan = "BY" + Format(Date, "yymmdd") + "01"
            NomorByr = Urutan
        Else
            If Mid(!NomorByr, 3, 6) <> Format(Date, "yymmdd") Then
                Urutan = "BY" + Format(Date, "yymmdd") + "01"
            Else
                Hitung = Right(!NomorByr, 2) + 1
                Urutan = "BY" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        NomorByr = Urutan
    End With
End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "nomor kredit harus diisi"
        Combo1.SetFocus
    Else
        TxtAngsuran.SetFocus
    End If
End If
End Sub

Private Sub COMBO1_Click()
Call BukaDB
RSBeliKredit.Open "select * from belikredit where idkredit='" & Combo1 & "'", CONN
If Not RSBeliKredit.EOF Then
    'jika belum pernah membayar angsuran maka
    'jatuh tempo pembayaran adalah dimulai dari tanggal beli + 30 hari
    If RSBeliKredit!angsuranke = 0 Then
        LblTanggalTempo = RSBeliKredit!Tanggal + (30 * 1)
    Else
    'jika pernah ada angsuran, maka angsuran berikutnya
    'adalah 30 hari X jumlah angsuran yang penah dibayar
        LblTanggalTempo = RSBeliKredit!Tanggal + (30 * (RSBeliKredit!angsuranke + 1))
    End If
    'jumlah denda adalah 5000 x hari keterlambatan dati tgl jatuh tempo
    If CDate(lbltanggalbayar) > CDate(LblTanggalTempo) Then
        LblTerlambat = CDate(lbltanggalbayar) - CDate(LblTanggalTempo)
        LblDenda = 5000 * LblTerlambat
    Else
        LblTerlambat = 0
        LblDenda = 0
    End If
    
    LblHargaKredit = Format(RSBeliKredit!harga, "###,###,###,###")
    If RSBeliKredit!telahbayar = 0 Then
        LblTelahBayar = 0
    Else
        LblTelahBayar = Format(RSBeliKredit!telahbayar, "###,###,###,###")
    End If
    
'    TxtAngsuran = Format(RSBeliKredit!angsuran, "###,###,###,###")
'    LblSisaLalu = Format(RSBeliKredit!sisa, "###,###,###,###")
    
    'mencari identitas customer yang dihasilkan dari query belikredit
    RSCustomer.Open "select * from customer where kodecus='" & RSBeliKredit!Kodecus & "'", CONN
    If Not RSCustomer.EOF Then
        LblNama = RSCustomer!Nama
        LblAlamat = RSCustomer!Alamat
        LblTelepon = RSCustomer!Telepon
        LblHP = RSCustomer!HP
    End If
    
    'mencari identitas motor yang dihasilkan dari query belikredit
    RSMotor.Open "select * from Motor where kodemtr='" & RSBeliKredit!Kodemtr & "'", CONN
    If Not RSMotor.EOF Then
        LblMerk = RSMotor!merk
        LblWarna = RSMotor!warna
    End If
End If
TxtAngsuran = Format(RSBeliKredit!angsuran, "###,###,###,###")
    LblSisaLalu = Format(RSBeliKredit!sisa, "###,###,###,###")
End Sub

Private Sub Command1_Click()
If Combo1 = "" Or TxtAngsuran = "" Or LblSisaSekarang = "" Then
    MsgBox "Data belum lengkap, coba enter di angsuran + denda"
    TxtAngsuran.SetFocus
    Exit Sub
End If
Call BukaDB
SIMPANBAYARCICILAN = "INSERT INTO bayarcicilan(nomorbyr,tanggalbyr,idkredit,JUMLAH,sisa,CICILAN,keterangan) VALUES " & _
"('" & NomorByr & "','" & LblTanggalbyr & "','" & Combo1 & "','" & TxtAngsuran & "','" & LblSisaSekarang & "','" & LblCicilanKe & "','" & TxtKeterangan & "')"
CONN.Execute SIMPANBAYARCICILAN

'sisa pembayaran terus berkurang akibat pembayaran
'jumlah telah bayar terus bertambah
'jika sisa sekarang = 0 maka keterangan =lunas
'indikasi angsuran terus berubah 1,2,3 dan seterusnya
RSBeliKredit.Open "SELECT * FROM BELIKREDIT WHERE IDKREDIT='" & Combo1 & "'", CONN
If Not RSBeliKredit.EOF Then
    If LblSisaSekarang = 0 Then
        updatedata = "UPDATE BeliKredit SET SISA='" & LblSisaSekarang & "',telahbayar= '" & RSBeliKredit!telahbayar + TxtAngsuran & "',ANGSURANKE='" & LblCicilanKe & "',keterangan='LUNAS' WHERE idkredit='" & Combo1 & "'"
        CONN.Execute updatedata
        CONN.Close
    Else
        updatedata = "UPDATE BeliKredit SET SISA='" & RSBeliKredit!sisa - TxtAngsuran & "',telahbayar= '" & RSBeliKredit!telahbayar + TxtAngsuran & "',ANGSURANKE='" & LblCicilanKe & "',keterangan='-' WHERE idkredit='" & Combo1 & "'"
        CONN.Execute updatedata
        CONN.Close
    End If
    
    Call BukaDB
    RSBeliKredit.Open "SELECT * FROM BeliKredit WHERE IDKredit='" & Combo1 & "' AND SISA=0", CONN
    If Not RSBeliKredit.EOF Then
        UBAHKET = "UPDATE BeliKredit SET KETerangan='LUNAS' WHERE IDKredit='" & Combo1 & "'"
        CONN.Execute UBAHKET
    End If
    Form_Activate
    Call Bersihkan
    Combo1.SetFocus
End If
End Sub


Private Sub TxtAngsuran_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSBeliKredit.Open "SELECT * FROM belikredit WHERE idkredit='" & Combo1 & "'", CONN
    'jika angsuran melebihi sisa pembayaran,
    'maka tampilkan dalam keterangan uang kembaliannya
    If Val(TxtAngsuran) > RSBeliKredit!sisa Then
        TxtAngsuran = Format(TxtAngsuran, "###,###,###,###")
        TxtKeterangan = "kembali" & Space(1) & Format(TxtAngsuran - RSBeliKredit!sisa, "###,###,###,###") & Space(1) & "LUNAS"
        LblCicilanKe = 1
        LblSisaSekarang = 0
    Else
        'sisa sekarang tampil setelah dikurang angsuran
        'indikasi cicilan terus berubah yaitu cicilan bulan lalu + 1
        LblSisaSekarang = Format(LblSisaLalu - TxtAngsuran, "###,###,###,###")
        RSBayarCicilan.Open "SELECT COUNT(idkredit) AS KETEMU FROM bayarcicilan WHERE idkredit='" & Combo1 & "'", CONN
        If Not RSBayarCicilan.EOF Then
            LblCicilanKe = RSBayarCicilan!ketemu + 1
        Else
            LblCicilanKe = 1
        End If
        'tampilkan dalam keterangan indikasi pembayaran bulan jatuh tempo
        TxtKeterangan = "Pembayaran Bulan" & Space(1) & Format(LblTanggalTempo, "MMMM YYYY")
    End If
    TxtKeterangan.SetFocus
End If
End Sub

Private Sub TxtKeterangan_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If TxtKeterangan = "" Then
        TxtKeterangan = "-"
    Else
        Command1.SetFocus
    End If
End If
End Sub

Private Sub Command2_Click()
Call Bersihkan
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Sub Bersihkan()
Combo1 = ""
LblNama = ""
LblAlamat = ""
LblTelepon = ""
LblHP = ""
LblMerk = ""
LblWarna = ""
LblHargaKredit = ""
LblTanggalTempo = ""
LblTerlambat = ""
LblTelahBayar = ""
LblSisaLalu = ""
LblDenda = ""
TxtAngsuran = ""
LblCicilanKe = ""
LblSisaSekarang = ""
TxtKeterangan = "-"
End Sub

