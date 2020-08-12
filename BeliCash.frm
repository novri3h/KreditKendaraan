VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form BeliCash 
   Caption         =   "Pembelian Cash"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
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
   ScaleHeight     =   4950
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   3720
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton Cmdtutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   1000
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   840
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   5040
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox TxtKet 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
      Width           =   1500
   End
   Begin VB.TextBox TxtDibayar 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5040
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BeliCash.frx":0000
      Height          =   1815
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3201
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4560
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.Label LblNama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5040
      TabIndex        =   27
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label LblAlamat 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5040
      TabIndex        =   26
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label LblTelepon 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5040
      TabIndex        =   25
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label LblMerk 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   24
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label LblWarna 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   23
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label LblHarga 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   22
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dibayar"
      Height          =   345
      Left            =   3480
      TabIndex        =   21
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Keterangan"
      Height          =   345
      Left            =   3480
      TabIndex        =   20
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ID Cash"
      Height          =   345
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Beli"
      Height          =   345
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label IdCash 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   16
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1800
      TabIndex        =   15
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Motor"
      Height          =   345
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Merk"
      Height          =   345
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Warna"
      Height          =   345
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Harga"
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Customer"
      Height          =   345
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   3480
      TabIndex        =   9
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   3480
      TabIndex        =   8
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   3480
      TabIndex        =   7
      Top             =   1200
      Width           =   1500
   End
End
Attribute VB_Name = "BeliCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBKredit.mdb"
Adodc1.RecordSource = "belicash"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

'menampilkan daftar kode customer dalam combo1
Call BukaDB
RSCustomer.Open "Customer", CONN
Combo1.Clear
Do Until RSCustomer.EOF
    Combo1.AddItem RSCustomer!Kodecus
    RSCustomer.MoveNext
Loop

'menampilkan daftar kode motor di combo2
RSMotor.Open "Motor", CONN
Combo2.Clear
Do Until RSMotor.EOF
    Combo2.AddItem RSMotor!Kodemtr
    RSMotor.MoveNext
Loop

Call Auto 'memanggil IDCash otomatis dengan pola tanggal
Tanggal = Date
End Sub

'memanggil IDCash otomatis dengan pola tanggal
'buka tabel becash dan cari IDCash yang paling besar
'jika tidak ada maka dibentuk yang baru
'jika sudah ada yang yang paling besar + 1
Private Sub Auto()
Call BukaDB
RSBeliCash.Open "select * from BeliCash Where IdCash In(Select Max(IdCash)From BeliCash)Order By IdCash Desc", CONN
RSBeliCash.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSBeliCash
        If .EOF Then
            Urutan = "CS" + Format(Date, "yymmdd") + "01"
            IdCash = Urutan
        Else
            If Mid(!IdCash, 3, 6) <> Format(Date, "yymmdd") Then
                Urutan = "CS" + Format(Date, "yymmdd") + "01"
            Else
                Hitung = Right(!IdCash, 2) + 1
                Urutan = "CS" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        IdCash = Urutan
    End With
End Sub

'menampilkan identitas customer yang dipilih di combo1
Private Sub COMBO1_Click()
Call BukaDB
RSCustomer.Open "select * from customer where kodecus='" & Combo1 & "'", CONN
If RSCustomer.EOF Then
    MsgBox "kode customer tidak terdaftar"
    Combo1.SetFocus
Else
    LblNama = RSCustomer!Nama
    LblAlamat = RSCustomer!Alamat
    LblTelepon = RSCustomer!Telepon
End If
End Sub

'menampilkan identitas motor yang dipilih di combo2
Private Sub Combo2_Click()
Call BukaDB
RSMotor.Open "select * from Motor where kodemtr='" & Combo2 & "'", CONN
If RSMotor.EOF Then
    MsgBox "kode Motor tidak terdaftar"
    Combo2.SetFocus
Else
    LblMerk = RSMotor!merk
    LblWarna = RSMotor!warna
    LblHarga = Format(RSMotor!harga, "###,###,###,###")
End If
End Sub

Private Sub TxtDibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        TxtDibayar = Format(TxtDibayar, "###,###,###")
        If TxtDibayar = "" Or TxtDibayar < LblHarga Then
            TxtKet = "Kurang" & Space(1) & Format(LblHarga - TxtDibayar, "###,###,###")
            CmdSimpan.Enabled = True
            CmdSimpan.SetFocus
        Else
            If TxtDibayar = LblHarga Then
                TxtKet = 0
            Else
                TxtKet = "Kembali" & Space(1) & Format(TxtDibayar - LblHarga, "###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub CmdSimpan_Keypress(Keyascii As Integer)
    If Keyascii = 27 Then
        TxtDibayar = ""
        TxtKet = ""
        TxtDibayar.SetFocus
    End If
End Sub

Private Sub CmdSimpan_Click()
If Combo1 = "" Or Combo2 = "" Or TxtDibayar = "" Or TxtKet = "" Then
    MsgBox "data belum lengkap"
Else
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into BeliCash(IdCash,Tanggal,kodecus,kodemtr,harga,dibayar,keterangan)" & _
    "values('" & IdCash & "','" & Tanggal & "','" & Combo1 & "','" & Combo2 & "','" & LblHarga & "','" & TxtDibayar & "','" & TxtKet & "')"
    CONN.Execute (SQLTambahJual)
    Form_Activate
    Call Bersihkan
    Form_Activate
    Call cetak
End If
End Sub

Sub cetak()
    CR.ReportFileName = App.Path & "\kwitansi beli cash.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Bersihkan()
    Combo1 = ""
    Combo2 = ""
    LblNama = ""
    LblAlamat = ""
    LblTelepon = ""
    LblMerk = ""
    LblWarna = ""
    LblHarga = ""
    TxtDibayar = ""
    TxtKet = ""
End Sub

Private Sub CmdBatal_Click()
Call Bersihkan
Form_Activate
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub




