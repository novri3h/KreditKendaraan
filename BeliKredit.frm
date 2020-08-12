VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BeliKredit 
   Caption         =   "Transaksi Kredit"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
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
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtBunga 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   1500
   End
   Begin VB.TextBox TxtLama 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   4920
      TabIndex        =   4
      Top             =   1560
      Width           =   1500
   End
   Begin VB.TextBox TxtDP 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   4920
      TabIndex        =   2
      Top             =   840
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4920
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1500
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   840
      Width           =   1500
   End
   Begin VB.CommandButton Cmdtutup 
      Caption         =   "&Tutup"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BeliKredit.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   8
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
      Left            =   4440
      Top             =   2520
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
   Begin VB.Label LblAngsuran 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   27
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Lama Cicilan (Bln)"
      Height          =   345
      Left            =   3360
      TabIndex        =   26
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bunga (%) / Thn"
      Height          =   345
      Left            =   3360
      TabIndex        =   25
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   3360
      TabIndex        =   24
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Customer"
      Height          =   345
      Left            =   3360
      TabIndex        =   23
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Harga Kredit"
      Height          =   345
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Harga Cash"
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Merk"
      Height          =   345
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Motor"
      Height          =   345
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Tanggal 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   18
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label IdKredit 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   17
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Beli"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ID Kredit"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Angsuran/Bln"
      Height          =   345
      Left            =   3360
      TabIndex        =   14
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Muka"
      Height          =   345
      Left            =   3360
      TabIndex        =   13
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label LblHargaKredit 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   12
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label LblHargaCash 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   11
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label LblMerk 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label LblNama 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4920
      TabIndex        =   9
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "BeliKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo2 = "" Then
        MsgBox "kode customer harus diisi"
        Combo2.SetFocus
        Exit Sub
    Else
        TxtDP.SetFocus
    End If
End If
End Sub

Private Sub Combo2_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo2 = "" Then
        MsgBox "kode motor harus diisi"
        Combo2.SetFocus
        Exit Sub
    Else
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBKredit.mdb"
Adodc1.RecordSource = "BeliKredit"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Call BukaDB
RSCustomer.Open "Customer", CONN
Combo1.Clear
Do Until RSCustomer.EOF
    Combo1.AddItem RSCustomer!Kodecus
    RSCustomer.MoveNext
Loop

RSMotor.Open "Motor", CONN
Combo2.Clear
Do Until RSMotor.EOF
    Combo2.AddItem RSMotor!Kodemtr
    RSMotor.MoveNext
Loop

Call Auto
Tanggal = Date
End Sub

Private Sub Auto()
Call BukaDB
RSBeliKredit.Open "select * from BeliKredit Where IdKredit In(Select Max(IdKredit)From BeliKredit)Order By IdKredit Desc", CONN
RSBeliKredit.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSBeliKredit
        If .EOF Then
            Urutan = "CR" + Format(Date, "yymmdd") + "01"
            IdKredit = Urutan
        Else
            If Mid(!IdKredit, 3, 6) <> Format(Date, "yymmdd") Then
                Urutan = "CR" + Format(Date, "yymmdd") + "01"
            Else
                Hitung = Right(!IdKredit, 2) + 1
                Urutan = "CR" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        IdKredit = Urutan
    End With
End Sub

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

Private Sub Combo2_Click()
Call BukaDB
RSMotor.Open "select * from Motor where kodemtr='" & Combo2 & "'", CONN
If RSMotor.EOF Then
    MsgBox "kode Motor tidak terdaftar"
    Combo2.SetFocus
Else
    LblMerk = RSMotor!merk
    LblHargaCash = Format(RSMotor!harga, "###,###,###,###")
End If
End Sub

Private Sub TxtDibayar_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If TxtDibayar = "" Or Val(TxtDibayar) < (LblHarga) Then
            TxtKet = "kurang" & Space(1) & Format(LblHarga - TxtDibayar, "###,###,###")
        Else
            
            If TxtDibayar = LblHarga Then
                TxtKet = TxtDibayar - LblHarga
                TxtDibayar = Format(TxtDibayar, "###,###,###")
            Else
                TxtKet = "kembali" & Space(1) & Format(TxtDibayar - LblHarga, "###,###,###")
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
If Combo1 = "" Or Combo2 = "" Or TxtDP = "" Or TxtBunga = "" Or TxtLama = "" Then
    MsgBox "data belum lengkap"
Else
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into BeliKredit(IdKredit,Tanggal,kodecus,kodemtr,harga,uangmuka,bunga,lamacicilan,angsuran,sisa,keterangan)" & _
    "values('" & IdKredit & "','" & Tanggal & "','" & Combo1 & "','" & Combo2 & "','" & LblHargaKredit & "','" & TxtDP & "','" & TxtBunga & "','" & TxtLama & "','" & LblAngsuran & "','" & LblHargaKredit & "','-')"
    CONN.Execute (SQLTambahJual)
    Form_Activate
    Call Bersihkan
    Form_Activate
    Combo2.SetFocus
End If
End Sub

Private Sub Bersihkan()
    Combo1 = ""
    Combo2 = ""
    LblNama = ""
    TxtDP = ""
    TxtBunga = ""
    TxtLama = ""
    LblMerk = ""
    LblHargaCash = ""
    LblHargaKredit = ""
    LblAngsuran = ""
End Sub

Private Sub CmdBatal_Click()
Call Bersihkan
Form_Activate
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Private Sub TxtBunga_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If TxtBunga = "" Then
        MsgBox "Bunga harus diisi"
        TxtBunga.SetFocus
        Exit Sub
    Else
        TxtLama.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TxtDP_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If TxtDP = "" Then
        MsgBox "Uang Muka harus diisi"
        TxtDP.SetFocus
        Exit Sub
    Else
        TxtDP = Format(TxtDP, "###,###,###,###")
        TxtBunga.SetFocus
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
    
End Sub

'mencari harga motor kredit dan angsuran perbulan
Private Sub TxtLama_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    LblAngsuran = Round(Pmt(TxtBunga / 100 / 12, TxtLama, LblHargaCash), 0) * -1
    LblAngsuran = Format(LblAngsuran, "###,###,###,###")
    LblHargaKredit = Round(FV(TxtBunga / 100 / 12, TxtLama, LblAngsuran), 0) * -1
    LblHargaKredit = Format(LblHargaKredit, "###,###,###,###")
    CmdSimpan.SetFocus
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub
