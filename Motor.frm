VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Motor 
   Caption         =   "Data Motor"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5970
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
   ScaleHeight     =   5430
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5655
      Begin VB.TextBox Text4 
         Height          =   350
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1250
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   1250
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   4260
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1250
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3480
         Top             =   1200
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
         CommandType     =   2
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
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   5295
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Harga"
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Warna"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Merk"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Kode"
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   5655
      Begin VB.CommandButton Cmdinput 
         Caption         =   "&Input"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdtutup 
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1845
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3254
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
End
Attribute VB_Name = "Motor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBkredit.mdb"
Adodc1.RecordSource = "Motor"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Form_Load()
Call BukaDB
Text1.MaxLength = 5
Text2.MaxLength = 10
Text3.MaxLength = 10
Text4.MaxLength = 8
KondisiAwal
End Sub

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSMotor
        If Not RSMotor.EOF Then
            Text2 = RSMotor!merk
            Text3 = RSMotor!warna
            Text4 = RSMotor!harga
        End If
    End With
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Text1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Motor (KodeMtr,Merk,Warna,Harga) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
            CONN.Execute SQLTambah
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        If Text2 = "" Or Text3 = "" Or Text4 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Motor Set Merk= '" & Text2 & "', Warna='" & Text3 & "', Harga='" & Text4 & "' where KodeMtr='" & Text1 & "'"
            CONN.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Function CariData()
    Call BukaDB
    RSMotor.Open "Select * From Motor where KodeMtr='" & Text1 & "'", CONN
End Function

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Text1) < 5 Then
        MsgBox "Kode Harus 5 Digit"
        Text1.SetFocus
        Exit Sub
    Else
        Text2.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
        If Not RSMotor.EOF Then
            TampilkanData
            MsgBox "Kode Motor Sudah Ada"
            KosongkanText
            Text1.SetFocus
        Else
            Text2.SetFocus
        End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
        If Not RSMotor.EOF Then
            TampilkanData
            Text1.Enabled = False
            Text2.SetFocus
        Else
            MsgBox "Kode Motor Tidak Ada"
            Text1 = ""
            Text1.SetFocus
        End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
        If Not RSMotor.EOF Then
            TampilkanData
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim SQLHapus As String
                SQLHapus = "Delete From Motor where kodeMtr= '" & Text1 & "'"
                CONN.Execute SQLHapus
                KondisiAwal
                Form_Activate
            Else
                KondisiAwal
                CmdHapus.SetFocus
            End If
        Else
            MsgBox "Data Tidak ditemukan"
            Text1.SetFocus
        End If
    End If
End If
End Sub

Private Sub text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub text3_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_Change()
Label5.Caption = TerbilangDesimal(Text4.Text)
End Sub

Private Sub Text4_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


Public Function TerbilangDesimal(InputCurrency As String, Optional MataUang As String = "rupiah") As String
 Dim strInput As String
 Dim strBilangan As String
 Dim strPecahan As String
   On Error GoTo Pesan
   Dim strValid As String, huruf As String * 1
   Dim i As Integer
   'Periksa setiap karakter yg diketikkan ke kotak
   'UserID
   strValid = "1234567890,"
   For i% = 1 To Len(InputCurrency)
     huruf = Chr(Asc(Mid(InputCurrency, i%, 1)))
     If InStr(strValid, huruf) = 0 Then
       Set AngkaTerbilang = Nothing
       MsgBox "Harus karakter angka!", _
              vbCritical, "Karakter Tidak Valid"
       Exit Function
     End If
   Next i%
 
 If InputCurrency = "" Then Exit Function
 If Len(Trim(InputCurrency)) > 15 Then GoTo Pesan
 
 strInput = CStr(InputCurrency) 'Konversi ke string
 'Periksa apakah ada tanda "," jika ya berarti pecahan
 If InStr(1, strInput, ",", vbBinaryCompare) Then
      
  strBilangan = Left(strInput, InStr(1, strInput, _
                ",", vbBinaryCompare) - 1)
  'strBilangan = Right(strInput, InStr(1, strInput, _
  '              ".", vbBinaryCompare) - 2)
  strPecahan = Trim(Right(strInput, Len(strInput) - Len(strBilangan) - 1))
  
  If MataUang <> "" Then
      
  If CLng(Trim(strPecahan)) > 99 Then
     strInput = Format(Round(CDbl(strInput), 2), "#0.00")
     strPecahan = Format((Right(strInput, Len(strInput) - Len(strBilangan) - 1)), "00")
    End If
    
    If Len(Trim(strPecahan)) = 1 Then
       strInput = Format(Round(CDbl(strInput), 2), _
                  "#0.00")
       strPecahan = Format((Right(strInput, _
          Len(strInput) - Len(strBilangan) - 1)), "00")
    End If
    
    If CLng(Trim(strPecahan)) = 0 Then
    TerbilangDesimal = (KonversiBilangan(strBilangan) & MataUang & " " & KonversiBilangan(strPecahan))
 Else
  TerbilangDesimal = (KonversiBilangan(strBilangan) & MataUang & " " & KonversiBilangan(strPecahan) & "sen")
    End If
  Else
    TerbilangDesimal = (KonversiBilangan(strBilangan) & "koma " & KonversiPecahan(strPecahan))
  End If
  
 Else
    TerbilangDesimal = (KonversiBilangan(strInput))
  End If
 Exit Function
Pesan:
  TerbilangDesimal = "(maksimal 15 digit)"
End Function

'Fungsi ini untuk mengkonversi nilai pecahan (setelah 'angka 0)
Private Function KonversiPecahan(strAngka As String) As String
Dim i%, strJmlHuruf$, Urai$, Kar$
 If strAngka = "" Then Exit Function
    strJmlHuruf = Trim(strAngka)
    Urai = ""
    Kar = ""
    For i = 1 To Len(strJmlHuruf)
      'Tampung setiap satu karakter ke Kar
      Kar = Mid(strAngka, i, 1)
      Urai = Urai & Kata(CInt(Kar))
    Next i
    KonversiPecahan = Urai
End Function

'Fungsi ini untuk menterjemahkan setiap satu angka ke 'kata
Private Function Kata(angka As Byte) As String
   Select Case angka
          Case 1: Kata = "satu "
          Case 2: Kata = "dua "
          Case 3: Kata = "tiga "
          Case 4: Kata = "empat "
          Case 5: Kata = "lima "
          Case 6: Kata = "enam "
          Case 7: Kata = "tujuh "
          Case 8: Kata = "delapan "
          Case 9: Kata = "sembilan "
          Case 0: Kata = "nol "
   End Select
End Function

'Ini untuk mengkonversi nilai bilangan sebelum pecahan
Private Function KonversiBilangan(strAngka As String) As String
Dim strJmlHuruf$, intPecahan As Integer, strPecahan$, Urai$, Bil1$, strTot$, Bil2$
 Dim X, Y, z As Integer

 If strAngka = "" Then Exit Function
    strJmlHuruf = Trim(strAngka)
    X = 0
    Y = 0
    Urai = ""
    While (X < Len(strJmlHuruf))
      X = X + 1
      strTot = Mid(strJmlHuruf, X, 1)
      Y = Y + Val(strTot)
      z = Len(strJmlHuruf) - X + 1
      Select Case Val(strTot)
      'Case 0
       '   Bil1 = "NOL "
      Case 1
          If (z = 1 Or z = 7 Or z = 10 Or z = 13) Then
              Bil1 = "satu "
          ElseIf (z = 4) Then
              If (X = 1) Then
                  Bil1 = "se"
              Else
                  Bil1 = "satu "
              End If
          ElseIf (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
              X = X + 1
              strTot = Mid(strJmlHuruf, X, 1)
              z = Len(strJmlHuruf) - X + 1
              Bil2 = ""
              Select Case Val(strTot)
              Case 0
                  Bil1 = "sepuluh "
              Case 1
                  Bil1 = "sebelas "
              Case 2
                  Bil1 = "dua belas "
              Case 3
                  Bil1 = "tiga belas "
              Case 4
                  Bil1 = "empat belas "
              Case 5
                  Bil1 = "lima belas "
              Case 6
                  Bil1 = "enam belas "
              Case 7
                  Bil1 = "tujuh belas "
              Case 8
                  Bil1 = "delapan belas "
              Case 9
                  Bil1 = "sembilan belas "
              End Select
          Else
              Bil1 = "se"
          End If
      
      Case 2
          Bil1 = "dua "
      Case 3
          Bil1 = "tiga "
      Case 4
          Bil1 = "empat "
      Case 5
          Bil1 = "lima "
      Case 6
          Bil1 = "enam "
      Case 7
          Bil1 = "tujuh "
      Case 8
          Bil1 = "delapan "
      Case 9
          Bil1 = "sembilan "
      Case Else
          Bil1 = ""
      End Select
       
      If (Val(strTot) > 0) Then
         If (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
            Bil2 = "puluh "
         ElseIf (z = 3 Or z = 6 Or z = 9 Or z = 12 Or z = 15) Then
            Bil2 = "ratus "
         Else
            Bil2 = ""
         End If
      Else
         Bil2 = ""
      End If
      If (Y > 0) Then
          Select Case z
          Case 4
              Bil2 = Bil2 + "ribu "
              Y = 0
          Case 7
              Bil2 = Bil2 + "juta "
              Y = 0
          Case 10
              Bil2 = Bil2 + "milyar "
              Y = 0
          Case 13
              Bil2 = Bil2 + "trilyun "
              Y = 0
          End Select
      End If
      Urai = Urai + Bil1 + Bil2
  Wend
  KonversiBilangan = Urai
End Function




