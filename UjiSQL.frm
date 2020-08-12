VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UjiSQL 
   Caption         =   "Uji Coba SQL"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   5640
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4440
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   7832
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   5640
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4680
      Width           =   13320
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   375
      Left            =   1200
      Top             =   5640
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "DT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label TANGGAL 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Data"
      Height          =   195
      Left            =   7080
      TabIndex        =   2
      Top             =   5760
      Width           =   885
   End
End
Attribute VB_Name = "UjiSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
TANGGAL = Date
On Error Resume Next
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBkredit.mdb"
DT.RecordSource = "select * from bayarcicilan where idkredit='xxx'"
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
Text1 = ""
Text1.SetFocus
Command1.Default = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Command1_Click()
    Dim Pesan As String
    On Error GoTo salah
    DT.RecordSource = Text1
    DT.Refresh
    Text2 = DT.Recordset.RecordCount
    
    If DT.Recordset.EOF Then
        Pesan = MsgBox("Data Tidak Ditemukan")
        DT.Refresh
        Text1.SetFocus
    End If
    On Error GoTo 0
    Exit Sub
salah:
    MsgBox "Syntax SQL Salah..!"
End Sub


