VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3630
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
   ScaleHeight     =   1470
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   2000
   End
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2000
      End
      Begin VB.TextBox Text2 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1000
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   2000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1005
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Text1.MaxLength = 30
Text2.MaxLength = 10
Text2.PasswordChar = "X"
Text2.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then End
If Keyascii = 13 Then
    Call BukaDB
    RSOperator.Open "Select NamaOPR from Operator where NamaOPR ='" & Text1 & "'", CONN
    If RSOperator.EOF Then
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & Text1 & "' tidak dikenal"
            Text1 = ""
            Text1.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & Text1 & "' tidak dikenal"
            Text1 = ""
            Text1.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & Text1 & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            Unload Me
        End If
    Else
        Text1.Enabled = False
        Text2.Enabled = True
        Text2.SetFocus
    End If
End If
End Sub


Private Sub text2_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then End
Dim KodeOperator As String
Dim NamaOperator As String
If Keyascii = 13 Then
    Call BukaDB
    RSOperator.Open "Select * from Operator where NamaOPR ='" & Text1 & "' and PasswordOPR='" & Text2 & "'", CONN
    If RSOperator.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            Text2 = ""
            Text2.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            Text2 = ""
            Text2.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            Unload Me
        End If
    Else
        Text3 = RSOperator!KodeOPR
        Text4.Enabled = False
        Text4 = RSOperator!StatusOPR
        
        Me.Visible = False
        Menu.Show
        
        Menu.STBar.Panels(1).Text = Login.Text1
        Menu.STBar.Panels(2).Text = Login.Text4
        Menu.STBar.Panels(3).Text = Login.Text3
        Menu.STBar.Panels(3).Visible = False
        
        If Menu.STBar.Panels(2).Text = "USER" Then
            Menu.mnoperator.Enabled = False
'            'Menu.mnadm.Enabled = False
'        ElseIf Menu.STBar.Panels(2).Text = "ADM" Then
'            Menu.mnfile.Enabled = False
'            Menu.mnapoteker.Enabled = False
        End If
    End If
End If
End Sub

