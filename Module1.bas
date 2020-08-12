Attribute VB_Name = "Module1"

Public CONN As New ADODB.Connection
Public RSMotor As ADODB.Recordset
Public RSCustomer As ADODB.Recordset
Public RSOperator As ADODB.Recordset
Public RSBeliCash As ADODB.Recordset
Public RSBeliKredit As ADODB.Recordset
Public RSDetailKredit As ADODB.Recordset
Public RSBayarCicilan As ADODB.Recordset

Public Sub BukaDB()
Set CONN = New ADODB.Connection
Set RSMotor = New ADODB.Recordset
Set RSCustomer = New ADODB.Recordset
Set RSOperator = New ADODB.Recordset
Set RSBeliCash = New ADODB.Recordset
Set RSBeliKredit = New ADODB.Recordset
Set RSDetailKredit = New ADODB.Recordset
Set RSBayarCicilan = New ADODB.Recordset
CONN.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBkredit.mdb"
End Sub


