Attribute VB_Name = "DB"
' ================================================================
' ===================== Payroll System ===========================
' ================================================================
' Copyright (C) 2013  Jhon Kenneth N. Carino
'                     Email Address: jkennethcarino@yahoo.com
'                     Contact No:  +639163369826
'
' ============= Premiere Computer Learning Center ================

Public EmpNumb As String
Public EmpPosition As String
Public conn As New ADODB.Connection
Public RS As New ADODB.Recordset
Public cmd As New ADODB.Command
Public SQL As String


Public Sub connectDB()
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=./Database/PayrollSystem.mdb;Persist Security Info=False"
conn.Open
With RS
    .ActiveConnection = conn
    .Open SQL, conn, 3, 3
End With
End Sub


Public Sub connOpen()
' For ListView Records
' If working with "LIKE" SQL Statement then use this
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=./Database/PayrollSystem.mdb;Persist Security Info=False"
conn.Open

With cmd
    .ActiveConnection = conn
    .CommandType = adCmdText
    .CommandText = SQL
    Set RS = .Execute
End With
End Sub

Public Sub connClose()
' Close Database Connection
Set conn = Nothing
Set RS = Nothing
End Sub
