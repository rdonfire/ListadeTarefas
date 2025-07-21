Attribute VB_Name = "Module1"
Option Explicit

Public conn As New ADODB.Connection
Public Sub ConectarBD()
On Error GoTo ErroConexao

Dim connectionString As String


connectionString = "Provider=SQLOLEDB;Server=DEV_AMORIM\PDVNET;Database=AGENDA;User Id=SA;Password=inter#system;"

conn.Open connectionString
Exit Sub

ErroConexao:
MsgBox "Erro ao conectar ao banco de dados: " & Err.Description, vbCritical
End
End Sub
