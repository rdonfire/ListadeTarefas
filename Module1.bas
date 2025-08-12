Attribute VB_Name = "Connections"
Option Explicit

Public conn As New ADODB.Connection
Public Function ConectarBD()
On Error GoTo ErroConexao
Dim connectionString As String
Dim ConectarStatus As Boolean

If conn.State <> adStateOpen Then
    connectionString = "Provider=SQLOLEDB;Server=DEV_AMORIM\PDVNET;Database=AGENDA;User Id=SA;Password=inter#system;"
    conn.Open connectionString
    ConectarStatus = True
    Exit Function
End If

ConectarStatus = False
Exit Function

ErroConexao:
MsgBox "Erro ao conectar ao banco de dados: " & Err.Description, vbCritical
End Function

Public Function ExecutarSQL(ByVal parSql As String) As Boolean
On Error GoTo Trata

Dim cmd As New ADODB.Command

If conn.State <> adStateOpen Then
    Call ConectarBD
    If conn.State <> adStateOpen Then
        ExecutarSQL = False
        Exit Function
    End If
End If

Set cmd.ActiveConnection = conn
cmd.CommandText = parSql
cmd.CommandType = adCmdText

cmd.Execute

Set cmd = Nothing
ExecutarSQL = True
Exit Function

Trata:
MsgBox "Erro", vbCritical
ExecutarSQL = False
If Not cmd Is Nothing Then Set cmd = Nothing
End Function

Public Function CarregarCombo(ByRef objComboBox As Object, ByVal sTabela As String, ByVal sColunaExibicao As String, Optional ByVal sColunaValor As String = "") As Boolean
On Error GoTo Trata

Dim rs As New ADODB.Recordset
Dim sSql As String


If conn.State <> adStateOpen Then
    Call ConectarBD
    If conn.State <> adStateOpen Then
        CarregarCombo = False
        Exit Function
    End If
End If

objComboBox.Clear

sSql = "SELECT " & sColunaExibicao & " FROM " & sTabela & " ORDER BY " & sColunaValor

rs.Open sSql, conn, adOpenKeyset, adLockOptimistic

Do While Not rs.EOF
    If sColunaValor <> "" Then
        objComboBox.AddItem rs.Fields(sColunaExibicao).Value
    Else
        objComboBox.AddItem rs.Fields(sColunaExibicao).Value
    End If
    rs.MoveNext
Loop

rs.Close
Set rs = Nothing
CarregarCombo = True
Exit Function
Trata:
If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
CarregarCombo = False
End Function

Public Function PegarCampo(ByVal parSql As String) As Variant

On Error GoTo Trata
Dim sDyna As New ADODB.Recordset
Dim valorRetorno As Variant

ConectarBD

sDyna.Open parSql, conn, adOpenKeyset, adLockReadOnly


If Not sDyna.EOF Then
    If Not IsNull(sDyna.Fields(0)) Then
        valorRetorno = sDyna.Fields(0).Value
    Else
        valorRetorno = Null
    End If
Else
    valorRetorno = Null
End If

sDyna.Close
Set sDyna = Nothing
PegarCampo = valorRetorno
Exit Function

Trata:
MsgBox "Ocorreu um erro na função PegarCampo: " & Err.Description, vbCritical
If Not sDyna Is Nothing Then If sDyna.State = adStateOpen Then sDyna.Close
Set sDyna = Nothing
PegarCampo = Null
End Function
