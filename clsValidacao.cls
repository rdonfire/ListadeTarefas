VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private fDescricao As String
Private fData As Date
Private fPrioridade As String
Private fConcluida As Boolean
Private fDataConcluida As Date
Private fMetaID As Integer
Private fDataCadastro As Date

Public Property Get Descricao() As String
Descricao = fDescricao
End Property
Public Property Let Descricao(ByVal parDescricao As String)
fDescricao = parDescricao
End Property
Public Property Get Data() As Date
Data = fData
End Property
Public Property Let Data(ByVal parData As Date)
fData = parData
End Property
Public Property Get Prioridade() As String
Prioridade = fPrioridade
End Property
Public Property Let Prioridade(ByVal parPrioridade As String)
fPrioridade = parPrioridade
End Property
Public Property Get Concluida() As Boolean
Concluida = fConcluida
End Property
Public Property Let Concluida(ByVal parConcluida As Boolean)
fConcluida = parConcluida
End Property
Public Property Get DataConcluida() As Date
DataConcluida = fDataConcluida
End Property
Public Property Let DataConcluida(ByVal parDataConcluida As Date)
fDataConcluida = parDataConcluida
End Property
Public Property Get MetaID() As Integer
MetaID = fMetaID
End Property
Public Property Let MetaID(ByVal parMetaID As Integer)
fMetaID = parMetaID
End Property
Public Property Get DataCadastro() As Date
DataCadastro = fDataCadastro
End Property
Public Property Let DataCadastro(ByVal parDataCadastro As Date)
fDataCadastro = parDataCadastro
End Property
Public Function Consulta(ByVal parCodigo As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String, sDyna As New ADODB.Recordset
Dim bResultado As Boolean
If ConectarBD Then Exit Function

sSql = "SELECT ID, DESCRICAO, DATAVENCIMENTO, PRIORIDADE, CONCLUIDA FROM METAS WHERE ID = " & parCodigo

With sDyna
    .Open sSql, conn, adOpenForwardOnly, adLockReadOnly
    If Not .EOF Then
        If IsNull(!ID) Then
            MsgBox "Meta não localizada!", vbExclamation
        Else
            fMetaID = !ID
            fDescricao = !Descricao
            fPrioridade = !Prioridade
            fConcluida = !Concluida
            fData = !Datavencimento
        End If
        bResultado = True
    End If
    .Close
End With
Consulta = bResultado
Set sDyna = Nothing
Exit Function

Trata:
MsgBox "ERRO NA CONSULTAR: " & Err.Description, "clsMeta.Consulta", vbInformation
Consulta = False
If Not sDyna Is Nothing Then
    If sDyna.State = adStateOpen Then
        sDyna.Close
    End If
    Set sDyna = Nothing
End If
End Function
Public Function Adicionar() As Boolean
Dim sSql As String
If Not ValidaMetaID(MetaID) Then
    sSql = "INSERT INTO METAS (DESCRICAO, DATAVENCIMENTO, PRIORIDADE, DATACADASTRO, CONCLUIDA, DATACONCLUIDA) VALUES ('" & fDescricao & "','" & fData & "','" & fPrioridade & "','" & fDataCadastro & "','" & fConcluida & "','" & fDataConcluida & "')"
Else
    sSql = "UPDATE METAS SET DESCRICAO = '" & Descricao & "'," & _
            "DATAVENCIMENTO = '" & Data & "'," & _
            "PRIORIDADE = '" & Prioridade & "'," & _
            "CONCLUIDA = '" & Concluida & "'," & _
            "DATACONCLUIDA = '" & DataConcluida & "'" & _
            "WHERE ID = " & MetaID
End If
If ExecutarSQL(sSql) Then Adicionar = True
End Function
Public Function ValidaMetaID(ByVal parCodigo As Integer) As Boolean
On Error GoTo Trata
Dim sSql As String, sDyna As New ADODB.Recordset
Dim bResultado As Boolean
If ConectarBD Then Exit Function

sSql = "SELECT ID, DESCRICAO, DATAVENCIMENTO, PRIORIDADE, CONCLUIDA FROM METAS WHERE ID = " & parCodigo

With sDyna
    .Open sSql, conn, adOpenForwardOnly, adLockReadOnly
    If Not .EOF Then
        If IsNull(!ID) Then
            bResultado = False
        End If
        bResultado = True
    End If
    .Close
End With
ValidaMetaID = bResultado
Set sDyna = Nothing
Exit Function

Trata:
MsgBox "ERRO NA CONSULTAR: " & Err.Description, "clsMeta.Consulta", vbInformation
ValidaMetaID = False
If Not sDyna Is Nothing Then
    If sDyna.State = adStateOpen Then
        sDyna.Close
    End If
    Set sDyna = Nothing
End If
End Function
Public Function ExcluirMeta(ByVal parCodigo As Integer) As Boolean
Dim sSql As String

sSql = "DELETE METAS WHERE ID = " & parCodigo
If ExecutarSQL(sSql) Then ExcluirMeta = True
End Function

