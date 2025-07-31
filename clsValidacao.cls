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
Private fConcluida As Byte
Private fDataConcluida As Date
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
Public Property Get Concluida() As Integer
Concluida = fConcluida
End Property
Public Property Let Concluida(ByVal parConcluida As Integer)
fConcluida = parConcluida
End Property
Public Property Get DataConcluida() As Date
DataConcluida = fDataConcluida
End Property
Public Property Let DataConcluida(ByVal parDataConcluida As Date)
fDataConcluida = parDataConcluida
End Property

Public Function Consultar() As Boolean
On Error GoTo Trata
Dim sSql As String, sDyna As New ADODB.Recordset

With sDyna
    sSql = "SELECT ID, DESCRICAO, DATAVENCIMENTO, PRIORIDADE, CONCLUIDA FROM METAS"

End With

Trata:

End Function

Public Function Adicionar(ByVal parDescricao As String, ByVal parData As Date, ByVal parPrioridade As String, Optional ByVal parConcluida As Byte = 0, Optional ByVal parDataConclusao As Date) As Boolean
Dim sSql As String

sSql = "INSERT INTO METAS (DESCRICAO, DATAVENCIMENTO, PRIORIDADE, CONCLUIDA, DATACONCLUIDA) VALUES ('" & parDescricao & "','" & parData & "','" & parPrioridade & "','" & parConcluida & "','" & parDataConclusao & "')"

If Not ExecutarSQL(sSql) Then
    Adicionar = False
Else
    Adicionar = True
End If
End Function

Public Function atualizar() As Boolean

End Function

