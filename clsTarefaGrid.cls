VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarefaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public Descricao As String
Public Prioridade As String
Public DataVencimento As Date
Public Concluida As Boolean

Public Sub CarregarDaLinhaDoGrid(ByVal pSpread As Object, ByVal pRow As Long)
    On Error GoTo ErroCarregar
    
    Dim rsGrid As ADODB.Recordset
    
    Set rsGrid = pSpread.DataSource
    
    Me.ID = rsGrid.Fields("ID").Value
    Me.Descricao = rsGrid.Fields("Descricao").Value
    Me.Prioridade = rsGrid.Fields("Prioridade").Value
    Me.DataVencimento = rsGrid.Fields("DataVencimento").Value
    Me.Concluida = rsGrid.Fields("Concluida").Value
    
    Exit Sub
    
ErroCarregar:
    MsgBox "Erro ao carregar dados da grade para a classe: " & Err.Description, vbCritical
End Sub

