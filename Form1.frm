VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda de tarefas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn 
      Caption         =   "Incluir"
      Height          =   855
      Index           =   0
      Left            =   9240
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton btn 
      Caption         =   "Sair"
      Height          =   855
      Index           =   3
      Left            =   12840
      Picture         =   "Form1.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin FPSpreadADO.fpSpread GridPrincipal 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   13935
      _Version        =   458752
      _ExtentX        =   24580
      _ExtentY        =   5953
      _StockProps     =   64
      BackColorStyle  =   3
      DAutoCellTypes  =   0   'False
      DAutoFill       =   0   'False
      DAutoSave       =   0   'False
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   -2147483633
      GridColor       =   -2147483633
      MaxCols         =   8
      MaxRows         =   1
      RestrictCols    =   -1  'True
      RestrictRows    =   -1  'True
      SelectBlockOptions=   9
      ShadowColor     =   -2147483626
      ShadowDark      =   -2147483638
      SpreadDesigner  =   "Form1.frx":2194
      Appearance      =   2
   End
   Begin VB.CommandButton btn 
      Caption         =   "Cancelar"
      Height          =   855
      Index           =   2
      Left            =   11640
      Picture         =   "Form1.frx":262D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame QuaPesquisa 
      Caption         =   "Pesquisar"
      Height          =   2655
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   13815
      Begin VB.CommandButton btn 
         Caption         =   "Pesquisar"
         Height          =   855
         Index           =   4
         Left            =   12000
         Picture         =   "Form1.frx":36F7
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cmbPesquisa 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   720
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131923969
         CurrentDate     =   45869
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131923969
         CurrentDate     =   45869
      End
   End
   Begin VB.Frame frAddMeta 
      Caption         =   "Adicionar nova Meta"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   14175
      Begin VB.TextBox txtMetaID 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         Height          =   375
         HideSelection   =   0   'False
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtDataConclusao 
         Height          =   375
         Left            =   10080
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131923969
         CurrentDate     =   45869
      End
      Begin VB.CheckBox chk_concluida 
         Caption         =   "Concluida"
         Height          =   255
         Left            =   10080
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker txtDataPrevista 
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131923969
         CurrentDate     =   45852
      End
      Begin VB.ComboBox cmbPrioridade 
         Height          =   315
         Left            =   7200
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtDescricao 
         Height          =   1695
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label lblMetaID 
         Caption         =   "META ID"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDataConclusao 
         Caption         =   "Data de Conclusão"
         Height          =   255
         Left            =   10080
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Data Prevista"
         Height          =   255
         Left            =   7200
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Prioridade"
         Height          =   255
         Left            =   7200
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton btn 
      Caption         =   "Excluir"
      Height          =   855
      Index           =   6
      Left            =   11640
      Picture         =   "Form1.frx":47C1
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton btn 
      Caption         =   "Editar"
      Height          =   855
      Index           =   5
      Left            =   10440
      Picture         =   "Form1.frx":588B
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton btn 
      Caption         =   "Confirmar"
      Height          =   855
      Index           =   1
      Left            =   10440
      Picture         =   "Form1.frx":6955
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame frameButton 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   13935
      Begin VB.Label lbl_Cabecalho 
         Alignment       =   2  'Center
         Caption         =   "AGENDA DE METAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   6615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fContAlteracaoDosCampos As Integer
Dim fDataCadastro As Date
Dim sCodigo As String
Dim fHabilitada As Boolean
Const MAX_LENGTH As Long = 255

Enum CheckConcluida
    eNumHabilitada = 1
    eNumDesatbilida = 0
End Enum

Enum eNumAtualizar
    eAtualizar = 0
    eNaoAtualizar = 1
End Enum

Enum eNumAcao
    eIncluir = 0
    eConfirmar = 1
    eCancelar = 2
    eSair = 3
    ePesquisar = 4
    eEditar = 5
    eExcluir = 6
End Enum
Private Sub Btn_Click(Index As Integer)
On Error GoTo ErroAcoes

Select Case Index
    Case eNumAcao.eIncluir
        ajusteButton Index
        txtDescricao.SetFocus
    Case eNumAcao.eConfirmar
        If Not ValidaD Then Exit Sub
        GerarID
        PassarDados eNumAtualizar.eAtualizar
        LimparCampos
        AtualizarGrid GridPrincipal
        ajusteButton Index
    Case eNumAcao.eCancelar
        ajusteButton Index
        AtualizarGrid GridPrincipal
        LimparCampos
    Case eNumAcao.eSair
        Dim sResultado As VbMsgBoxResult
        sResultado = MsgBox("Deseja Sair?", vbYesNo)
        If sResultado = vbYes Then Unload Me
    Case eNumAcao.ePesquisar
        AtualizarGrid GridPrincipal
    Case eNumAcao.eEditar
        If Not SelecionarMeta Then Exit Sub
        ajusteButton Index
        txtDescricao.SetFocus
    Case eNumAcao.eExcluir
        ExcluirMeta
    Case Else
End Select
Exit Sub
ErroAcoes:
MsgBox "Ocorreu um erro na ação do botão: " & Err.Description, vbCritical
End Sub
Private Sub LimparCampos()
txtDescricao.Text = ""
txtDataPrevista.Value = Now
chk_concluida.Value = vbUnchecked
txtMetaID = ""
End Sub
Private Sub chk_concluida_Click()
If chk_concluida.Value = eNumHabilitada Then
    lblDataConclusao.Visible = True
    txtDataConclusao.Visible = True
    txtDataConclusao.Value = Date
    fHabilitada = True
Else
    lblDataConclusao.Visible = False
    txtDataConclusao.Visible = False
    txtDataConclusao.Value = 0
    fHabilitada = False
End If
End Sub
Private Sub Form_Load()
ajusteButton 1
'cmbPesquisa.AddItem
End Sub
Private Sub ajusteButton(Optional ByVal parBt As Integer = 0)
If parBt <> 0 And parBt <> 5 Then
    btn(0).Visible = True
    btn(2).Visible = False
    frAddMeta.Visible = False
    GridPrincipal.Visible = True
    txtDataConclusao.Visible = False
    lblDataConclusao.Visible = False
    txtDataPrevista.Value = Now
    QuaPesquisa.Visible = True
    btn(1).Visible = False
    btn(4).Visible = True
    btn(5).Visible = True
    btn(6).Visible = True
    Exit Sub
End If
btn(0).Visible = False
btn(2).Visible = True
btn(5).Visible = False
btn(6).Visible = False
If parBt = 5 Then
    btn(1).Visible = True
Else
    btn(1).Visible = False
    LimparCampos
    If Not CarregarCombo(cmbPrioridade, "Prioridades", "DescricaoPrioridade", "IDPrioridade") Then
        MsgBox "Não foi possível carregar as prioridades.", vbExclamation
    End If
End If
btn(4).Visible = False
frAddMeta.Visible = True
GridPrincipal.Visible = False
QuaPesquisa.Visible = False
btn(1).Visible = True
End Sub
Private Function ValidaD() As Boolean

If Trim(txtDescricao.Text) = "" Then
    MsgBox "A descrição da tarefa não pode estar vazia.", vbExclamation
    txtDescricao.SetFocus
    Exit Function
End If

If Trim(cmbPrioridade.Text) = "" Then
    MsgBox "Selecione um nivel de prioridade.", vbExclamation
    cmbPrioridade.SetFocus
    Exit Function
End If

If txtDataPrevista.Value = "" Then
    MsgBox "A data prevista não pode estar vazia.", vbExclamation
    txtDataPrevista.SetFocus
    Exit Function
End If

If Len(txtDescricao.Text) > MAX_LENGTH Then
    MsgBox "A descrição não pode ter mais de " & MAX_LENGTH & " caracteres.", vbExclamation
    txtDescricao.Text = Left(txtDescricao.Text, MAX_LENGTH)
    txtDescricao.SetFocus
    Exit Function
End If

ValidaD = True
End Function
Private Function PassarDados(Optional ByVal parCondicao As Integer)
On Error GoTo Trata
Dim sclsMeta As New clsMeta

fDataCadastro = Now()

With sclsMeta

    .Descricao = txtDescricao
    .Data = txtDataPrevista
    .Prioridade = cmbPrioridade
    .DataCadastro = fDataCadastro
    .Concluida = chk_concluida
    .DataConcluida = txtDataConclusao
    .MetaID = txtMetaID

    If parCondicao = eNumAtualizar.eAtualizar Then
        If Not .Adicionar() Then
            MsgBox "Não foi possível adicionar a meta. Verifique os dados.", vbExclamation, "Erro ao Salvar"
            PassarDados = False
            Exit Function
        End If
    Else
        If Not .ExcluirMeta(sCodigo) Then
            MsgBox "Não foi possivel realizar a exclusão da meta selecionada!", vbInformation
            PassarDados = False
            Exit Function
        End If
    End If
End With

PassarDados = True
Exit Function
Trata:
MsgBox "Ocorreu um erro inesperado ao processar os dados.", vbCritical, "Erro de Sistema"
PassarDados = False
End Function
Public Function AtualizarGrid(ByRef objSpread As fpSpread) As Boolean
On Error GoTo Trata
Dim rs As New ADODB.Recordset
Dim iRow As Long
Dim sSql As String

sSql = "SELECT ID, DESCRICAO, PRIORIDADE, DATAVENCIMENTO, CONCLUIDA, DATACONCLUIDA FROM METAS ORDER BY DATACADASTRO DESC"

If conn.State <> adStateOpen Then
    Call ConectarBD
    If conn.State <> adStateOpen Then
        AtualizarGrid = False
        Exit Function
    End If
End If

objSpread.BlockMode = True
objSpread.MaxRows = 0

rs.Open sSql, conn, adOpenKeyset, adLockOptimistic

iRow = 0
Do While Not rs.EOF
    iRow = iRow + 1
    objSpread.MaxRows = iRow

    If Not IsNull(rs.Fields("ID").Value) Then
        objSpread.SetText 1, iRow, rs.Fields("ID").Value
    Else
        objSpread.SetText 1, iRow, ""
    End If
    If Not IsNull(rs.Fields("DESCRICAO").Value) Then
        objSpread.SetText 2, iRow, rs.Fields("DESCRICAO").Value
    Else
        objSpread.SetText 2, iRow, ""
    End If

    If Not IsNull(rs.Fields("PRIORIDADE").Value) Then
        objSpread.SetText 3, iRow, rs.Fields("PRIORIDADE").Value
    Else
        objSpread.SetText 3, iRow, ""
    End If

    If Not IsNull(rs.Fields("DATAVENCIMENTO").Value) Then
        objSpread.SetText 4, iRow, Format(rs.Fields("DATAVENCIMENTO").Value, "dd/mm/yyyy")
    Else
        objSpread.SetText 4, iRow, ""
    End If

    If Not IsNull(rs.Fields("CONCLUIDA").Value) Then
        If CBool(rs.Fields("CONCLUIDA").Value) = True Then
            objSpread.SetText 5, iRow, "CONCLUÍDA"
        Else
            objSpread.SetText 5, iRow, ""
        End If
    Else
        objSpread.SetText 5, iRow, ""
    End If

    If Not IsNull(rs.Fields("DATACONCLUIDA").Value) Then
        objSpread.SetText 6, iRow, Format(rs.Fields("DATACONCLUIDA").Value, "dd/mm/yyyy")
    Else
        objSpread.SetText 6, iRow, ""
    End If

    objSpread.SetText 7, iRow, "Aguardando Cáculo"
    rs.MoveNext
Loop

rs.Close
Set rs = Nothing

objSpread.BlockMode = False
objSpread.Refresh

AtualizarGrid = True
Exit Function

Trata:
MsgBox "Erro ao Atualizar Grid: " & Err.Description, vbCritical
If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
objSpread.BlockMode = False
AtualizarGrid = False
End Function
Private Function RecebeDados(ByVal parCodigo As String) As Boolean
On Error GoTo Tratar
Dim sclsMeta As New clsMeta
If Not parCodigo = "" Then
    With sclsMeta
        If .Consulta(parCodigo) Then
            txtDescricao = .Descricao
            cmbPrioridade = .Prioridade
            txtDataPrevista = .Data
            txtMetaID = .MetaID
            habilitaCheckConcluida (.Concluida)
            RecebeDados = True
            Exit Function
        End If
    End With
Else
    RecebeDados = False
    Exit Function
End If
Tratar:
MsgBox "Erro ao coletar dados! ", vbCritical
RecebeDados = False
End Function
Private Sub habilitaCheckConcluida(ByVal parCheck As Boolean)
If parCheck Then
    fHabilitada = True
    chk_concluida = eNumHabilitada
End If
End Sub
Private Function VerificaSeAlterou(ByVal parCodigo As String) As Boolean
Dim sclsMeta As New clsMeta
fContAlteracaoDosCampos = 0

With sclsMeta
    If .Consulta(parCodigo) Then
        If txtDescricao <> .Descricao Then
            fContAlteracaoDosCampos = 1
            Exit Function
        End If
        If cmbPrioridade <> .Prioridade Then
            fContAlteracaoDosCampos = 1
            Exit Function
        End If
        If txtDataPrevista <> .Data Then
            fContAlteracaoDosCampos = 1
            Exit Function
        End If
        If fHabilitada <> .Concluida Then
            fContAlteracaoDosCampos = 1
            Exit Function
        End If
    End If
End With
End Function
Public Sub GerarID()
Dim sUltimoID As Integer

sUltimoID = PegarCampo("SELECT TOP 1 ID FROM METAS ORDER BY DATACADASTRO DESC")

If txtMetaID = "" Then
    txtMetaID = sUltimoID + "0001"
End If
Exit Sub
End Sub
Private Sub GridPrincipal_click(ByVal Row As Long, ByVal Col As Long)
Dim sclsMeta As clsMeta
MarcaLinha GridPrincipal, Row, 1, Col, sCodigo
txtMetaID = sCodigo
If Row = 8 Then ExcluirMeta
End Sub
Private Sub ExcluirMeta()
Dim sResultado As VbMsgBoxResult
If Not SelecionarMeta Then Exit Sub
sResultado = MsgBox("Deseja excluir a meta " & sCodigo & "?", vbYesNo)
If sResultado = vbYes Then
    If PassarDados(2) Then MsgBox "Meta excluída com sucesso!", vbInformation
    AtualizarGrid GridPrincipal
End If
End Sub
Private Function SelecionarMeta() As Boolean
SelecionarMeta = True
If Not RecebeDados(txtMetaID) Then
    MsgBox "Selecione uma META!", vbExclamation
    SelecionarMeta = False
    Exit Function
End If
End Function

