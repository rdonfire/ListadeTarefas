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
   Begin VB.CommandButton Btn 
      Caption         =   "Editar"
      Height          =   495
      Index           =   5
      Left            =   6600
      TabIndex        =   21
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Pesquisar"
      Height          =   495
      Index           =   4
      Left            =   6600
      TabIndex        =   20
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Sair"
      Height          =   495
      Index           =   3
      Left            =   12360
      TabIndex        =   10
      Top             =   360
      Width           =   1455
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
      MaxCols         =   7
      MaxRows         =   1
      RestrictCols    =   -1  'True
      RestrictRows    =   -1  'True
      SpreadDesigner  =   "Form1.frx":0000
      Appearance      =   2
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   2
      Left            =   10560
      TabIndex        =   12
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame QuaPesquisa 
      Caption         =   "Pesquisar"
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   13815
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   720
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   2760
         TabIndex        =   18
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         Format          =   66781185
         CurrentDate     =   45869
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Format          =   66781185
         CurrentDate     =   45869
      End
   End
   Begin VB.Frame fram_AddMeta 
      Caption         =   "Adicionar nova Meta"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   13935
      Begin MSComCtl2.DTPicker txtDataConclusao 
         Height          =   375
         Left            =   10080
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   66781185
         CurrentDate     =   45869
      End
      Begin VB.CheckBox chk_concluida 
         Caption         =   "Concluida"
         Height          =   255
         Left            =   9960
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker txtDataPrevista 
         Height          =   255
         Left            =   7200
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   66781185
         CurrentDate     =   45852
      End
      Begin VB.ComboBox cmbPrioridade 
         Height          =   315
         Left            =   7200
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtDescricao 
         Height          =   1575
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label lblDataConclusao 
         Caption         =   "Data de Conclusão"
         Height          =   255
         Left            =   10080
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Data Prevista"
         Height          =   255
         Left            =   7200
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Prioridade"
         Height          =   255
         Left            =   7320
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Confirmar"
      Height          =   495
      Index           =   1
      Left            =   8640
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame frameButton 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   13935
      Begin VB.CommandButton Btn 
         Caption         =   "Incluir"
         Height          =   495
         Index           =   0
         Left            =   8520
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
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
         TabIndex        =   13
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
Private Sub Btn_Click(Index As Integer)
On Error GoTo ErroAcoes
    
    Select Case Index
        Case 0
            ajusteButton Index
            
        Case 1
            If Not ValidaD Then Exit Sub
             If Not PassarDados Then Exit Sub
                MsgBox "Meta adicionada com Sucesso", vbInformation
                LimparCampos
        Case 2
            ajusteButton Index
            LimparCampos
        Case 3
            Unload Me
        Case 4
            AtualizarGrid GridPrincipal
        Case 5
            ajusteButton Index
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
cmbPrioridade.ListIndex = 0
End Sub

Private Sub chk_concluida_Click()
If chk_concluida.Value = 1 Then
    lblDataConclusao.Visible = True
    txtDataConclusao.Visible = True
    txtDataConclusao.Value = Date
Else
    lblDataConclusao.Visible = False
    txtDataConclusao.Visible = False
    txtDataConclusao.Value = 0
End If
End Sub

Private Sub Form_Load()
ajusteButton 1
End Sub
Private Sub ajusteButton(Optional ByVal parBt As Integer = 0)
If parBt <> 0 And parBt <> 5 Then
    Btn(0).Visible = True
    Btn(5).Visible = False
    Btn(2).Visible = False
    fram_AddMeta.Visible = False
    GridPrincipal.Visible = True
    txtDataConclusao.Visible = False
    lblDataConclusao.Visible = False
    txtDataPrevista.Value = Now
    QuaPesquisa.Visible = True
    Btn(1).Visible = False
    Exit Sub
End If
Btn(0).Visible = False
Btn(2).Visible = True
Btn(5).Visible = True
If parBt = 5 Then
    Btn(1).Visible = True
Else
    Btn(1).Visible = False
End If
Btn(4).Visible = False
fram_AddMeta.Visible = True
GridPrincipal.Visible = False
QuaPesquisa.Visible = False


If Not CarregarCombo(cmbPrioridade, "Prioridades", "DescricaoPrioridade", "IDPrioridade") Then
    MsgBox "Não foi possível carregar as prioridades.", vbExclamation
End If
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

Const MAX_LENGTH As Long = 255
If Len(txtDescricao.Text) > MAX_LENGTH Then
    MsgBox "A descrição não pode ter mais de " & MAX_LENGTH & " caracteres.", vbExclamation
    txtDescricao.Text = Left(txtDescricao.Text, MAX_LENGTH)
    txtDescricao.SetFocus
    Exit Function
End If

ValidaD = True
End Function
Private Function PassarDados()
On Error GoTo Trata
Dim sClsMeta As New clsMeta

With sClsMeta
    If Not .Adicionar(txtDescricao, txtDataPrevista, cmbPrioridade, chk_concluida, txtDataConclusao) Then
        MsgBox "Não foi possível adicionar a meta. Verifique os dados.", vbExclamation, "Erro ao Salvar"
        PassarDados = False
        Exit Function
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

sSql = "SELECT ID, DESCRICAO, PRIORIDADE, DATAVENCIMENTO, CONCLUIDA, DATACONCLUIDA FROM METAS ORDER BY DATAVENCIMENTO DESC"

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

    objSpread.SetText 1, iRow, "Consultar"

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
AtualizarGrid = False    '
End Function

Private Sub GridPrincipal_ButtonClick()

  MsgBox "Button!", vbInformation
End Sub

