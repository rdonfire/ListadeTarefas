VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda de tarefas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread fpTarefas 
      Height          =   3255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   8415
      _Version        =   458752
      _ExtentX        =   14843
      _ExtentY        =   5741
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   5
      SpreadDesigner  =   "Form1.frx":0000
      UserResize      =   2
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Sair"
      Height          =   495
      Index           =   3
      Left            =   7080
      TabIndex        =   9
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Excluir"
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   8
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Editar"
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Btn 
      Caption         =   "Adicionar"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox txtDescricao 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adicionar nova tarefa"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.CheckBox chk_concluida 
         Caption         =   "Concluida"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpDataPrevista 
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   168755201
         CurrentDate     =   45852
      End
      Begin VB.ComboBox cmbPrioridade 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Data Prevista"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Prioridade"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
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

Dim Descricao As String
Dim Prioridade As String
Dim dataPrevista As Date
Dim sql As String
Dim concluidaStatus As Integer
Dim taskId As Long

Select Case Index
    Case 0
        If Trim(txtDescricao.Text) = "" Then
            MsgBox "A descrição da tarefa não pode estar vazia.", vbExclamation
            Exit Sub
        End If

        Descricao = txtDescricao.Text
        Prioridade = cmbPrioridade.Text
        dataPrevista = dtpDataPrevista.Value
        concluidaStatus = chk_concluida.Value
        concluidaStatus = IIf(chk_concluida.Value = vbChecked, 1, 0)

        If IsNumeric(Me.Tag) Then
            taskId = CLng(Me.Tag)
        Else
            taskId = 0
        End If

        If taskId > 0 Then    ' Se taskId for maior que 0, estamos no modo de edição
            sql = "UPDATE Tarefas SET Descricao = '" & Descricao & "', " & _
                    "Prioridade = '" & Prioridade & "', " & _
                    "DataVencimento = CONVERT(DATETIME, '" & Format(dataPrevista, "yyyy-mm-dd hh:mm:ss") & "', 120), " & _
                    "Concluida = " & concluidaStatus & " WHERE ID = " & taskId

            conn.Execute sql
            MsgBox "Tarefa atualizada com sucesso!", vbInformation
        Else    ' Se taskId for 0, estamos no modo de adição
            sql = "INSERT INTO Tarefas (Descricao, Prioridade, DataVencimento, Concluida) VALUES ('" & _
                    Descricao & "', '" & Prioridade & "', CONVERT(DATETIME, '" & Format(dataPrevista, "yyyy-mm-dd hh:mm:ss") & "', 120), " & concluidaStatus & ")"

            conn.Execute sql

            MsgBox "Tarefa adicionada com sucesso!", vbInformation
        End If

        LimparCampos
        PreencherGrid
    Case 1
        If fpTarefas.ActiveRow <= 0 Or fpTarefas.ActiveRow > fpTarefas.MaxRows - 1 Then
            MsgBox "Selecione uma tarefa na grade para editar.", vbExclamation
            Exit Sub
        End If

        Dim oTarefa As New clsTarefaGrid

        oTarefa.CarregarDaLinhaDoGrid fpTarefas, fpTarefas.ActiveRow

        txtDescricao.Text = oTarefa.Descricao
        cmbPrioridade.Text = oTarefa.Prioridade
        dtpDataPrevista.Value = oTarefa.DataVencimento
        chk_concluida.Value = IIf(oTarefa.Concluida, vbChecked, vbUnchecked)

        Me.Tag = oTarefa.ID
        Btn(0).Caption = "Salvar Edição"

    Case 2
        If fpTarefas.ActiveRow <= 0 Or fpTarefas.ActiveRow > fpTarefas.MaxRows - 1 Then
            MsgBox "Selecione uma tarefa na grade para excluir.", vbExclamation
            Exit Sub
        End If

        Dim oTarefaExcluir As New clsTarefaGrid
        oTarefaExcluir.CarregarDaLinhaDoGrid fpTarefas, fpTarefas.ActiveRow
        taskId = oTarefaExcluir.ID

        If MsgBox("Tem certeza que deseja excluir a tarefa '" & oTarefaExcluir.Descricao & "' (ID: " & taskId & ")?", vbYesNo + vbQuestion, "Confirmar Exclusão") = vbYes Then
            sql = "DELETE FROM Tarefas WHERE ID = " & taskId
            conn.Execute sql
            MsgBox "Tarefa excluída com sucesso!", vbInformation

            Call LimparCampos
            PreencherGrid
        Else
            MsgBox "Exclusão cancelada.", vbInformation
        End If
    Case 3
        Unload Me

    Case Else
        MsgBox "Ação não reconhecida.", vbExclamation
End Select

Exit Sub

ErroAcoes:
MsgBox "Ocorreu um erro na ação do botão: " & Err.Description, vbCritical
End Sub
Private Sub LimparCampos()
txtDescricao.Text = ""
cmbPrioridade.ListIndex = 0
dtpDataPrevista.Value = Now
chk_concluida.Value = vbUnchecked
Me.Tag = ""
Btn(0).Caption = "Adicionar"
End Sub
Private Sub Form_Load()
With cmbPrioridade
    .Clear
    .AddItem "Baixa"
    .AddItem "Média"
    .AddItem "Alta"
    .ListIndex = 0
End With
ConectarBD
LimparCampos
PreencherGrid
dtpDataPrevista.Value = Now
End Sub
Private Sub PreencherGrid()
    On Error GoTo ErroPreencherGrade
    
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    If conn.State = adStateClosed Then
        Call ConectarBD
    End If
    
    Set fpTarefas.DataSource = Nothing
    
    sql = "SELECT ID, Descricao, Prioridade, DataVencimento, Concluida FROM Tarefas ORDER BY DataVencimento"
    
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    Set fpTarefas.DataSource = rs
    
    fpTarefas.MaxRows = rs.RecordCount + 1
    'rs.Close
    
    Exit Sub
    
ErroPreencherGrade:
    MsgBox "Erro ao preencher a grade: " & Err.Description, vbCritical
End Sub
