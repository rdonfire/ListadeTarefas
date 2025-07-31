Attribute VB_Name = "Grid"
Option Explicit

Public Function AtualizarGrid(ByRef objSpread As fpSpread, ByVal sSql As String) As Boolean
On Error GoTo Trata
Dim rs As New ADODB.Recordset
Dim iRow As Long

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
